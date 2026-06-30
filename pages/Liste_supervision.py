import streamlit as st
import pandas as pd
import io
import zipfile
import re
from xlsxwriter.utility import xl_col_to_name
from streamlit_gsheets import GSheetsConnection

# Configuration de la page
st.set_page_config(page_title="Tri Supervision ODICEE", layout="wide")
st.title("Tri de la liste de supervision")

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1pkj6frncXmzUUVAClAWp63HY_UpQUahOCLz19w-UseI/edit?usp=sharing"
conn = st.connection("gsheets", type=GSheetsConnection)

@st.cache_data(ttl=30)
def load_liste_initiales():
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="Liste_initiales")
        if df.empty: return {}
        df_clean = df.copy().dropna(subset=["Initiales"])
        return dict(zip(df_clean['Initiales'].astype(str).str.strip(), df_clean['Couleur'].astype(str).str.strip()))
    except Exception:
        return {}

dict_initiales = load_liste_initiales()

def lire_colonnes_numeriques_brutes(file_bytes, noms_colonnes_attendus):
    # Fonction de lecture XML (inchangée)
    resultats = {}
    try:
        with zipfile.ZipFile(file_bytes) as z:
            sheet_path = 'xl/worksheets/sheet1.xml'
            if sheet_path not in z.namelist(): return resultats
            with z.open('xl/workbook.xml') as f: pass 
            with z.open(sheet_path) as f: content = f.read().decode('utf-8')

            header_row_match = re.search(r'<row r="1"[^>]*>(.*?)</row>', content, re.DOTALL)
            if not header_row_match: return resultats
            header_xml = header_row_match.group(1)

            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f: ss_content = f.read().decode('utf-8')
                for si_match in re.finditer(r'<si>(.*?)</si>', ss_content, re.DOTALL):
                    texts = re.findall(r'<t[^>]*>([^<]*)</t>', si_match.group(1))
                    shared_strings.append(''.join(texts))

            col_letter_to_name = {}
            for cell_match in re.finditer(r'<c r="([A-Z]+)1"[^>]*?(?:\st="(\w+)")?[^>]*>(.*?)</c>', header_xml):
                col_letter, cell_type, cell_inner = cell_match.groups()
                v_match = re.search(r'<v>([^<]*)</v>', cell_inner or '')
                if not v_match: continue
                raw_val = v_match.group(1)
                if cell_type == 's':
                    try: nom = shared_strings[int(raw_val)]
                    except (ValueError, IndexError): nom = raw_val
                else: nom = raw_val
                col_letter_to_name[col_letter] = nom.strip()

            lettres_cibles = [l for l, nom in col_letter_to_name.items() if nom in noms_colonnes_attendus]

            for lettre in lettres_cibles:
                nom_col = col_letter_to_name[lettre]
                valeurs = {}
                pattern = re.compile(r'<c r="' + lettre + r'(\d+)"[^>]*>(.*?)</c>')
                for m in pattern.finditer(content):
                    row_num = int(m.group(1))
                    if row_num == 1: continue 
                    cell_xml = m.group(2)
                    v_match = re.search(r'<v>([^<]*)</v>', cell_xml)
                    if v_match:
                        try: valeurs[row_num] = float(v_match.group(1))
                        except ValueError: valeurs[row_num] = None
                    else: valeurs[row_num] = None
                resultats[nom_col] = valeurs
    except Exception: return {}
    return resultats

# Zone de dépôt des fichiers
col1, col2 = st.columns(2)
with col1:
    file_odicee = st.file_uploader(
        "1. Importer la supervision (ODICEE)",
        type=["xlsx", "csv"],
        help="ODICEE / Pilotage / Dossiers au stade 3F.xlsx"
    )
with col2:
    file_controle = st.file_uploader(
        "2. Importer les paramètres (Export_Controle)",
        type=["xlsx", "csv"],
        help="ODICEE / Accueil / A Controler / Export_Controle_lots_de_travaux.xlsx"
    )

st.subheader("Paramètres de tri", help=(
    "**1 - URGENCE** : Dossiers dont le passage en stade 3F remonte à plus d'un mois — triés du plus ancien au plus récent.\n\n"
    "**2 - PRIORITÉ** : Dossiers dont la date de réalisation réelle est antérieure ou égale à la date de fin sélectionnée, ou appartenant à un lot de contrôle 3F dont la première date de réalisation est antérieure ou égale à cette date — les dossiers d'un même lot sont regroupés ensemble et triés selon la date de réalisation la plus ancienne du lot, du plus ancien au plus récent.\n\n"
    "**3 - Classique** : Dossiers ne répondant à aucun critère d'urgence ou de priorité — triés par date de réalisation réelle du plus ancien au plus récent."
))
st.markdown("**Conditions :** Définir la date limite de réalisation réelle (toutes les dates antérieures sont prises en compte)", help="L'intervalle doit prendre en compte la date de réalisation des dossiers pour le prochain dépôt")
date_fin = st.date_input("Date de fin", value=None, format="DD/MM/YYYY")
date_debut = None


if file_odicee is not None and file_controle is not None:
    if st.button("🚀 Générer le fichier trié"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            status_text.info("⏳ Étape 1 : Lecture des fichiers...")
            df = pd.read_csv(file_odicee) if file_odicee.name.endswith('.csv') else pd.read_excel(file_odicee)
            df_ctrl = pd.read_csv(file_controle) if file_controle.name.endswith('.csv') else pd.read_excel(file_controle)
            
            df.columns = df.columns.str.strip()
            df_ctrl.columns = df_ctrl.columns.str.strip()
            df['Ordre_Import'] = range(len(df))
            progress_bar.progress(25)

            status_text.info("⏳ Étape 2 : Vérification des colonnes...")
            # NOMS EXACTS ISSUS DE TON FICHIER D'EXEMPLE
            col_stade_3f = 'Date stade 3F'
            col_rea = 'Date de réalisation réelle'
            col_volume_classique = 'Volume classique (kWhc)'
            col_volume_precarite = 'Volume précarité (kWhc)'
            col_odi_id = 'Numéro du dossier'
            
            col_ctr_id = 'N°dossier'
            col_ctr_date = 'Date réelle de réalisation' # <--- CORRIGÉ ICI
            col_ctr_lot = 'Lot de contrôle'
            col_ctr_stade = 'Stade'

            # VÉRIFICATION
            manque_odi = [c for c in [col_stade_3f, col_rea, col_odi_id] if c not in df.columns]
            manque_ctr = [c for c in [col_ctr_id, col_ctr_date, col_ctr_lot, col_ctr_stade] if c not in df_ctrl.columns]
            
            if manque_odi or manque_ctr:
                status_text.error(f"❌ Arrêt. Colonnes introuvables :\nODICEE : {manque_odi}\nContrôle : {manque_ctr}")
                st.stop()
            progress_bar.progress(50)

            status_text.info("⏳ Étape 3 : Traitement des données...")
            if not file_odicee.name.endswith('.csv'):
                file_odicee.seek(0)
                valeurs_brutes = lire_colonnes_numeriques_brutes(file_odicee, {col_volume_classique, col_volume_precarite})
                file_odicee.seek(0)
                for nom_col, vals in valeurs_brutes.items():
                    if nom_col in df.columns:
                        df[nom_col] = [vals.get(idx + 2) for idx in range(len(df))]

            for c in [col_volume_classique, col_volume_precarite]:
                if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce')

            # FORMATAGE DES DATES
            df[col_stade_3f] = pd.to_datetime(df[col_stade_3f], errors='coerce', dayfirst=True)
            df[col_rea] = pd.to_datetime(df[col_rea], errors='coerce', dayfirst=True)
            df_ctrl[col_ctr_date] = pd.to_datetime(df_ctrl[col_ctr_date], errors='coerce', dayfirst=True)

            # TRI & CONDITIONS
            df['Priorité_Tri'] = 99
            un_mois_avant = pd.Timestamp.today() - pd.DateOffset(months=1)
            
            # Condition 1
            mask_cond1 = df[col_stade_3f] < un_mois_avant
            df.loc[mask_cond1, 'Priorité_Tri'] = 1

            df['_date_ref_lot'] = pd.NaT  # initialisé ici, rempli si lot trouvé

            if date_fin:
                end_dt = pd.to_datetime(date_fin)
                
                # Condition 2 : toutes les dates de réalisation jusqu'à la date de fin
                mask_cond2 = (df[col_rea] <= end_dt) & (df['Priorité_Tri'] != 1) 
                df.loc[mask_cond2, 'Priorité_Tri'] = 2

                # Condition 2bis (Lots)
                df_ctrl_3f = df_ctrl[df_ctrl[col_ctr_stade].astype(str).str.contains('3F', case=False, na=False)].copy()
                if not df_ctrl_3f.empty:
                    dates_min_par_lot = df_ctrl_3f.groupby(col_ctr_lot)[col_ctr_date].min().reset_index()
                    lots_valides = dates_min_par_lot[
                        dates_min_par_lot[col_ctr_date] <= end_dt
                    ][col_ctr_lot]
                    
                    dossiers_valides = df_ctrl_3f[df_ctrl_3f[col_ctr_lot].isin(lots_valides)][col_ctr_id].astype(str).str.strip()
                    df['Temp_ID'] = df[col_odi_id].astype(str).str.strip()
                    mask_cond2bis = df['Temp_ID'].isin(dossiers_valides) & (df['Priorité_Tri'] != 1)
                    df.loc[mask_cond2bis, 'Priorité_Tri'] = 2

                    # Date de référence du lot : date min du lot pour trier les dossiers groupés ensemble
                    df_ctrl_3f_valides = df_ctrl_3f[df_ctrl_3f[col_ctr_lot].isin(lots_valides)][[col_ctr_id, col_ctr_lot]].copy()
                    df_ctrl_3f_valides[col_ctr_id] = df_ctrl_3f_valides[col_ctr_id].astype(str).str.strip()
                    dates_min_par_lot_valides = dates_min_par_lot[dates_min_par_lot[col_ctr_lot].isin(lots_valides)].rename(columns={col_ctr_date: '_date_ref_lot'})
                    df_ctrl_3f_valides = df_ctrl_3f_valides.merge(dates_min_par_lot_valides[[col_ctr_lot, '_date_ref_lot']], on=col_ctr_lot, how='left')
                    map_id_to_date_lot = dict(zip(df_ctrl_3f_valides[col_ctr_id], df_ctrl_3f_valides['_date_ref_lot']))
                    df['_date_ref_lot'] = df['Temp_ID'].map(map_id_to_date_lot)

                    df.drop(columns=['Temp_ID'], inplace=True)
            progress_bar.progress(75)

            status_text.info("⏳ Étape 4 : Tri final...")

            # Clé de tri secondaire pour le groupe 2 :
            # - dossiers via lot → date min du lot (tous regroupés ensemble)
            # - dossiers via date individuelle → leur propre date de réalisation
            df['_sort_rea'] = df['_date_ref_lot'].fillna(df[col_rea])
            df['_sort_stade_3f'] = df[col_stade_3f]

            df_sorted = df.sort_values(
                by=['Priorité_Tri', '_sort_stade_3f', '_sort_rea', 'Ordre_Import'],
                ascending=[True, True, True, True]
            ).drop(columns=['_sort_stade_3f', '_sort_rea', '_date_ref_lot'])
            
            df_sorted['Bandeau Priorité'] = df_sorted['Priorité_Tri'].map({
                1: '1 - URGENCE (> 1 mois Stade 3F)',
                2: '2 - PRIORITÉ (Date de réalisation ou Lot)',
                99: '3 - Classique'
            })
            
            cols = df_sorted.columns.tolist()
            cols.insert(0, cols.pop(cols.index('Bandeau Priorité')))
            df_sorted = df_sorted[cols].drop(columns=['Priorité_Tri', 'Ordre_Import'])

            df_sorted[col_stade_3f] = df_sorted[col_stade_3f].dt.strftime('%d/%m/%Y')
            df_sorted[col_rea] = df_sorted[col_rea].dt.strftime('%d/%m/%Y')

            st.session_state['df_resultat'] = df_sorted
            progress_bar.progress(100)
            status_text.success("✅ Tri terminé avec succès ! 👇")

        except Exception as e:
            status_text.empty()
            st.error(f"🚨 Une erreur critique a bloqué le script : {e}")

# Affichage et Export
if 'df_resultat' in st.session_state:
    df_result = st.session_state['df_resultat']
    st.write(f"**Aperçu des dossiers triés ({len(df_result)} lignes) :**")
    st.dataframe(df_result.head(20), use_container_width=True)

    # Insertion colonne Initiales en position 0 (même logique qu'Export_liste)
    df_export = df_result.copy()
    if 'Initiales' not in df_export.columns:
        df_export.insert(0, 'Initiales', '')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Supervision')
        workbook  = writer.book
        worksheet = writer.sheets['Supervision']
        max_row, max_col = df_export.shape

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max_row, max_col - 1)

        # Largeur des colonnes
        def safe_len(val):
            if pd.isna(val): return 0
            return len(str(val))
        for i, col in enumerate(df_export.columns):
            max_data_len = df_export[col].map(safe_len).max()
            if pd.isna(max_data_len): max_data_len = 0
            column_len = max(int(max_data_len), len(str(col))) + 2
            worksheet.set_column(i, i, min(int(column_len), 40))

        # Formatage conditionnel colonne Initiales (identique à fonctions.py)
        if dict_initiales and max_row > 0:
            dossier_col_letter = xl_col_to_name(df_export.columns.get_loc('Numéro du dossier')) if 'Numéro du dossier' in df_export.columns else None
            total_lignes = max_row + 1
            plage_dossier = f"${dossier_col_letter}$2:${dossier_col_letter}${total_lignes}" if dossier_col_letter else ""
            plage_init = f"$A$2:$A${total_lignes}"

            formats_init = {}
            for init, color in dict_initiales.items():
                color_clean = str(color).strip()
                if color_clean and color_clean.lower() != 'nan':
                    if not color_clean.startswith('#') and len(color_clean) == 6:
                        color_clean = f"#{color_clean}"
                    formats_init[init] = workbook.add_format({'bg_color': color_clean, 'font_color': '#000000'})

            for init, fmt_init in formats_init.items():
                if dossier_col_letter:
                    formula = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}2, {plage_init}, "{init}")>0'
                else:
                    formula = f'=$A2="{init}"'
                worksheet.conditional_format(1, 0, max_row, 2, {'type': 'formula', 'criteria': formula, 'format': fmt_init})

            # Couleur par défaut si initiales renseignées mais non reconnues
            fmt_en_cours = workbook.add_format({'bg_color': '#FFE699', 'font_color': '#595959'})
            if dossier_col_letter:
                formula_default = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}2, {plage_init}, "<>")>0'
            else:
                formula_default = '=$A2<>""'
            worksheet.conditional_format(1, 0, max_row, 2, {'type': 'formula', 'criteria': formula_default, 'format': fmt_en_cours})

    output.seek(0)
    st.download_button(
        label="📥 Télécharger l'Excel trié",
        data=output,
        file_name="Supervision_Triee.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )