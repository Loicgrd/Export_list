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

@st.cache_data(ttl=10)
def load_worksheet_bailleurs(sheet_name):
    """Charge une base de bailleurs (Confort ou CDC) depuis Google Sheets.
    Retourne un dict {Raison sociale: SIREN} — la clé sert à comparer avec
    'Raison sociale du bénéficiaire' pour router chaque dossier vers la bonne feuille."""
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name)
        if df.empty or len(df.columns) < 2:
            return {}
        df = df.dropna(subset=[df.columns[0]])
        return dict(zip(df.iloc[:, 0], df.iloc[:, 1].astype(str)))
    except Exception:
        return {}

dict_confort = load_worksheet_bailleurs("Confort")
dict_national = load_worksheet_bailleurs("CDC")

def ajouter_feuille_avec_bandeau(writer, df_prio, df_classique, nom_feuille, dict_initiales=None, dossier_col_nom='Numéro du dossier'):
    """Écrit une feuille Excel avec un bandeau rouge encadrant les dossiers prioritaires
    (Priorité 1 - URGENCE et 2 - PRIORITÉ), suivis des dossiers classiques (Priorité 3).
    Contrairement à fonctions.ajouter_feuille_dcr, le bandeau est TOUJOURS affiché même sans
    aucun dossier prioritaire (avec une ligne vide pour ajout manuel), et il n'y a pas de bandeau
    secondaire "Puis basculer sur DCR" — ce texte n'a de sens que pour la feuille DCR d'origine
    de l'app Export_liste, pas pour un usage générique multi-feuilles."""
    if dict_initiales is None:
        dict_initiales = {}
    if df_prio.empty and df_classique.empty:
        return

    colonnes = df_classique.columns.tolist() if not df_classique.empty else df_prio.columns.tolist()

    if not df_prio.empty and 'Initiales' not in df_prio.columns:
        df_prio.insert(0, 'Initiales', '')
    if not df_classique.empty and 'Initiales' not in df_classique.columns:
        df_classique.insert(0, 'Initiales', '')

    # 'Bandeau Priorité' toujours en dernière colonne
    colonnes_finales = [c for c in colonnes if c != 'Bandeau Priorité']
    if 'Bandeau Priorité' in colonnes:
        colonnes_finales.append('Bandeau Priorité')
    if not df_prio.empty:
        df_prio = df_prio[colonnes_finales]
    if not df_classique.empty:
        df_classique = df_classique[colonnes_finales]

    dossier_col_letter = None
    if dossier_col_nom in colonnes_finales:
        dossier_col_letter = xl_col_to_name(colonnes_finales.index(dossier_col_nom))

    workbook = writer.book
    worksheet = workbook.add_worksheet(nom_feuille)
    worksheet.freeze_panes(2, 0)

    fmt_rouge = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'red', 'align': 'center', 'valign': 'vcenter'})
    fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2'})
    fmt_en_cours = workbook.add_format({'bg_color': '#FFE699', 'font_color': '#595959'})

    current_row = 0
    max_col = len(colonnes_finales)

    nb_lignes_prio = len(df_prio) if not df_prio.empty else 1  # 1 ligne vide si pas de prioritaires
    first_data_row_excel = 3
    last_data_row_excel = nb_lignes_prio + len(df_classique) + 5
    plage_dossier = f"${dossier_col_letter}${first_data_row_excel}:${dossier_col_letter}${last_data_row_excel}" if dossier_col_letter else ""
    plage_init = f"$A${first_data_row_excel}:$A${last_data_row_excel}"

    formats_init = {}
    for init, color in dict_initiales.items():
        color_clean = str(color).strip()
        if color_clean and color_clean.lower() != 'nan':
            if not color_clean.startswith('#') and len(color_clean) == 6:
                color_clean = f"#{color_clean}"
            formats_init[init] = workbook.add_format({'bg_color': color_clean, 'font_color': '#000000'})

    # --- Bandeau prioritaire : toujours affiché, même vide ---
    worksheet.merge_range(current_row, 0, current_row, max_col - 1, "↓ /!\\ Liste prioritaire /!\\ ↓", fmt_rouge)
    current_row += 1
    for i, val in enumerate(colonnes_finales):
        worksheet.write(current_row, i, val, fmt_header)
    current_row += 1
    start_prio = current_row

    if not df_prio.empty:
        df_prio.to_excel(writer, sheet_name=nom_feuille, startrow=current_row, header=False, index=False)
        end_prio = current_row + len(df_prio) - 1
        current_row += len(df_prio)
    else:
        # Ligne vide pour ajout manuel si aucun dossier prioritaire
        end_prio = current_row
        current_row += 1

    for init, fmt_init in formats_init.items():
        if dossier_col_letter:
            f_prio_init = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}{start_prio + 1}, {plage_init}, "{init}")>0'
        else:
            f_prio_init = f'=$A{start_prio + 1}="{init}"'
        worksheet.conditional_format(start_prio, 0, end_prio, 2, {'type': 'formula', 'criteria': f_prio_init, 'format': fmt_init})

    if dossier_col_letter:
        formule_prio = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}{start_prio + 1}, {plage_init}, "<>")>0'
    else:
        formule_prio = f'=$A{start_prio + 1}<>""'
    worksheet.conditional_format(start_prio, 0, end_prio, 2, {'type': 'formula', 'criteria': formule_prio, 'format': fmt_en_cours})

    worksheet.merge_range(current_row, 0, current_row, max_col - 1, "↑ /!\\ Liste prioritaire /!\\ ↑", fmt_rouge)
    current_row += 1

    # --- Dossiers classiques (Priorité 3) ---
    ligne_entete_classique = current_row
    for i, val in enumerate(colonnes_finales):
        worksheet.write(current_row, i, val, fmt_header)
    current_row += 1

    if not df_classique.empty:
        start_classique = current_row
        df_classique.to_excel(writer, sheet_name=nom_feuille, startrow=current_row, header=False, index=False)
        end_classique = current_row + len(df_classique) - 1

        for init, fmt_init in formats_init.items():
            if dossier_col_letter:
                f_classique_init = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}{start_classique + 1}, {plage_init}, "{init}")>0'
            else:
                f_classique_init = f'=$A{start_classique + 1}="{init}"'
            worksheet.conditional_format(start_classique, 0, end_classique, 2, {'type': 'formula', 'criteria': f_classique_init, 'format': fmt_init})

        if dossier_col_letter:
            formule_classique = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}{start_classique + 1}, {plage_init}, "<>")>0'
        else:
            formule_classique = f'=$A{start_classique + 1}<>""'
        worksheet.conditional_format(start_classique, 0, end_classique, 2, {'type': 'formula', 'criteria': formule_classique, 'format': fmt_en_cours})
        worksheet.autofilter(ligne_entete_classique, 0, end_classique, max_col - 1)

    for i in range(max_col):
        worksheet.set_column(i, i, 16)


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
    "**1 - URGENCE** : Dossiers dont le passage en stade 3F remonte à plus de 25 jours — triés du plus ancien au plus récent.\n\n"
    "**2 - PRIORITÉ** : Dossiers dont la date de réalisation réelle est antérieure ou égale à la date de fin sélectionnée, ou appartenant à un lot de contrôle 3F dont la première date de réalisation est antérieure ou égale à cette date — les dossiers d'un même lot sont regroupés ensemble et triés selon la date de réalisation la plus ancienne du lot, du plus ancien au plus récent.\n\n"
    "**3 - Classique** : Dossiers ne répondant à aucun critère d'urgence ou de priorité — triés par date stade 3F du plus ancien au plus récent (même ordre que l'export ODICEE stade 3F d'origine)."
))
st.markdown("**Conditions :** Définir la date limite de réalisation réelle (toutes les dates antérieures sont prises en compte)", help="L'intervalle doit prendre en compte la date de réalisation des dossiers pour le prochain dépôt")
date_fin_defaut = pd.Timestamp.today() - pd.DateOffset(months=11)
date_fin = st.date_input("Date de fin", value=date_fin_defaut.date(), format="DD/MM/YYYY")
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
            seuil_urgence = pd.Timestamp.today() - pd.Timedelta(days=25)
            
            # Condition 1
            mask_cond1 = df[col_stade_3f] < seuil_urgence
            df.loc[mask_cond1, 'Priorité_Tri'] = 1

            df['_date_ref_lot'] = pd.NaT  # initialisé ici, rempli si lot trouvé

            if date_fin:
                end_dt = pd.to_datetime(date_fin)
                
                # Condition 2 : toutes les dates de réalisation jusqu'à la date de fin
                mask_cond2 = (df[col_rea] <= end_dt) & (df['Priorité_Tri'] != 1) 
                df.loc[mask_cond2, 'Priorité_Tri'] = 2

                # Condition 2bis (Lots)
                # "NEANT" dans la colonne Lot de contrôle signifie qu'il n'y a PAS de vrai lot —
                # ce n'est pas un identifiant de regroupement. Sans cette exclusion, tous les
                # dossiers marqués NEANT seraient groupés ensemble par erreur (le groupby les
                # traite comme un lot unique), leur donnant tous la même date de tri minimale
                # alors qu'ils n'ont aucun rapport entre eux.
                df_ctrl_3f = df_ctrl[
                    df_ctrl[col_ctr_stade].astype(str).str.contains('3F', case=False, na=False)
                    & (df_ctrl[col_ctr_lot].astype(str).str.strip().str.upper() != 'NEANT')
                ].copy()
                if not df_ctrl_3f.empty:
                    dates_min_par_lot = df_ctrl_3f.groupby(col_ctr_lot)[col_ctr_date].min().reset_index()
                    lots_valides = dates_min_par_lot[
                        dates_min_par_lot[col_ctr_date] <= end_dt
                    ][col_ctr_lot]
                    
                    dossiers_valides = df_ctrl_3f[df_ctrl_3f[col_ctr_lot].isin(lots_valides)][col_ctr_id].astype(str).str.strip()
                    df['Temp_ID'] = df[col_odi_id].astype(str).str.strip()
                    mask_cond2bis = df['Temp_ID'].isin(dossiers_valides) & (df['Priorité_Tri'] != 1)
                    df.loc[mask_cond2bis, 'Priorité_Tri'] = 2

                    # Date de référence du lot : date min du lot pour trier les dossiers groupés ensemble.
                    # Un dossier peut apparaître dans PLUSIEURS lots de contrôle différents (une fiche/lot
                    # de travaux par lot). Avant cette correction, dict(zip(...)) retenait arbitrairement
                    # le dernier lot rencontré dans l'ordre du fichier — pas nécessairement le plus ancien.
                    # On prend maintenant systématiquement la date la plus ancienne parmi tous ses lots.
                    df_ctrl_3f_valides = df_ctrl_3f[df_ctrl_3f[col_ctr_lot].isin(lots_valides)][[col_ctr_id, col_ctr_lot]].copy()
                    df_ctrl_3f_valides[col_ctr_id] = df_ctrl_3f_valides[col_ctr_id].astype(str).str.strip()
                    dates_min_par_lot_valides = dates_min_par_lot[dates_min_par_lot[col_ctr_lot].isin(lots_valides)].rename(columns={col_ctr_date: '_date_ref_lot'})
                    df_ctrl_3f_valides = df_ctrl_3f_valides.merge(dates_min_par_lot_valides[[col_ctr_lot, '_date_ref_lot']], on=col_ctr_lot, how='left')
                    map_id_to_date_lot = df_ctrl_3f_valides.groupby(col_ctr_id)['_date_ref_lot'].min().to_dict()
                    df['_date_ref_lot'] = df['Temp_ID'].map(map_id_to_date_lot)

                    df.drop(columns=['Temp_ID'], inplace=True)
            progress_bar.progress(75)

            status_text.info("⏳ Étape 4 : Tri final...")

            # Clé de tri unique selon le groupe de priorité :
            # - Groupe 1 (URGENCE)   : trié par date stade 3F (plus ancien -> plus récent)
            # - Groupe 2 (PRIORITÉ)  : trié par date de réalisation (date min du lot si applicable,
            #                          sinon date de réalisation individuelle)
            # - Groupe 99 (Classique): trié par date stade 3F (même ordre que l'export ODICEE
            #                          stade 3F d'origine) — PAS par date de réalisation
            df['_sort_rea'] = df['_date_ref_lot'].fillna(df[col_rea])
            df['_sort_key'] = df['_sort_rea']
            df.loc[df['Priorité_Tri'] == 1, '_sort_key'] = df.loc[df['Priorité_Tri'] == 1, col_stade_3f]
            df.loc[df['Priorité_Tri'] == 99, '_sort_key'] = df.loc[df['Priorité_Tri'] == 99, col_stade_3f]

            df_sorted = df.sort_values(
                by=['Priorité_Tri', '_sort_key', 'Ordre_Import'],
                ascending=[True, True, True]
            ).drop(columns=['_sort_key', '_sort_rea', '_date_ref_lot'])
            
            df_sorted['Bandeau Priorité'] = df_sorted['Priorité_Tri'].map({
                1: '1 - URGENCE (> 25 jours Stade 3F)',
                2: '2 - PRIORITÉ (Date de réalisation ou Lot)',
                99: '3 - Classique'
            })
            
            cols = df_sorted.columns.tolist()
            cols.append(cols.pop(cols.index('Bandeau Priorité')))
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

    col_beneficiaire = 'Raison sociale du bénéficiaire'

    # ─────────────────────────────────────────────
    # Routage vers Confort / National / DCR selon la raison sociale du bénéficiaire
    # (même logique et même ordre de priorité que l'app Export_liste : Confort d'abord,
    # puis National sur le reste, le solde va sur DCR)
    # ─────────────────────────────────────────────
    if col_beneficiaire in df_result.columns:
        mask_confort = df_result[col_beneficiaire].isin(dict_confort.keys())
        df_confort = df_result[mask_confort].copy()
        df_reste = df_result[~mask_confort].copy()

        mask_national = df_reste[col_beneficiaire].isin(dict_national.keys())
        df_national = df_reste[mask_national].copy()
        df_dcr = df_reste[~mask_national].copy()
    else:
        st.warning(f"⚠️ Colonne '{col_beneficiaire}' introuvable — tous les dossiers restent sur la feuille DCR.")
        df_confort = pd.DataFrame(columns=df_result.columns)
        df_national = pd.DataFrame(columns=df_result.columns)
        df_dcr = df_result.copy()

    def split_prio_classique(df):
        """Sépare un DataFrame en (prioritaires [1+2], classiques [3]) selon 'Bandeau Priorité'."""
        if df.empty or 'Bandeau Priorité' not in df.columns:
            return df.iloc[0:0].copy(), df.copy()
        mask_prio = df['Bandeau Priorité'].str.startswith(('1', '2'))
        return df[mask_prio].copy(), df[~mask_prio].copy()

    dcr_prio, dcr_classique = split_prio_classique(df_dcr)
    national_prio, national_classique = split_prio_classique(df_national)
    confort_prio, confort_classique = split_prio_classique(df_confort)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        ajouter_feuille_avec_bandeau(writer, dcr_prio, dcr_classique, 'DCR', dict_initiales)
        ajouter_feuille_avec_bandeau(writer, national_prio, national_classique, 'National', dict_initiales)
        ajouter_feuille_avec_bandeau(writer, confort_prio, confort_classique, 'Confort', dict_initiales)

    output.seek(0)
    nom_fichier = f"Liste_Supervision_{pd.Timestamp.today().strftime('%Y-%m-%d')}.xlsx"
    st.download_button(
        label="📥 Télécharger l'Excel trié",
        data=output,
        file_name=nom_fichier,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("📄 DCR", f"{len(df_dcr)} lignes")
    c2.metric("🇫🇷 National", f"{len(df_national)} lignes")
    c3.metric("🏢 Confort", f"{len(df_confort)} lignes")
