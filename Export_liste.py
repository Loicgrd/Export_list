import streamlit as st
import pandas as pd
import io
import requests
import zipfile
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
from xlsxwriter.utility import xl_col_to_name  # <--- NOUVEL IMPORT NÉCESSAIRE

st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

# ==========================================
# CONNEXION AU GOOGLE SHEET
# ==========================================
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1pkj6frncXmzUUVAClAWp63HY_UpQUahOCLz19w-UseI/edit?usp=sharing"
conn = st.connection("gsheets", type=GSheetsConnection)

@st.cache_data(ttl=10)
def load_worksheet(sheet_name):
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name)
        if df.empty or len(df.columns) < 3:
            df = pd.DataFrame(columns=["Nom", "SIREN", "Date d'ajout"])
        if len(df.columns) < 3:
            df["Date d'ajout"] = ""
            
        df = df.dropna(subset=[df.columns[0]])
        df = df.iloc[:, :3]
        
        dict_siren = dict(zip(df.iloc[:, 0], df.iloc[:, 1].astype(str)))
        return dict_siren, df
    except Exception as e:
        st.error(f"Erreur de lecture de la feuille '{sheet_name}' : {e}")
        return {}, pd.DataFrame(columns=["Nom", "SIREN", "Date d'ajout"])

@st.cache_data(ttl=10)
def load_admin():
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="ADMIN")
        if df.empty:
            return [], []
            
        mots_docs = [str(x).strip() for x in df.get("Nom du document", pd.Series()).dropna() if str(x).strip() and str(x).strip().lower() != 'nan']
        mots_coms = [str(x).strip() for x in df.get("Commentaire", pd.Series()).dropna() if str(x).strip() and str(x).strip().lower() != 'nan']
        
        return mots_docs, mots_coms
    except Exception as e:
        st.error(f"Erreur de lecture de la feuille 'ADMIN' : {e}")
        return [], []

dict_confort, df_confort_gsheet = load_worksheet("Confort")
dict_cdc, df_cdc_gsheet = load_worksheet("CDC")
mots_docs_admin, mots_coms_admin = load_admin()

st.title("Générateur d'Exports CEE")

tab_generateur, tab_confort, tab_cdc, tab_admin = st.tabs([
    "📊 Générateur", "⚙️ Base Confort", "⚙️ Base CDC", "🛡️ Filtres ADMIN"
])

# ==========================================
# FONCTIONS UTILITAIRES
# ==========================================
def afficher_gestion_base(sheet_name, df_gsheet):
    st.header(f"Base de données '{sheet_name}'")
    state_recherche = f"recherche_{sheet_name}"
    state_trouves = f"trouves_{sheet_name}"
    if state_recherche not in st.session_state:
        st.session_state[state_recherche] = False
        st.session_state[state_trouves] = []
    
    st.subheader("➕ Ajouter plusieurs bailleurs (par SIREN)")
    liste_sirens_brut = st.text_area(f"Collez vos SIREN ici :", key=f"input_siren_{sheet_name}")
    
    if st.button("🔍 Rechercher", key=f"btn_search_{sheet_name}"):
        sirens = [s.strip() for s in liste_sirens_brut.replace('\n', ',').split(',') if s.strip()]
        trouves = []
        with st.spinner("Recherche..."):
            for s in sirens:
                resp = requests.get(f"https://recherche-entreprises.api.gouv.fr/search?q={s}")
                if resp.status_code == 200 and resp.json().get("results"):
                    res = resp.json()["results"][0]
                    trouves.append({"Nom": res.get('sigle') or res.get('nom_raison_sociale'), "SIREN": res.get('siren')})
        st.session_state[state_trouves] = trouves
        st.session_state[state_recherche] = True

    if st.session_state[state_recherche]:
        if st.session_state[state_trouves]:
            st.success(f"✅ {len(st.session_state[state_trouves])} trouvés !")
            st.dataframe(pd.DataFrame(st.session_state[state_trouves]), hide_index=True)
            if st.button("✅ Confirmer l'ajout", key=f"btn_conf_{sheet_name}", type="primary"):
                date_jour = datetime.now().strftime("%d/%m/%Y")
                nom_col, siren_col, date_col = df_gsheet.columns[0], df_gsheet.columns[1], df_gsheet.columns[2]
                nouveaux = [{nom_col: b['Nom'], siren_col: b['SIREN'], date_col: date_jour} for b in st.session_state[state_trouves]]
                df_updated = pd.concat([df_gsheet, pd.DataFrame(nouveaux)], ignore_index=True)
                conn.update(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, data=df_updated)
                st.session_state[state_recherche] = False
                st.cache_data.clear()
                st.rerun()
    
    st.divider()
    st.subheader("📋 Liste actuelle")
    if not df_gsheet.empty:
        df_display = df_gsheet[[df_gsheet.columns[2], df_gsheet.columns[0], df_gsheet.columns[1]]]
        st.dataframe(df_display, use_container_width=True, hide_index=True)

    if st.button("🗑️ Supprimer la sélection", type="primary", key=f"btn_del_{sheet_name}"):
        st.rerun()

def sauvegarder_admin(l_docs, l_coms):
    max_len = max(len(l_docs), len(l_coms))
    df_new = pd.DataFrame({
        "Nom du document": l_docs + [""] * (max_len - len(l_docs)),
        "Commentaire": l_coms + [""] * (max_len - len(l_coms))
    })
    conn.update(spreadsheet=SPREADSHEET_URL, worksheet="ADMIN", data=df_new)
    st.cache_data.clear()
    st.rerun()

with tab_confort: afficher_gestion_base("Confort", df_confort_gsheet)
with tab_cdc: afficher_gestion_base("CDC", df_cdc_gsheet)
with tab_admin:
    st.header("🛠️ Mots-clés pour le tri automatique ADMIN")
    col_d, col_c = st.columns(2)
    with col_d:
        st.subheader("📄 Colonne 'Nom du document'")
        nouveau_doc = st.text_input("Ajouter un mot-clé (ex: Visa) :")
        if st.button("➕ Ajouter", key="add_doc") and nouveau_doc:
            if nouveau_doc not in mots_docs_admin: sauvegarder_admin(mots_docs_admin + [nouveau_doc], mots_coms_admin)
        if mots_docs_admin:
            a_suppr_doc = st.multiselect("Supprimer :", mots_docs_admin, key="suppr_doc")
            if st.button("🗑️ Enlever", key="btn_suppr_doc") and a_suppr_doc:
                sauvegarder_admin([m for m in mots_docs_admin if m not in a_suppr_doc], mots_coms_admin)
            st.dataframe(pd.DataFrame(mots_docs_admin, columns=["Mots-clés (Documents)"]), hide_index=True, use_container_width=True)

    with col_c:
        st.subheader("💬 Colonne 'Commentaire'")
        nouveau_com = st.text_input("Ajouter un mot-clé (ex: Abandon) :")
        if st.button("➕ Ajouter", key="add_com") and nouveau_com:
            if nouveau_com not in mots_coms_admin: sauvegarder_admin(mots_docs_admin, mots_coms_admin + [nouveau_com])
        if mots_coms_admin:
            a_suppr_com = st.multiselect("Supprimer :", mots_coms_admin, key="suppr_com")
            if st.button("🗑️ Enlever", key="btn_suppr_com") and a_suppr_com:
                sauvegarder_admin(mots_docs_admin, [m for m in mots_coms_admin if m not in a_suppr_com])
            st.dataframe(pd.DataFrame(mots_coms_admin, columns=["Mots-clés (Commentaires)"]), hide_index=True, use_container_width=True)

# ==========================================
# GÉNÉRATEURS EXCEL (Formules croisées magiques incluses)
# ==========================================
def generer_excel_formate(df, nom_feuille):
    if not df.empty and 'Pris par (Initiales)' not in df.columns:
        df.insert(0, 'Pris par (Initiales)', '')

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
        df.to_excel(writer, index=False, sheet_name=nom_feuille)
        worksheet = writer.sheets[nom_feuille]
        (max_row, max_col) = df.shape

        if max_col > 0:
            worksheet.autofilter(0, 0, max_row, max_col - 1)
            worksheet.freeze_panes(1, 0)
            
            # Application de la coloration de groupe
            dossier_col_letter = None
            if 'Numéro dossier' in df.columns:
                dossier_col_letter = xl_col_to_name(df.columns.get_loc('Numéro dossier'))
                
            fmt_en_cours = writer.book.add_format({'bg_color': '#FFE699', 'font_color': '#595959'})
            
            if dossier_col_letter:
                # Si n'importe quelle ligne avec le même dossier a des initiales, on colorie
                formula = f'=COUNTIFS(${dossier_col_letter}:${dossier_col_letter}, ${dossier_col_letter}2, $A:$A, "<>")>0'
            else:
                formula = '=$A2<>""'
                
            worksheet.conditional_format(1, 0, max_row, max_col - 1, {'type': 'formula', 'criteria': formula, 'format': fmt_en_cours})

            for i in range(max_col): worksheet.set_column(i, i, 16) 
    return buffer.getvalue()

def generer_excel_dcr(df_prio, df_classique, nom_feuille):
    if not df_prio.empty and 'Pris par (Initiales)' not in df_prio.columns:
        df_prio.insert(0, 'Pris par (Initiales)', '')
    if not df_classique.empty and 'Pris par (Initiales)' not in df_classique.columns:
        df_classique.insert(0, 'Pris par (Initiales)', '')

    # Identification de la colonne "Numéro dossier"
    dossier_col_letter = None
    if not df_classique.empty and 'Numéro dossier' in df_classique.columns:
        dossier_col_letter = xl_col_to_name(df_classique.columns.get_loc('Numéro dossier'))
    elif not df_prio.empty and 'Numéro dossier' in df_prio.columns:
        dossier_col_letter = xl_col_to_name(df_prio.columns.get_loc('Numéro dossier'))

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(nom_feuille)
        
        fmt_rouge = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'red', 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
        fmt_bleu = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#3a75c4', 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
        fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2'})
        fmt_en_cours = workbook.add_format({'bg_color': '#FFE699', 'font_color': '#595959'}) # Orange clair
        
        current_row = 0
        max_col = len(df_classique.columns) if not df_classique.empty else (len(df_prio.columns) if not df_prio.empty else 1)
        
        # --- 1. LISTE PRIORITAIRE ---
        if not df_prio.empty:
            worksheet.merge_range(current_row, 0, current_row, max_col - 1, "↓ /!\\ Liste prioritaire (à faire avant de passer sur une DCR) /!\\ ↓", fmt_rouge)
            current_row += 1
            for i, val in enumerate(df_prio.columns.values): worksheet.write(current_row, i, val, fmt_header)
            current_row += 1
            
            start_prio = current_row
            df_prio.to_excel(writer, sheet_name=nom_feuille, startrow=current_row, header=False, index=False)
            end_prio = current_row + len(df_prio) - 1
            
            # Formule magique
            if dossier_col_letter:
                formule_prio = f'=COUNTIFS(${dossier_col_letter}:${dossier_col_letter}, ${dossier_col_letter}{start_prio + 1}, $A:$A, "<>")>0'
            else: formule_prio = f'=$A{start_prio + 1}<>""'
            worksheet.conditional_format(start_prio, 0, end_prio, max_col - 1, {'type': 'formula', 'criteria': formule_prio, 'format': fmt_en_cours})
            
            current_row += len(df_prio)
            worksheet.merge_range(current_row, 0, current_row, max_col - 1, "↑ /!\\ Liste prioritaire (à faire avant de passer sur une DCR) /!\\ ↑", fmt_rouge)
            current_row += 1
            
        # --- 2. LISTE CLASSIQUE ---
        if not df_classique.empty or df_prio.empty:
            worksheet.merge_range(current_row, 0, current_row, max_col - 1, "Puis basculer sur DCR (pensez à vérifier votre planning).", fmt_bleu)
            current_row += 1
            ligne_entete = current_row
            for i, val in enumerate(df_classique.columns.values): worksheet.write(current_row, i, val, fmt_header)
            current_row += 1
            
            if not df_classique.empty:
                start_classique = current_row
                df_classique.to_excel(writer, sheet_name=nom_feuille, startrow=current_row, header=False, index=False)
                end_classique = current_row + len(df_classique) - 1
                
                # Formule magique
                if dossier_col_letter:
                    formule_classique = f'=COUNTIFS(${dossier_col_letter}:${dossier_col_letter}, ${dossier_col_letter}{start_classique + 1}, $A:$A, "<>")>0'
                else: formule_classique = f'=$A{start_classique + 1}<>""'
                worksheet.conditional_format(start_classique, 0, end_classique, max_col - 1, {'type': 'formula', 'criteria': formule_classique, 'format': fmt_en_cours})
                
                worksheet.autofilter(ligne_entete, 0, end_classique, max_col - 1)
                
        for i in range(max_col): worksheet.set_column(i, i, 16)
        
    return buffer.getvalue()


# ==========================================
# ONGLET 1 : GÉNÉRATEUR PRINCIPAL
# ==========================================
with tab_generateur:
    uploaded_file = st.file_uploader("Importer le fichier Excel (Liste globale)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df_source = pd.read_excel(uploaded_file)
            st.success(f"Fichier chargé ! ({len(df_source)} lignes)")

            for col in df_source.columns:
                if 'date' in str(col).lower() or 'période' in str(col).lower():
                    df_source[col] = pd.to_datetime(df_source[col], errors='coerce')

            # --- ANALYSE ET DÉFINITION DE LA PRIORITÉ ---
            st.subheader("📅 Analyse et définition de la priorité (DCR)")
            
            if st.toggle("📊 Afficher le tableau de synthèse"):
                if 'Date réception' in df_source.columns and 'DCR' in df_source.columns:
                    df_synth = df_source.dropna(subset=['Date réception']).copy()
                    df_synth['Date réception'] = pd.to_datetime(df_synth['Date réception']).dt.date
                    df_synth['DCR'] = df_synth['DCR'].fillna('Non renseigné')
                    
                    tableau = pd.crosstab(index=df_synth['Date réception'], columns=df_synth['DCR'], margins=True, margins_name='Total')
                    tableau_dates = tableau.drop('Total').sort_index(ascending=True)
                    tableau_final = pd.concat([tableau_dates, tableau.loc[['Total']]])
                    
                    index_formate = []
                    for idx in tableau_final.index:
                        if isinstance(idx, str): index_formate.append(idx)
                        else: index_formate.append(idx.strftime('%d/%m/%Y'))
                    tableau_final.index = index_formate

                    dates_sans_total = [d for d in tableau_final.index if d != 'Total']
                    top_5_dates = dates_sans_total[-5:]

                    def coloriser_delais(row):
                        styles = []
                        for col_name in row.index:
                            if row.name == 'Total' or col_name == 'Total':
                                styles.append('background-color: #e6e6e6; font-weight: bold; color: black')
                            else:
                                if row.name in top_5_dates: styles.append('background-color: #d4edda; color: #155724')
                                else: styles.append('background-color: #f8d7da; color: #721c24')
                        return styles

                    st.dataframe(tableau_final.style.apply(coloriser_delais, axis=1), use_container_width=True)
                else:
                    st.warning("⚠️ Les colonnes 'Date réception' ou 'DCR' sont manquantes pour générer le tableau.")
            
            date_prio = st.date_input("Dossiers reçus JUSQU'À cette date (incluse) = Prioritaires :", value=None, format="DD/MM/YYYY")

            # --- CONFORT & CDC ---
            df_confort, df_cdc = pd.DataFrame(), pd.DataFrame()
            if 'Bénéficiaire' in df_source.columns:
                df_confort = df_source[df_source['Bénéficiaire'].isin(dict_confort.keys())].copy()
                if not df_confort.empty: df_confort.insert(0, 'SIREN', df_confort['Bénéficiaire'].map(dict_confort))
                
                df_cdc = df_source[df_source['Bénéficiaire'].isin(dict_cdc.keys())].copy()
                if not df_cdc.empty: df_cdc.insert(0, 'SIREN', df_cdc['Bénéficiaire'].map(dict_cdc))

            # --- LISTE À EXPORTER (Base) ---
            df_export = df_source.copy()
            if 'Numéro dossier' in df_export.columns:
                if not df_confort.empty: df_export = df_export[~df_export['Numéro dossier'].isin(df_confort['Numéro dossier'])]
                if not df_cdc.empty: df_export = df_export[~df_export['Numéro dossier'].isin(df_cdc['Numéro dossier'])]

            # --- ADMIN ---
            df_admin = pd.DataFrame()
            mask_admin = pd.Series(False, index=df_export.index)
            if 'Nom du document' in df_export.columns and mots_docs_admin:
                for mot in mots_docs_admin: mask_admin = mask_admin | df_export['Nom du document'].astype(str).str.contains(mot, case=False, na=False, regex=False)
            if 'Commentaire' in df_export.columns and mots_coms_admin:
                for mot in mots_coms_admin: mask_admin = mask_admin | df_export['Commentaire'].astype(str).str.contains(mot, case=False, na=False, regex=False)
            if mask_admin.any():
                df_admin = df_export[mask_admin].copy()
                df_export = df_export[~mask_admin]

            # --- TRI GLOBAL DES BASES (Astuce pour tout regrouper) ---
            if 'Numéro dossier' in df_export.columns: df_export = df_export.sort_values(by=['Numéro dossier', 'Date réception'], na_position='last')
            if 'Numéro dossier' in df_confort.columns: df_confort = df_confort.sort_values(by=['Numéro dossier', 'Date réception'], na_position='last')
            if 'Numéro dossier' in df_cdc.columns: df_cdc = df_cdc.sort_values(by=['Numéro dossier', 'Date réception'], na_position='last')
            if 'Numéro dossier' in df_admin.columns: df_admin = df_admin.sort_values(by=['Numéro dossier', 'Date réception'], na_position='last')

            # --- SÉPARATION DCR (Prioritaire / Classique) ---
            df_prio, df_classique = pd.DataFrame(), df_export.copy()
            if 'Date réception' in df_export.columns and date_prio is not None:
                dates_reception = pd.to_datetime(df_export['Date réception']).dt.date
                mask_prio = dates_reception <= date_prio
                df_prio = df_export[mask_prio].copy()
                df_classique = df_export[~mask_prio].copy()

            # --- CRÉATION FICHIERS ET ZIP ---
            st.divider()
            date_export = datetime.now().strftime("%d-%m-%Y")
            fichiers_a_zipper = {}

            nom_dcr = f"ODICEE-{date_export}-DCR_export_doc_com_non_vus.xlsx"
            fichiers_a_zipper[nom_dcr] = generer_excel_dcr(df_prio, df_classique, 'Liste à exporter')
            
            if not df_confort.empty:
                nom_confort = f"ODICEE-{date_export}-CONFORT_export_doc_com_non_vus.xlsx"
                fichiers_a_zipper[nom_confort] = generer_excel_formate(df_confort, 'Confort')
                
            if not df_cdc.empty:
                nom_cdc = f"ODICEE-{date_export}-CDC_export_doc_com_non_vus.xlsx"
                fichiers_a_zipper[nom_cdc] = generer_excel_formate(df_cdc, 'CDC')
                
            if not df_admin.empty:
                nom_admin = f"ODICEE-{date_export}-ADMIN_export_doc_com_non_vus.xlsx"
                fichiers_a_zipper[nom_admin] = generer_excel_formate(df_admin, 'ADMIN')

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for n, d in fichiers_a_zipper.items(): zf.writestr(n, d)
            
            # --- AFFICHAGE BOUTONS ---
            st.download_button("📦 TÉLÉCHARGER TOUS LES EXPORTS (.zip)", zip_buffer.getvalue(), f"ODICEE-{date_export}-TOUS_LES_EXPORTS.zip", use_container_width=True, type="primary")
            st.markdown("<p style='text-align: center; color: gray;'>Ou télécharger individuellement :</p>", unsafe_allow_html=True)
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.subheader("📊 DCR")
                st.text(f"{len(df_prio) + len(df_classique)} lignes.")
                st.download_button("📥 Télécharger DCR", fichiers_a_zipper.get(nom_dcr, b''), nom_dcr, use_container_width=True)
            with c2:
                st.subheader("🏢 Confort")
                st.text(f"{len(df_confort)} lignes.")
                if not df_confort.empty: st.download_button("📥 Télécharger", fichiers_a_zipper.get(nom_confort, b''), nom_confort, use_container_width=True)
            with c3:
                st.subheader("🏛️ CDC")
                st.text(f"{len(df_cdc)} lignes.")
                if not df_cdc.empty: st.download_button("📥 Télécharger", fichiers_a_zipper.get(nom_cdc, b''), nom_cdc, use_container_width=True)
            with c4:
                st.subheader("🛡️ ADMIN")
                st.text(f"{len(df_admin)} lignes.")
                if not df_admin.empty: st.download_button("📥 Télécharger", fichiers_a_zipper.get(nom_admin, b''), nom_admin, use_container_width=True)

        except Exception as e:
            st.error(f"Erreur lors de la génération de l'Excel : {e}")
