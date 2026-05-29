import streamlit as st
import pandas as pd
import io
import requests
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

# ==========================================
# CONNEXION AU GOOGLE SHEET
# ==========================================
# Remplace cette URL par celle de ton fichier exact si besoin
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1pkj6frncXmzUUVAClAWp63HY_UpQUahOCLz19w-UseI/edit?usp=sharing"

# Création de la connexion
conn = st.connection("gsheets", type=GSheetsConnection)

# Fonction pour lire les données (avec mise en cache pour ne pas surcharger Google)

@st.cache_data(ttl=10)
def load_bailleurs():
    try:
        # Ajout de worksheet="Confort" ici
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="Confort", usecols=[0, 1])
        df = df.dropna(subset=[df.columns[0]])
        return dict(zip(df.iloc[:, 0], df.iloc[:, 1].astype(str))), df
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        return {}, pd.DataFrame()

# Chargement de la base de données GSheet
dict_siren, df_bailleurs_gsheet = load_bailleurs()

st.title("Générateur d'Exports CEE")

tab_generateur, tab_reglages = st.tabs(["📊 Générateur d'Exports", "⚙️ Base de données Bailleurs"])

# ==========================================
# ONGLET 2 : GESTION DANS GOOGLE SHEETS
# ==========================================
with tab_reglages:
    st.header("Base de données 'Confort'")
    st.markdown("Ces bailleurs sont sauvegardés dans votre Google Sheet.")
    
    col_search, col_list = st.columns([1, 1])
    
    with col_search:
        st.subheader("🔍 Ajouter à la base")
        query = st.text_input("Entrez le SIREN ou le nom de l'organisme :")
        
        if query:
            response = requests.get(f"https://recherche-entreprises.api.gouv.fr/search?q={query}")
            if response.status_code == 200:
                results = response.json().get("results", [])
                if results:
                    options = {f"{r.get('nom_complet')} (SIREN: {r.get('siren')})": r for r in results}
                    selected_label = st.selectbox("Sélectionnez l'organisme :", list(options.keys()))
                    selected_data = options[selected_label]
                    
                    siren_api = selected_data.get('siren')
                    nom_defaut = selected_data.get('sigle') or selected_data.get('nom_raison_sociale')
                    
                    nom_excel = st.text_input("Nom exact dans l'Excel :", value=nom_defaut)
                    
                    if st.button("➕ Enregistrer dans le Google Sheet", type="primary"):
                        if nom_excel in dict_siren:
                            st.warning("Ce bailleur est déjà dans la base !")
                        else:
                            # 1. Créer la nouvelle ligne
                            nouvelle_ligne = pd.DataFrame([{
                                df_bailleurs_gsheet.columns[0]: nom_excel, 
                                df_bailleurs_gsheet.columns[1]: siren_api
                            }])
                            # 2. Ajouter à l'ancien tableau
                            df_updated = pd.concat([df_bailleurs_gsheet, nouvelle_ligne], ignore_index=True)
                            # 3. Écrire dans le Google Sheet !
                            conn.update(spreadsheet=SPREADSHEET_URL, worksheet="Confort", data=df_updated)
                            
                            # 4. Vider le cache pour forcer la relecture
                            st.cache_data.clear()
                            st.success(f"{nom_excel} a été ajouté définitivement !")
                            st.rerun()
                else:
                    st.warning("Aucun résultat.")
            else:
                st.error("Erreur API.")

    with col_list:
        st.subheader("📋 Bailleurs enregistrés")
        st.dataframe(df_bailleurs_gsheet, use_container_width=True)

# ==========================================
# ONGLET 1 : GÉNÉRATEUR D'EXCEL
# ==========================================
with tab_generateur:
    uploaded_file = st.file_uploader("Importer le fichier Excel (Liste globale)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df_source = pd.read_excel(uploaded_file)
            st.success(f"Fichier chargé ! ({len(df_source)} lignes)")

            # Formatage dates
            for col in ['Date réception', "Date d'engagement", 'Date prévisionnelle de réalisation', 'Date réelle de réalisation']:
                if col in df_source.columns:
                    df_source[col] = pd.to_datetime(df_source[col], errors='coerce').dt.date

            # --- TRAITEMENT CONFORT ---
            bailleurs_cibles = list(dict_siren.keys())
            df_confort = pd.DataFrame()
            
            if 'Bénéficiaire' in df_source.columns:
                df_confort = df_source[df_source['Bénéficiaire'].isin(bailleurs_cibles)].copy()
                if not df_confort.empty:
                    df_confort.insert(0, 'SIREN', df_confort['Bénéficiaire'].map(dict_siren))
                    df_confort.insert(1, 'BS Confort', df_confort['Bénéficiaire'])
                    
                    cols_attendues = ['SIREN', 'BS Confort', 'Date réception', 'Numéro dossier', 'Bénéficiaire', 'Stade']
                    df_confort = df_confort[[c for c in cols_attendues if c in df_confort.columns]]

            # --- TRAITEMENT EXPORT ---
            df_export = df_source.copy()
            if 'Contrôle' in df_export.columns:
                df_export = df_export[df_export['Contrôle'] != 'Non concerné']
            
            if not df_confort.empty and 'Numéro dossier' in df_export.columns:
                dossiers_confort = df_confort['Numéro dossier'].dropna().unique()
                df_export = df_export[~df_export['Numéro dossier'].isin(dossiers_confort)]

            # --- TÉLÉCHARGEMENTS ---
            st.divider()
            c1, c2 = st.columns(2)
            
            with c1:
                st.subheader("📊 Liste à exporter")
                st.text(f"{len(df_export)} lignes.")
                buffer_export = io.BytesIO()
                with pd.ExcelWriter(buffer_export, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                    df_export.to_excel(w, index=False, sheet_name='Liste à exporter')
                st.download_button("📥 Télécharger", buffer_export.getvalue(), "Liste_a_exporter.xlsx", use_container_width=True)
                
            with c2:
                st.subheader("🏢 Fichier Confort")
                st.text(f"{len(df_confort)} lignes.")
                if not df_confort.empty:
                    buffer_confort = io.BytesIO()
                    with pd.ExcelWriter(buffer_confort, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                        df_confort.to_excel(w, index=False, sheet_name='Confort')
                    st.download_button("📥 Télécharger", buffer_confort.getvalue(), "Fichier_Confort.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"Erreur : {e}")
