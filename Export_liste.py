import streamlit as st
import pandas as pd
import io
import requests
from streamlit_gsheets import GSheetsConnection

# ==========================================
# CONFIGURATION DE LA PAGE
# ==========================================
st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

# ==========================================
# CONNEXION AU GOOGLE SHEET
# ==========================================
# L'URL de ton fichier de base de données
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1pkj6frncXmzUUVAClAWp63HY_UpQUahOCLz19w-UseI/edit?usp=sharing"

# Initialisation de la connexion GSheets
conn = st.connection("gsheets", type=GSheetsConnection)

# Fonction de lecture avec mise en cache
@st.cache_data(ttl=10)
def load_bailleurs():
    try:
        # On cible explicitement l'onglet "Confort"
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="Confort", usecols=[0, 1])
        
        # S'il n'y a pas de colonnes, on crée une structure par défaut
        if df.empty or len(df.columns) < 2:
            df = pd.DataFrame(columns=["Nom du Bailleur", "SIREN"])
            
        df = df.dropna(subset=[df.columns[0]])
        # On crée un dictionnaire { "NOM": "SIREN" }
        return dict(zip(df.iloc[:, 0], df.iloc[:, 1].astype(str))), df
    except Exception as e:
        st.error(f"Erreur de lecture du Google Sheet : {e}")
        return {}, pd.DataFrame(columns=["Nom du Bailleur", "SIREN"])

# Chargement immédiat des données
dict_siren, df_bailleurs_gsheet = load_bailleurs()

# ==========================================
# INTERFACE UTILISATEUR
# ==========================================
st.title("Générateur d'Exports CEE")

tab_generateur, tab_reglages = st.tabs(["📊 Générateur d'Exports", "⚙️ Base de données Bailleurs"])

# ------------------------------------------
# ONGLET 1 : GESTION DANS GOOGLE SHEETS
# ------------------------------------------
with tab_reglages:
    st.header("Base de données 'Confort'")
    st.markdown("Recherchez un bailleur via l'API et enregistrez-le définitivement dans votre Google Sheet.")
    
    col_search, col_list = st.columns([1, 1])
    
    with col_search:
        st.subheader("🔍 Ajouter à la base")
        query = st.text_input("Entrez le SIREN ou le nom de l'organisme (ex: AGEN HABITAT) :")
        
        if query:
            try:
                response = requests.get(f"https://recherche-entreprises.api.gouv.fr/search?q={query}")
                if response.status_code == 200:
                    results = response.json().get("results", [])
                    if results:
                        options = {f"{r.get('nom_complet')} (SIREN: {r.get('siren')})": r for r in results}
                        selected_label = st.selectbox("Sélectionnez l'organisme :", list(options.keys()))
                        selected_data = options[selected_label]
                        
                        siren_api = selected_data.get('siren')
                        nom_defaut = selected_data.get('sigle') or selected_data.get('nom_raison_sociale')
                        
                        st.info("⚠️ Le nom ci-dessous doit correspondre **exactement** à la colonne 'Bénéficiaire' de votre Excel.")
                        nom_excel = st.text_input("Nom exact dans l'Excel :", value=nom_defaut)
                        
                        if st.button("➕ Enregistrer dans le Google Sheet", type="primary"):
                            if nom_excel in dict_siren:
                                st.warning("Ce bailleur est déjà dans la base !")
                            else:
                                # Préparation de la nouvelle ligne
                                nom_col_1 = df_bailleurs_gsheet.columns[0]
                                nom_col_2 = df_bailleurs_gsheet.columns[1]
                                
                                nouvelle_ligne = pd.DataFrame([{
                                    nom_col_1: nom_excel, 
                                    nom_col_2: siren_api
                                }])
                                
                                df_updated = pd.concat([df_bailleurs_gsheet, nouvelle_ligne], ignore_index=True)
                                
                                # Écriture dans l'onglet "Confort" du Google Sheet
                                conn.update(spreadsheet=SPREADSHEET_URL, worksheet="Confort", data=df_updated)
                                
                                # Vidage du cache pour forcer la mise à jour visuelle
                                st.cache_data.clear()
                                st.success(f"{nom_excel} a été ajouté définitivement !")
                                st.rerun()
                    else:
                        st.warning("Aucun résultat trouvé pour cette recherche.")
                else:
                    st.error("Erreur de communication avec l'API du gouvernement.")
            except Exception as e:
                st.error(f"Une erreur réseau est survenue : {e}")

    with col_list:
        st.subheader("📋 Bailleurs enregistrés")
        st.dataframe(df_bailleurs_gsheet, use_container_width=True)

# ------------------------------------------
# ONGLET 2 : GÉNÉRATEUR D'EXCEL
# ------------------------------------------
with tab_generateur:
    st.markdown("Chargez votre **Liste globale** pour générer les fichiers d'export et Confort.")
    
    uploaded_file = st.file_uploader("Importer le fichier Excel source (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df_source = pd.read_excel(uploaded_file)
            st.success(f"Fichier chargé avec succès ! ({len(df_source)} lignes trouvées)")

            # Formatage des dates (suppression des heures)
            colonnes_dates = ['Date réception', "Date d'engagement", 'Date prévisionnelle de réalisation', 'Date réelle de réalisation']
            for col in colonnes_dates:
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
                    colonnes_dispos = [c for c in cols_attendues if c in df_confort.columns]
                    df_confort = df_confort[colonnes_dispos]

            # --- TRAITEMENT EXPORT ---
            df_export = df_source.copy()
            if 'Contrôle' in df_export.columns:
                df_export = df_export[df_export['Contrôle'] != 'Non concerné']
            
            # Exclusion stricte des dossiers partis dans Confort
            if not df_confort.empty and 'Numéro dossier' in df_export.columns:
                dossiers_confort = df_confort['Numéro dossier'].dropna().unique()
                df_export = df_export[~df_export['Numéro dossier'].isin(dossiers_confort)]

            # --- TÉLÉCHARGEMENTS ---
            st.divider()
            c1, c2 = st.columns(2)
            
            with c1:
                st.subheader("📊 Liste à exporter")
                st.text(f"{len(df_export)} lignes (Excluant les dossiers Confort).")
                st.dataframe(df_export.head(3))
                
                buffer_export = io.BytesIO()
                with pd.ExcelWriter(buffer_export, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                    df_export.to_excel(w, index=False, sheet_name='Liste à exporter')
                st.download_button("📥 Télécharger la Liste à exporter", buffer_export.getvalue(), "Liste_a_exporter.xlsx", use_container_width=True)
                
            with c2:
                st.subheader("🏢 Fichier Confort")
                st.text(f"{len(df_confort)} lignes identifiées.")
                st.dataframe(df_confort.head(3) if not df_confort.empty else df_confort)
                
                if not df_confort.empty:
                    buffer_confort = io.BytesIO()
                    with pd.ExcelWriter(buffer_confort, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                        df_confort.to_excel(w, index=False, sheet_name='Confort')
                    st.download_button("📥 Télécharger le fichier Confort", buffer_confort.getvalue(), "Fichier_Confort.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"Erreur lors du traitement du fichier : {e}")
