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
                        st.warning("Aucun résultat trouvé pour cette recherche
