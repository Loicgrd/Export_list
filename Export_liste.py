import streamlit as st
import pandas as pd
import io
import requests
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

# ==========================================
# CONNEXION AU GOOGLE SHEET
# ==========================================
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1pkj6frncXmzUUVAClAWp63HY_UpQUahOCLz19w-UseI/edit?usp=sharing"
conn = st.connection("gsheets", type=GSheetsConnection)

# Fonction modifiée pour charger une feuille spécifique (Confort ou CDC)
@st.cache_data(ttl=10)
def load_worksheet(sheet_name):
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name)
        
        # Si la feuille est totalement vide, on crée les colonnes par défaut
        if df.empty or len(df.columns) < 3:
            df = pd.DataFrame(columns=["Nom", "SIREN", "Date d'ajout"])
            
        # Si la colonne Date d'ajout manque mais que les autres y sont
        if len(df.columns) < 3:
            df["Date d'ajout"] = ""
            
        df = df.dropna(subset=[df.columns[0]])
        df = df.iloc[:, :3]
        
        dict_siren = dict(zip(df.iloc[:, 0], df.iloc[:, 1].astype(str)))
        return dict_siren, df
    except Exception as e:
        st.error(f"Erreur de lecture de la feuille '{sheet_name}' : {e}")
        return {}, pd.DataFrame(columns=["Nom", "SIREN", "Date d'ajout"])

# Chargement des deux bases de données
dict_confort, df_confort_gsheet = load_worksheet("Confort")
dict_cdc, df_cdc_gsheet = load_worksheet("CDC")

st.title("Générateur d'Exports CEE")

# Ajout du troisième onglet
tab_generateur, tab_confort, tab_cdc = st.tabs([
    "📊 Générateur d'Exports", 
    "⚙️ Base Confort", 
    "⚙️ Base CDC"
])


# ==========================================
# FONCTION COMMUNE DE GESTION (POUR CONFORT ET CDC)
# ==========================================
def afficher_gestion_base(sheet_name, df_gsheet):
    st.header(f"Base de données '{sheet_name}'")

    # Clés dynamiques pour que Confort et CDC ne se mélangent pas
    state_recherche = f"recherche_{sheet_name}"
    state_trouves = f"trouves_{sheet_name}"

    if state_recherche not in st.session_state:
        st.session_state[state_recherche] = False
        st.session_state[state_trouves] = []

    # --- SECTION MULTI-AJOUT ---
    st.subheader("➕ Ajouter plusieurs bailleurs (par SIREN)")
    liste_sirens_brut = st.text_area(f"Collez vos SIREN ici :", key=f"input_siren_{sheet_name}")
    
    # ÉTAPE 1 : LA RECHERCHE
    if st.button("🔍 Rechercher les SIREN", key=f"btn_search_{sheet_name}"):
        sirens = [s.strip() for s in liste_sirens_brut.replace('\n', ',').split(',') if s.strip()]
        
        trouves = []
        with st.spinner("Recherche dans la base du Gouvernement..."):
            for s in sirens:
                resp = requests.get(f"https://recherche-entreprises.api.gouv.fr/search?q={s}")
                if resp.status_code == 200 and resp.json().get("results"):
                    res = resp.json()["results"][0]
                    trouves.append({
                        "Nom": res.get('sigle') or res.get('nom_raison_sociale'),
                        "SIREN": res.get('siren')
                    })
        
        st.session_state[state_trouves] = trouves
        st.session_state[state_recherche] = True

    # ÉTAPE 2 : LA CONFIRMATION
    if st.session_state[state_recherche]:
        if st.session_state[state_trouves]:
            st.success(f"✅ {len(st.session_state[state_trouves])} bailleur(s) trouvé(s) ! Veuillez vérifier avant d'ajouter :")
            st.dataframe(pd.DataFrame(st.session_state[state_trouves]), hide_index=True)
            
            col_btn1, col_btn2 = st.columns([0.2, 0.8])
            with col_btn1:
                if st.button("✅ Confirmer l'ajout", key=f"btn_conf_{sheet_name}", type="primary"):
                    date_jour = datetime.now().strftime("%d/%m/%Y")
                    nom_col, siren_col, date_col = df_gsheet.columns[0], df_gsheet.columns[1], df_gsheet.columns[2]
                    
                    nouveaux_bailleurs = [{nom_col: b['Nom'], siren_col: b['SIREN'], date_col: date_jour} for b in st.session_state[state_trouves]]
                    df_updated = pd.concat([df_gsheet, pd.DataFrame(nouveaux_bailleurs)], ignore_index=True)
                    
                    conn.update(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, data=df_updated)
                    
                    st.session_state[state_recherche] = False
                    st.cache_data.clear()
                    st.rerun()
                    
            with col_btn2:
                if st.button("❌ Annuler", key=f"btn_annul_{sheet_name}"):
                    st.session_state[state_recherche] = False
                    st.rerun()
        else:
            st.warning("❌ Aucun bailleur trouvé pour ces SIREN.")
            if st.button("Nouvelle recherche", key=f"btn_nouv_{sheet_name}"):
                st.session_state[state_recherche] = False
                st.rerun()

    st.divider()

    # --- LISTE ACTUELLE ---
    st.subheader("📋 Liste actuelle")
    if not df_gsheet.empty:
        df_display = df_gsheet.copy()
        col_nom, col_siren, col_date = df_display.columns[0], df_display.columns[1], df_display.columns[2]
        df_display = df_display[[col_date, col_nom, col_siren]]
        st.dataframe(df_display, use_container_width=True, hide_index=True)
    else:
        st.info("La liste est actuellement vide.")

    # --- SECTION SUPPRESSION ---
    st.subheader("🗑️ Supprimer des bailleurs")
    if not df_gsheet.empty:
        bailleurs_a_supprimer = st.multiselect("Sélectionner les bailleurs à supprimer :", options=df_gsheet.iloc[:, 0].tolist(), key=f"del_{sheet_name}")
        if st.button("Supprimer la sélection", type="primary", key=f"btn_del_{sheet_name}"):
            df_updated = df_gsheet[~df_gsheet.iloc[:, 0].isin(bailleurs_a_supprimer)]
            conn.update(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, data=df_updated)
            st.cache_data.clear()
            st.rerun()


# ==========================================
# AFFICHAGE DES ONGLETS DE RÉGLAGES
# ==========================================
with tab_confort:
    afficher_gestion_base("Confort", df_confort_gsheet)

with tab_cdc:
    afficher_gestion_base("CDC", df_cdc_gsheet)


# ==========================================
# ONGLET 1 : GÉNÉRATEUR D'EXCEL
# ==========================================
with tab_generateur:
    uploaded_file = st.file_uploader("Importer le fichier Excel (Liste globale)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df_source = pd.read_excel(uploaded_file)
            st.success(f"Fichier chargé ! ({len(df_source)} lignes)")

            for col in ['Date réception', "Date d'engagement", 'Date prévisionnelle de réalisation', 'Date réelle de réalisation']:
                if col in df_source.columns:
                    df_source[col] = pd.to_datetime(df_source[col], errors='coerce').dt.date

            # --- EXTRACTION CONFORT ---
            df_confort = pd.DataFrame()
            if 'Bénéficiaire' in df_source.columns:
                df_confort = df_source[df_source['Bénéficiaire'].isin(dict_confort.keys())].copy()
                if not df_confort.empty:
                    df_confort.insert(0, 'SIREN', df_confort['Bénéficiaire'].map(dict_confort))
                    df_confort.insert(1, 'BS Confort', df_confort['Bénéficiaire'])
                    cols_attendues = ['SIREN', 'BS Confort', 'Date réception', 'Numéro dossier', 'Bénéficiaire', 'Stade']
                    df_confort = df_confort[[c for c in cols_attendues if c in df_confort.columns]]

            # --- EXTRACTION CDC ---
            df_cdc = pd.DataFrame()
            if 'Bénéficiaire' in df_source.columns:
                df_cdc = df_source[df_source['Bénéficiaire'].isin(dict_cdc.keys())].copy()
                if not df_cdc.empty:
                    df_cdc.insert(0, 'SIREN', df_cdc['Bénéficiaire'].map(dict_cdc))
                    df_cdc.insert(1, 'BS CDC', df_cdc['Bénéficiaire'])
                    cols_attendues_cdc = ['SIREN', 'BS CDC', 'Date réception', 'Numéro dossier', 'Bénéficiaire', 'Stade']
                    df_cdc = df_cdc[[c for c in cols_attendues_cdc if c in df_cdc.columns]]

            # --- LISTE À EXPORTER (Filtrée) ---
            df_export = df_source.copy()
            if 'Contrôle' in df_export.columns:
                df_export = df_export[df_export['Contrôle'] != 'Non concerné']
            
            if 'Numéro dossier' in df_export.columns:
                if not df_confort.empty:
                    dossiers_confort = df_confort['Numéro dossier'].dropna().unique()
                    df_export = df_export[~df_export['Numéro dossier'].isin(dossiers_confort)]
                if not df_cdc.empty:
                    dossiers_cdc = df_cdc['Numéro dossier'].dropna().unique()
                    df_export = df_export[~df_export['Numéro dossier'].isin(dossiers_cdc)]

            # --- TÉLÉCHARGEMENTS EN 3 COLONNES ---
            st.divider()
            c1, c2, c3 = st.columns(3)
            
            with c1:
                st.subheader("📊 Liste Principale")
                st.text(f"{len(df_export)} lignes.")
                buffer_export = io.BytesIO()
                with pd.ExcelWriter(buffer_export, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                    df_export.to_excel(w, index=False, sheet_name='Liste à exporter')
                st.download_button("📥 Télécharger Liste", buffer_export.getvalue(), "Liste_a_exporter.xlsx", use_container_width=True)
                
            with c2:
                st.subheader("🏢 Fichier Confort")
                st.text(f"{len(df_confort)} lignes.")
                if not df_confort.empty:
                    buffer_confort = io.BytesIO()
                    with pd.ExcelWriter(buffer_confort, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                        df_confort.to_excel(w, index=False, sheet_name='Confort')
                    st.download_button("📥 Télécharger Confort", buffer_confort.getvalue(), "Fichier_Confort.xlsx", use_container_width=True)
                else:
                    st.info("Aucun bailleur Confort trouvé.")
                    
            with c3:
                st.subheader("🏛️ Fichier CDC")
                st.text(f"{len(df_cdc)} lignes.")
                if not df_cdc.empty:
                    buffer_cdc = io.BytesIO()
                    with pd.ExcelWriter(buffer_cdc, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                        df_cdc.to_excel(w, index=False, sheet_name='CDC')
                    st.download_button("📥 Télécharger CDC", buffer_cdc.getvalue(), "Fichier_CDC.xlsx", use_container_width=True)
                else:
                    st.info("Aucun bailleur CDC trouvé.")

        except Exception as e:
            st.error(f"Erreur : {e}")

        except Exception as e:
            st.error(f"Erreur : {e}")
