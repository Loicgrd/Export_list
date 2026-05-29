import streamlit as st
import pandas as pd
import io
import requests

# ==========================================
# CONFIGURATION & INITIALISATION
# ==========================================
st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

# Initialisation de la mémoire (session_state) pour garder les bailleurs actifs
if 'bailleurs_confort' not in st.session_state:
    st.session_state.bailleurs_confort = {
        "INOLYA": "780705703",
        "IMMOBILIERE RHONE ALPES SA D'HLM": "661750067",
        "OPH TROYES AUBE HABITAT": "902718998",
        "SEINE-SAINT-DENIS HABITAT": "279300198"
    }

st.title("Générateur d'Exports CEE")

# Création de deux onglets de navigation
tab_generateur, tab_reglages = st.tabs(["📊 Générateur d'Exports", "⚙️ Gestion des Bailleurs Confort"])

# ==========================================
# ONGLET 2 : GESTION DES BAILLEURS VIA API
# ==========================================
with tab_reglages:
    st.header("Paramétrage de la liste 'Confort'")
    st.markdown("Recherchez un bailleur social via l'API du gouvernement et ajoutez-le à la liste de traitement.")
    
    col_search, col_list = st.columns([1, 1])
    
    with col_search:
        st.subheader("🔍 Ajouter un bailleur")
        query = st.text_input("Entrez le SIREN ou le nom de l'organisme :")
        
        if query:
            try:
                # Requête vers l'API du gouvernement
                response = requests.get(f"https://recherche-entreprises.api.gouv.fr/search?q={query}")
                
                if response.status_code == 200:
                    results = response.json().get("results", [])
                    
                    if results:
                        # Création d'un dictionnaire pour la liste déroulante
                        options = {f"{r.get('nom_complet')} (SIREN: {r.get('siren')})": r for r in results}
                        selected_label = st.selectbox("Sélectionnez l'organisme trouvé :", list(options.keys()))
                        selected_data = options[selected_label]
                        
                        siren_api = selected_data.get('siren')
                        # On essaie de prendre le sigle (souvent utilisé dans les Excel), sinon le nom complet
                        nom_defaut = selected_data.get('sigle') or selected_data.get('nom_raison_sociale')
                        
                        st.info("⚠️ **Important :** Le nom ci-dessous doit correspondre **exactement** à ce qui est écrit dans la colonne 'Bénéficiaire' de votre fichier Excel.")
                        nom_excel = st.text_input("Nom exact dans l'Excel :", value=nom_defaut)
                        
                        if st.button("➕ Ajouter à la liste Confort", type="primary"):
                            st.session_state.bailleurs_confort[nom_excel] = siren_api
                            st.success(f"{nom_excel} a été ajouté avec succès !")
                            st.rerun() # Rafraîchit la page pour mettre à jour la liste à droite
                    else:
                        st.warning("Aucun résultat trouvé pour cette recherche.")
                else:
                    st.error("Erreur de communication avec l'API du gouvernement.")
            except Exception as e:
                st.error(f"Une erreur réseau est survenue : {e}")

    with col_list:
        st.subheader("📋 Liste actuelle active")
        if not st.session_state.bailleurs_confort:
            st.info("La liste est vide. Ajoutez des bailleurs via la recherche.")
        else:
            # Affichage de la liste avec bouton de suppression
            for nom, siren in list(st.session_state.bailleurs_confort.items()):
                c1, c2 = st.columns([4, 1])
                c1.write(f"🏢 **{nom}** *(SIREN: {siren})*")
                if c2.button("❌", key=f"del_{siren}_{nom}", help="Supprimer"):
                    del st.session_state.bailleurs_confort[nom]
                    st.rerun()

# ==========================================
# ONGLET 1 : GÉNÉRATEUR D'EXCEL
# ==========================================
with tab_generateur:
    st.markdown("Chargez votre fichier contenant la **Liste globale** pour générer les fichiers d'export.")

    uploaded_file = st.file_uploader("Importer le fichier Excel source (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            # Lecture
            df_source = pd.read_excel(uploaded_file)
            st.success(f"Fichier chargé avec succès ! ({len(df_source)} lignes trouvées)")

            # Formatage des dates
            colonnes_dates = ['Date réception', "Date d'engagement", 'Date prévisionnelle de réalisation', 'Date réelle de réalisation']
            for col in colonnes_dates:
                if col in df_source.columns:
                    df_source[col] = pd.to_datetime(df_source[col], errors='coerce').dt.date

            # --- TRAITEMENT CONFORT ---
            dict_siren = st.session_state.bailleurs_confort # On utilise la liste dynamique !
            bailleurs_cibles = list(dict_siren.keys())

            df_confort = pd.DataFrame()
            if 'Bénéficiaire' in df_source.columns:
                df_confort = df_source[df_source['Bénéficiaire'].isin(bailleurs_cibles)].copy()
                if not df_confort.empty:
                    df_confort.insert(0, 'SIREN', df_confort['Bénéficiaire'].map(dict_siren))
                    df_confort.insert(1, 'BS Confort', df_confort['Bénéficiaire'])
                    
                    colonnes_attendues = ['SIREN', 'BS Confort', 'Date réception', 'Numéro dossier', 'Bénéficiaire', 'Stade']
                    colonnes_dispos = [c for c in colonnes_attendues if c in df_confort.columns]
                    df_confort = df_confort[colonnes_dispos]

            # --- TRAITEMENT EXPORT ---
            df_export = df_source.copy()
            if 'Contrôle' in df_export.columns:
                df_export = df_export[df_export['Contrôle'] != 'Non concerné']
            
            # Exclusion des dossiers Confort
            if not df_confort.empty and 'Numéro dossier' in df_export.columns and 'Numéro dossier' in df_confort.columns:
                liste_dossiers_confort = df_confort['Numéro dossier'].dropna().unique()
                df_export = df_export[~df_export['Numéro dossier'].isin(liste_dossiers_confort)]

            # --- TÉLÉCHARGEMENTS ---
            st.divider()
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📊 Liste à exporter")
                st.text(f"{len(df_export)} lignes (Excluant les dossiers Confort).")
                st.dataframe(df_export.head(3))
                
                buffer_export = io.BytesIO()
                with pd.ExcelWriter(buffer_export, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Liste à exporter')
                
                st.download_button("📥 Télécharger la Liste à exporter", buffer_export.getvalue(), "Liste_a_exporter.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                
            with col2:
                st.subheader("🏢 Fichier Confort")
                st.text(f"{len(df_confort)} lignes identifiées.")
                st.dataframe(df_confort.head(3) if not df_confort.empty else df_confort)
                
                if not df_confort.empty:
                    buffer_confort = io.BytesIO()
                    with pd.ExcelWriter(buffer_confort, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                        df_confort.to_excel(writer, index=False, sheet_name='Confort')
                    
                    st.download_button("📥 Télécharger le fichier Confort", buffer_confort.getvalue(), "Fichier_Confort.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        except Exception as e:
            st.error(f"Une erreur est survenue lors de l'analyse du fichier : {e}")
