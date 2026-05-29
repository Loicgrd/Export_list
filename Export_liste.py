import streamlit as st
import pandas as pd
import io

# Configuration de la page
st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

st.title("Générateur d'Exports : Liste à Exporter & Confort")
st.markdown("Chargez votre fichier contenant la **Liste globale** pour générer les fichiers d'export.")

# 1. Upload du fichier
uploaded_file = st.file_uploader("Importer le fichier Excel source (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Lecture de la feuille principale
        df_source = pd.read_excel(uploaded_file)
        st.success(f"Fichier chargé avec succès ! ({len(df_source)} lignes trouvées)")

        # ==========================================
        # TRAITEMENT 1 : LISTE À EXPORTER
        # ==========================================
        df_export = df_source.copy()
        
        # Exemple de filtre métier (à adapter selon ton besoin exact)
        # Ici on retire par exemple les dossiers "Non concerné"
        if 'Contrôle' in df_export.columns:
            df_export = df_export[df_export['Contrôle'] != 'Non concerné']
        
        # Si tu as une colonne "Prioritaire", tu peux trier ici :
        # df_export = df_export.sort_values(by='Prioritaire', na_position='last')

        # ==========================================
        # TRAITEMENT 2 : CONFORT
        # ==========================================
        # Définition des bailleurs cibles pour l'onglet Confort
        bailleurs_cibles = [
            "INOLYA", 
            "IMMOBILIERE RHONE ALPES SA D'HLM", 
            "OPH TROYES AUBE HABITAT", 
            "SEINE-SAINT-DENIS HABITAT"
        ]
        
        # Dictionnaire de mapping SIREN (à compléter ou à remplacer par un merge avec un Excel de réf)
        dict_siren = {
            "INOLYA": "780705703",
            "IMMOBILIERE RHONE ALPES SA D'HLM": "661750067",
            "OPH TROYES AUBE HABITAT": "902718998",
            "SEINE-SAINT-DENIS HABITAT": "279300198"
        }

        df_confort = pd.DataFrame()
        if 'Bénéficiaire' in df_source.columns:
            # Filtrer sur les bailleurs
            df_confort = df_source[df_source['Bénéficiaire'].isin(bailleurs_cibles)].copy()
            
            # Création des colonnes spécifiques à "Confort"
            df_confort.insert(0, 'SIREN', df_confort['Bénéficiaire'].map(dict_siren))
            df_confort.insert(1, 'BS Confort', df_confort['Bénéficiaire']) # À adapter si le nom diffère
            
            # Sélection et réorganisation des colonnes cibles
            colonnes_attendues = ['SIREN', 'BS Confort', 'Date réception', 'Numéro dossier', 'Bénéficiaire', 'Stade']
            colonnes_dispos = [c for c in colonnes_attendues if c in df_confort.columns]
            df_confort = df_confort[colonnes_dispos]

        # ==========================================
        # INTERFACE DE TÉLÉCHARGEMENT
        # ==========================================
        st.divider()
        col1, col2 = st.columns(2)
        
        # --- Bloc Gauche : Liste à exporter ---
        with col1:
            st.subheader("📊 Liste à exporter")
            st.text(f"{len(df_export)} lignes prêtes pour l'export.")
            st.dataframe(df_export.head(3))
            
            # Conversion en Excel en mémoire
            buffer_export = io.BytesIO()
            with pd.ExcelWriter(buffer_export, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Liste à exporter')
            
            st.download_button(
                label="📥 Télécharger la Liste à exporter",
                data=buffer_export.getvalue(),
                file_name="Liste_a_exporter.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        # --- Bloc Droit : Confort ---
        with col2:
            st.subheader("🏢 Confort (Bailleurs ciblés)")
            st.text(f"{len(df_confort)} lignes identifiées.")
            st.dataframe(df_confort.head(3))
            
            # Conversion en Excel en mémoire
            buffer_confort = io.BytesIO()
            with pd.ExcelWriter(buffer_confort, engine='xlsxwriter') as writer:
                df_confort.to_excel(writer, index=False, sheet_name='Confort')
                
            st.download_button(
                label="📥 Télécharger le fichier Confort",
                data=buffer_confort.getvalue(),
                file_name="Fichier_Confort.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Une erreur est survenue lors de l'analyse du fichier : {e}")
