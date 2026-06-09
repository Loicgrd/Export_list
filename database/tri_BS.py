import pandas as pd

def creer_base_bs(chemin_fichier_entree, chemin_fichier_sortie):
    print(f"Chargement du fichier : {chemin_fichier_entree}...")
    
    # Lecture du fichier (adaptation si c'est un vrai Excel ou un CSV déguisé en .xlsx)
    try:
        df = pd.read_excel(chemin_fichier_entree)
    except ValueError:
        df = pd.read_csv(chemin_fichier_entree, sep=',')

    # Les noms exacts des colonnes dans ton export
    col_siren = "SIREN de l'organisme"
    col_nom = "Organisme"
    
    # 1. On isole les deux colonnes qui nous intéressent et on enlève les lignes vides
    df_reduit = df[[col_siren, col_nom]].dropna()
    
    # 2. On dédoublonne en se basant sur le SIREN
    df_unique = df_reduit.drop_duplicates(subset=[col_siren])
    
    # 3. On renomme pour ta structure de BDD
    df_final = df_unique.rename(columns={
        col_siren: "SIREN",
        col_nom: "Nom BS"
    })
    
    # 4. Nettoyage du format SIREN (conversion en texte brut, suppression des ".0" si Pandas l'a lu en float)
    df_final['SIREN'] = df_final['SIREN'].astype(str).str.replace(r'\.0$', '', regex=True)
    
    # 5. Export vers un nouvel Excel propre
    df_final.to_excel(chemin_fichier_sortie, index=False)
    
    print("\nExtraction terminée avec succès !")
    print(f"Fichier sauvegardé sous : {chemin_fichier_sortie}")
    print(f"Nombre de Bailleurs Sociaux uniques récupérés : {len(df_final)}")

# --- Exécution du script ---
if __name__ == "__main__":
    # Remplace par le chemin exact de ton fichier si besoin
    fichier_source = "database/Liste_BS_non_trie.xlsx" 
    fichier_resultat = "database/Base_De_Donnees_BS_Unique.xlsx"
    
    creer_base_bs(fichier_source, fichier_resultat)