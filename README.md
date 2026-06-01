Markdown
# ⚡ Générateur d'Exports CEE (ODICEE)

Une application web interactive développée avec **Streamlit** pour automatiser le traitement, le tri et la génération de fichiers Excel collaboratifs dans le cadre de la gestion des dossiers CEE (Certificats d'Économies d'Énergie).

L'outil permet de transformer un export global brut en plusieurs fichiers Excel sectorisés (DCR, Confort, CDC, ADMIN) avec des mises en forme conditionnelles avancées pour le travail en équipe sur SharePoint.

## 🚀 Fonctionnalités Principales

- **📥 Import & Traitement de données :** Chargement d'un fichier Excel global et traitement instantané via Pandas.
- **🔄 Tri Automatique :** Séparation intelligente des dossiers selon des bases de données de bénéficiaires (Confort, CDC) et des filtres de mots-clés (ADMIN).
- **☁️ Synchronisation Google Sheets :** Gestion dynamique des bases de données et des filtres de mots-clés directement synchronisés avec un Google Sheet partagé.
- **🔍 API Recherche Entreprises :** Intégration de l'API gouvernementale (`recherche-entreprises.api.gouv.fr`) pour retrouver et ajouter automatiquement les Noms et SIREN des bailleurs.
- **📊 Tableaux de Synthèse Dynamiques :** Génération de tableaux croisés dynamiques avec indicateurs visuels (codes couleurs) pour évaluer la charge de travail par date de réception.
- **🤝 Optimisation pour SharePoint :** Génération de fichiers Excel (`xlsxwriter`) incluant des formules de **mise en forme conditionnelle dynamiques**. Si un collaborateur saisit ses initiales sur une ligne d'un dossier, toutes les lignes de ce même dossier se grisent automatiquement pour toute l'équipe.
- **📦 Export ZIP :** Téléchargement individuel ou global (archive ZIP) des fichiers générés.

## 🛠️ Technologies Utilisées

- **Python 3.9+**
- **Streamlit** : Interface utilisateur web (Frontend & Backend).
- **Pandas** : Manipulation et traitement des DataFrames.
- **XlsxWriter** : Génération de fichiers Excel formatés avec intégration de formules complexes.
- **Requests** : Appels HTTP vers l'API du gouvernement.
- **st-gsheets-connection** : Connexion native entre Streamlit et Google Sheets.
- **Pytz** : Gestion stricte des fuseaux horaires (Europe/Paris) lors du déploiement cloud.

## ⚙️ Installation en local

1. **Cloner le dépôt :**
   ```bash
   git clone [https://github.com/TonNomUtilisateur/nom-du-repo.git](https://github.com/TonNomUtilisateur/nom-du-repo.git)
   cd nom-du-repo
Créer un environnement virtuel (recommandé) :

Bash
python -m venv env
source env/bin/activate  # Sur Mac/Linux
env\Scripts\activate     # Sur Windows
Installer les dépendances :
Crée un fichier requirements.txt contenant les librairies nécessaires :

Plaintext
streamlit
pandas
xlsxwriter
requests
st-gsheets-connection
pytz
openpyxl
Puis lance l'installation :

Bash
pip install -r requirements.txt
🔒 Configuration (Google Sheets & Secrets)
L'application utilise Google Sheets comme base de données légère. Pour que l'application puisse lire et écrire dans le fichier Google Sheets, vous devez configurer les identifiants Google Service Account.

Créez un dossier .streamlit/ à la racine de votre projet.

Ajoutez un fichier secrets.toml à l'intérieur.

Renseignez vos identifiants Google Cloud (Service Account JSON) :

Ini, TOML
[connections.gsheets]
type = "service_account"
project_id = "TON_PROJECT_ID"
private_key_id = "TON_PRIVATE_KEY_ID"
private_key = "TON_PRIVATE_KEY"
client_email = "TON_CLIENT_EMAIL"
client_id = "TON_CLIENT_ID"
auth_uri = "[https://accounts.google.com/o/oauth2/auth](https://accounts.google.com/o/oauth2/auth)"
token_uri = "[https://oauth2.googleapis.com/token](https://oauth2.googleapis.com/token)"
auth_provider_x509_cert_url = "[https://www.googleapis.com/oauth2/v1/certs](https://www.googleapis.com/oauth2/v1/certs)"
client_x509_cert_url = "TON_CERT_URL"
(Note : Assurez-vous d'ajouter .streamlit/secrets.toml à votre .gitignore pour ne pas exposer vos clés publiquement).

🖥️ Utilisation
Pour lancer l'application en local :

Bash
streamlit run Export_liste.py
L'application s'ouvrira automatiquement dans votre navigateur web par défaut à l'adresse http://localhost:8501.

💡 Note sur le workflow collaboratif
Ce projet résout un problème classique du travail d'équipe sur SharePoint : l'impossibilité d'utiliser des macros VBA en ligne. L'outil génère nativement dans le .xlsx une formule COUNTIFS injectée via XlsxWriter dans la mise en forme conditionnelle. Ainsi, l'Excel est 100% compatible avec Excel Online tout en offrant une réactivité en temps réel pour l'attribution des dossiers.
