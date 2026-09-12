# Installer les trois profils

## Préparer le Synology et télécharger

1. Déployer le service selon [INSTALLATION_NAS.md](Serveur/INSTALLATION_NAS.md). Pour les essais, utiliser un partage et une base distincts contenant seulement des données fictives.
2. Depuis le dépôt privé `Mandagoutolivier/chatgpt`, télécharger le ZIP, l'extraire sur le PC, puis ouvrir `versions/cabinet-2026.09.12`. Conserver toute son arborescence.
3. Fermer Word et Excel. Word et Excel doivent être installés. Le constructeur nécessite l'accès au modèle objet du projet VBA, à autoriser dans le centre de gestion de la confidentialité pendant la construction. Le script ne désactive pas les protections Office/Windows. Si une stratégie d'entreprise interdit les scripts/macros, faire signer ou autoriser le paquet selon cette stratégie.
4. À domicile, connecter le VPN du cabinet et vérifier l'accès au partage UNC et à l'adresse HTTPS du service.

## Choisir le poste

Double-cliquer `Installer.cmd` : **1 domicile, 2 secrétariat, 3 cabinet médecin**. Le mode par défaut prépare les fichiers et affiche le dossier obtenu. Les anciens compléments restent actifs jusqu'à l'activation.

Équivalents PowerShell, depuis le dossier de version :

```powershell
.\Installer.ps1 -Profil Domicile
.\Installer.ps1 -Profil Secretariat
.\Installer.ps1 -Profil Cabinet
```

Word et Excel sont contrôlés dans les trois profils. Le poste médecin reçoit `CabinetUnifie.dotm` et `sqlite3.exe`. Le secrétariat reçoit `Cabinet.xlsm` et un raccourci. `Normal.dotm` est conservé.

## Tester puis activer le même dossier

Compiler les projets VBA Word et Excel préparés avec **Débogage > Compiler**, enregistrer, fermer et rouvrir. Dérouler [RECETTE_WINDOWS.md](RECETTE_WINDOWS.md). La préparation n'est pas une preuve de compilation. Exemple pour domicile :

```powershell
.\Build\valider_preparation.ps1 -DossierPrepare 'C:\chemin\du\dossier\prepare' -CompilationWordValidee -CompilationExcelValidee -RecetteValidee
.\Installer.ps1 -Profil Domicile -Mode Installation -DossierPrepare 'C:\chemin\du\dossier\prepare' -RacineNas '\\DS224\CabinetCardio' -UrlService 'https://adresse-interne-du-service'
```

Remplacer les chemins et l'adresse d'exemple. Pour secrétariat, utiliser `-Profil Secretariat` et `-CompilationExcelValidee`. Pour médecin, `-Profil Cabinet` et `-CompilationWordValidee`. Le validateur rouvre les binaires, compare leur source au manifeste et vérifie les références Office. La compilation et les essais physiques restent attestés par l'opérateur ; ils ne sont pas simulés.

À l'activation, le script demande le jeton du compte NAS par saisie masquée, ou accepte `-FichierJeton` pointant vers un fichier local protégé. Il contrôle HTTPS, la version du protocole et les rôles. Il conserve le jeton dans `%APPDATA%\CabinetCardio\service.token` avec des droits limités au compte Windows, SYSTEM et administrateurs. Ne pas saisir un jeton dans la ligne de commande ni le déposer dans GitHub.

Le service doit avoir été initialisé et les classeurs importés avant cette activation. Les ressources initiales sont copiées seulement si absentes ; les bases Excel historiques ne sont jamais remplacées par l'installateur.

## Réglages du poste médecin

Le dossier ECG se règle avec `-DossierGdt 'C:\Mandagout'`. Vérifier dans Resting12Lead que ce dossier et le nom `IMPORT.GDT` correspondent à son interface GDT. Le logiciel ne suppose pas un accès documenté à sa base interne propriétaire.

Affecter les touches Dragon/PowerMic aux macros `Unifie_A_NouvelleLettre`, `Unifie_B_FormuleAppel`, `Unifie_C_InsererPatient`, `Unifie_D_Finaliser`. La clé OpenAI demeure locale au poste médecin, via `OPENAI_API_KEY` ou `%APPDATA%\CabinetCardio\openai.key`.

## Feuille de soins et comptabilité

Les boutons ajoutés à l'accueil donnent accès à l'assuré distinct, aux paramètres des destinataires et à l'encaissement d'une séance. Les réglages inhabituels restent hors de la saisie courante.

Compléter RPPS et numéro AM dans la configuration médecin. Les positions du formulaire sont dans `Config\cerfa_positions.txt` sur le NAS. Les champs `PATIENT_NOM`, `PATIENT_DDN` sont requis pour un assuré distinct ; `MEDECIN_RPPS`, `MEDECIN_AM` sont requis si le praticien n'est pas préimprimé. Aucune coordonnée physique inconnue n'a été inventée. Ajouter les positions après essai sur le formulaire utilisé.

Le calage et `poste.ini` sont **locaux à chaque imprimante/poste**. Après essai papier, renseigner :

```ini
[CERFA]
CalageValide=1
PraticienPreimprime=1
Imprimante=nom exact de l'imprimante si nécessaire
```

Mettre `PraticienPreimprime=1` seulement si les identifiants du médecin figurent déjà correctement sur les feuilles. Un NIR invalide ou une fiche assuré incomplète bloque l'impression ; l'absence de NIR ne bloque pas la consultation. Plus de quatre actes est refusé explicitement et demande une seconde feuille. Aucun connecteur FSE/CPS/SESAM-Vitale n'est livré.

Le journal exporté dans Excel est une copie ; le service conserve les écritures originales. Les montants sont numériques à l'export. L'encaissement couvre un règlement intégral ; le suivi de paiements partiels n'est pas implémenté.

## Retour arrière

Chaque activation écrit un dossier `%APPDATA%\CabinetCardio\Sauvegardes\...`. Simulation, puis restauration locale :

```powershell
.\Build\restaurer_poste.ps1 -DossierSauvegarde 'C:\chemin\sauvegarde'
.\Build\restaurer_poste.ps1 -DossierSauvegarde 'C:\chemin\sauvegarde' -Appliquer
```

Ce retour restaure les fichiers du poste, pas les transactions du NAS. Après migration, ne pas faire fonctionner simultanément les anciens clients Excel et les nouveaux clients serveur. Revenir à l'ancien logiciel nécessite un arrêt coordonné et une décision sur les données saisies depuis la migration.
