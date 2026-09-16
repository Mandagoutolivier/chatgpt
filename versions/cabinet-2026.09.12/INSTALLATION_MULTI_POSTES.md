# Installer les trois profils

**U2b (`2026.09.16-u2b`) reste en recette isolée. L'installation clinique actuelle doit être conservée. Les premiers essais ont lieu sur un PC secondaire formellement identifié — le PC de l'étage s'il est bien distinct — et sur `cabinetcardio-test-u2`, jamais sur le poste indispensable aux consultations du lendemain.**

## Lanceur autonome conseillé

Télécharger **[Demarrer_Installation_Cabinet.cmd](../../Installateur/Demarrer_Installation_Cabinet.cmd)** avec le bouton de téléchargement du fichier GitHub, ou utiliser la pièce jointe fournie dans la conversation. Un seul fichier suffit. Pour U2b, le lancer exclusivement dans le compte Windows dédié à la recette du PC secondaire, jamais dans le profil clinique U0, et sans élévation administrateur.

Le lanceur ouvre le navigateur pour télécharger une version précise du dépôt privé, repère le ZIP dans Téléchargements (ou ouvre un sélecteur), extrait la version, vérifie les SHA-256 et débloque ses fichiers. Votre navigateur doit être connecté à GitHub avec le compte autorisé ; en cas de page 404, se connecter puis rouvrir le lien affiché. Aucun jeton GitHub n'est demandé. Git et Python ne sont pas requis sur le PC.

Après les trois profils, il demande le **dossier réel des données sur le NAS** : saisir un chemin UNC complet, ou **B** pour parcourir le réseau. Le chemin précédemment enregistré ou le contenu de `chemin.txt` est proposé ; aucun nom de serveur ni de partage n'est imposé. Pour une installation neuve, il peut créer un sous-dossier dans un partage existant après saisie de **CREER**. Cette création ne recherche, ne déplace et n'importe aucune base existante. Choisir un dossier contenant le code du dépôt est refusé.

Le même dossier physique doit être accessible depuis les deux postes. `home` étant personnel au compte NAS, son utilisation demande une vérification explicite de l'accès commun. Les droits d'écriture sont testés par un fichier temporaire unique supprimé immédiatement. Le dossier choisi est mémorisé ; un changement de racine invalide la recette précédente.

Il affiche les trois profils, contrôle le dossier NAS choisi, autorise temporairement l'accès au projet VBA si les stratégies du poste le permettent, construit les fichiers, ouvre les projets à compiler puis le guide de recette. Après vos confirmations, il valide les mêmes fichiers et les active avec sauvegarde. La configuration d'accès VBA initiale est restaurée. Il n'active pas globalement les macros.

**Deux interventions restent nécessaires :** dans chaque éditeur Office ouvert, choisir **Débogage > Compiler**, puis confirmer le résultat dans la console ; réaliser les essais requis et saisir **RECETTE** seulement lorsqu'ils ont réussi. Le lanceur ne déclare pas ces essais réussis à votre place. Saisir **PAUSE** pour conserver la préparation. Relancer le même fichier avec le même profil reprend le parcours ; un binaire modifié invalide ses confirmations.

Le service Synology doit être configuré pour ce même emplacement de données (volume monté côté serveur et chemin UNC côté Windows) ; un choix de dossier dans Windows ne reconfigure pas le serveur. La liaison privée sécurisée entre le domicile et le cabinet doit déjà être configurée : ce lanceur installe les clients Windows, pas le serveur DSM. L'adresse HTTPS et le secret applicatif du poste seront demandés à l'activation. Le déploiement NAS, les réglages Dragon/ECG et le calage de l'imprimante restent décrits ci-dessous.

Les fichiers téléchargés restent dans `%LOCALAPPDATA%\CabinetCardio\Installation\Sources`. L'état de reprise est dans `%APPDATA%\CabinetCardio\Assistant`. Après un arrêt brutal, relancer le fichier pour restaurer aussi le réglage Office temporaire. Un paquet altéré ou contenant des fichiers locaux supplémentaires est automatiquement déplacé dans un dossier de quarantaine conservé. Le téléchargement peut reprendre ; les notes locales restent dans cette copie. Ne pas modifier les empreintes du lanceur.

## Installation depuis le dépôt complet

1. Déployer le service selon [INSTALLATION_NAS.md](Serveur/INSTALLATION_NAS.md), avec le projet `cabinetcardio-test-u2`, le port `8766`, le partage d'exemple `\\NAS-RECETTE\CabinetCardioTestU2` et des volumes propres à U2b. N'utiliser que des données fictives ou une copie isolée.
2. Depuis le dépôt privé `Mandagoutolivier/chatgpt`, télécharger le ZIP, l'extraire sur le PC, puis ouvrir `versions/cabinet-2026.09.12`. Conserver toute son arborescence.
3. Fermer Word et Excel. Word et Excel doivent être installés. Le constructeur nécessite l'accès au modèle objet du projet VBA, à autoriser dans le centre de gestion de la confidentialité pendant la construction. Le script ne désactive pas les protections Office/Windows. Si une stratégie d'entreprise interdit les scripts/macros, faire signer ou autoriser le paquet selon cette stratégie.
4. À domicile, connecter la liaison privée sécurisée du cabinet et vérifier l'accès au partage UNC et à l'adresse HTTPS du service.

## Choisir le poste

Double-cliquer `Installer.cmd` ouvre maintenant le même assistant complet : **1 domicile, 2 secrétariat, 3 cabinet médecin**. Les anciens compléments restent actifs jusqu'à l'activation. Les commandes PowerShell ci-dessous conservent le mode manuel de préparation.

Équivalents PowerShell, depuis le dossier de version :

```powershell
.\Installer.ps1 -Profil Domicile
.\Installer.ps1 -Profil Secretariat
.\Installer.ps1 -Profil Cabinet
```

Word et Excel sont contrôlés dans les trois profils. Le poste médecin reçoit `CabinetUnifie.dotm` ; le tri des arrivées s’effectue en mémoire. Le secrétariat reçoit `Cabinet.xlsm` et un raccourci. `Normal.dotm` est conservé.

## Tester puis activer le même dossier

Compiler les projets VBA Word et Excel préparés avec **Débogage > Compiler**, enregistrer, fermer et rouvrir. Dérouler [RECETTE_WINDOWS.md](RECETTE_WINDOWS.md). La préparation n'est pas une preuve de compilation. Exemple pour domicile :

Exécuter ensuite `Build\Tester_U2_Office.ps1`. La recette automatisée n'est valide que si le socle Word réussit au moins **34 contrôles**, l'extension Word U2 exactement **29/29**, Excel exactement **70/70**, et qu'aucun échec ni essai manquant n'est signalé.

```powershell
.\Build\valider_preparation.ps1 -DossierPrepare 'C:\chemin\du\dossier\prepare' -CompilationWordValidee -CompilationExcelValidee -RecetteValidee
.\Installer.ps1 -Profil Domicile -Mode Installation -DossierPrepare 'C:\chemin\du\dossier\prepare' -RacineNas '\\NAS-RECETTE\CabinetCardioTestU2' -UrlService 'https://adresse-interne-de-recette'
```

Remplacer les chemins et l'adresse d'exemple. Pour secrétariat, utiliser `-Profil Secretariat` et `-CompilationExcelValidee`. Pour médecin, `-Profil Cabinet` et `-CompilationWordValidee`. Le validateur rouvre les binaires, compare leur source au manifeste et vérifie les références Office. La compilation et les essais physiques restent attestés par l'opérateur ; ils ne sont pas simulés.

À l'activation, le script demande le secret applicatif du compte par saisie masquée, ou accepte `-FichierJeton` pointant vers un fichier local protégé. Il contrôle HTTPS, le protocole 2, la révision serveur `2026.09.16-u2b`, le schéma 2 et les rôles. Il conserve ce secret dans `%APPDATA%\CabinetCardio\service.token` avec des droits limités au compte Windows, SYSTEM et administrateurs. Ne jamais le saisir dans la ligne de commande ni le déposer dans GitHub.

Le service doit avoir été initialisé et les classeurs importés avant cette activation. Les ressources initiales sont copiées seulement si absentes ; les bases Excel historiques ne sont jamais remplacées par l'installateur.

## Réglages du poste médecin

Le dossier de recette ECG se règle avec `-DossierGdt 'C:\CabinetCardioTestU2\GDT'`. Il doit être réservé à U2b et ne jamais être surveillé par le profil Resting12Lead clinique. Si aucun profil ECG de test entièrement distinct n'est disponible, vérifier seulement la génération du fichier fictif `IMPORT.GDT`, sans l'importer dans Resting12Lead. Ne jamais utiliser `C:\ECG\GDT`, `C:\Mandagout` ou un autre dossier clinique pour cette recette. Le logiciel ne suppose pas un accès documenté à sa base interne propriétaire.

Affecter les touches Dragon/PowerMic aux macros `Unifie_A_NouvelleLettre`, `Unifie_B_FormuleAppel`, `Unifie_C_InsererPatient`, `Unifie_D_Finaliser`. La clé OpenAI demeure locale au poste médecin, via `OPENAI_API_KEY` ou `%APPDATA%\CabinetCardio\openai.key`.

## Feuille de soins et comptabilité

Les boutons ajoutés à l'accueil donnent accès à l'assuré distinct, aux paramètres des destinataires et à l'encaissement d'une séance. Les réglages inhabituels restent hors de la saisie courante.

La section `[SORTIE]` doit préciser `ExportActif`, `Dossier` et `NomFichier`. Commencer la recette avec `ExportActif=0`, `Dossier=Sorties` et `NomFichier=PublicationID`. L'absence d'une clé est une erreur explicite. `IdentitePublication` ajoute l'identité au nom du fichier et ne doit être choisi qu'après validation du besoin.

Compléter RPPS et numéro AM dans la configuration médecin. Les positions du formulaire sont dans `Config\cerfa_positions.txt` sur le NAS. Les champs `PATIENT_NOM`, `PATIENT_DDN` sont requis pour un assuré distinct ; `MEDECIN_RPPS`, `MEDECIN_AM` sont requis si le praticien n'est pas préimprimé. Aucune coordonnée physique inconnue n'a été inventée. Ajouter les positions après essai sur le formulaire utilisé.

Le calage et `poste.ini` sont **locaux à chaque imprimante/poste**. Après essai papier, renseigner :

```ini
[CERFA]
CalageValide=1
PraticienPreimprime=1
Imprimante=nom exact de l'imprimante
```

Mettre `PraticienPreimprime=1` seulement si les identifiants du médecin figurent déjà correctement sur les feuilles. Un NIR invalide ou une fiche assuré incomplète bloque l'impression ; l'absence de NIR ne bloque pas la consultation. Plus de quatre actes est refusé explicitement et demande une seconde feuille. Aucun connecteur FSE/CPS/SESAM-Vitale n'est livré.

Une impression demandée reste de résultat inconnu jusqu'à confirmation de la feuille sortie. Après interruption, vérifier le papier : confirmer la feuille existante ou demander explicitement une réimpression. Aucun acquittement du courrier n'est accepté tant que ce résultat reste inconnu.

Le journal exporté dans Excel est une copie ; le service conserve les écritures originales. Les montants sont numériques à l'export. L'encaissement couvre un règlement intégral ; le suivi de paiements partiels n'est pas implémenté.

## Retour arrière

La recette sur le PC de l'étage ne doit ni désinstaller ni remplacer l'installation du poste clinique utilisé le lendemain. Chaque activation écrit un dossier `%APPDATA%\CabinetCardio\Sauvegardes\...`. Simulation, puis restauration locale :

```powershell
.\Build\restaurer_poste.ps1 -DossierSauvegarde 'C:\chemin\sauvegarde'
.\Build\restaurer_poste.ps1 -DossierSauvegarde 'C:\chemin\sauvegarde' -Appliquer
```

Ce retour restaure les fichiers du poste d'essai, pas les transactions du NAS. Pour abandonner la recette, restaurer ce PC si nécessaire puis arrêter uniquement `cabinetcardio-test-u2`. Ne pas arrêter le projet clinique, ne pas réutiliser ses volumes et ne pas faire écrire simultanément U2b et l'ancien logiciel dans les mêmes données.
