# Déploiement C5-NP10 : AX8_MAX et ACCUEIL

Ce paquet installe les clients Office depuis **AX8_MAX**, sous **AX8_MAX\olivi** : Word pour le médecin et Excel pour **ACCUEIL\accueil**. Il remplace le précédent script de déploiement. Aucun logiciel ni fichier n'est installé sur RDC.

Le double clic sur `Lancer-Deploiement-C5.cmd` effectue seulement les contrôles. L'installation doit être déclenchée avec les options explicites ci-dessous. Elle ne démarre pas Word ou Excel.

## Prérequis

- Fermer Word et Excel sur AX8_MAX et ACCUEIL, dans toutes les sessions.
- Extraire tous les fichiers du ZIP dans un même dossier local sur AX8_MAX. Ne pas lancer depuis le ZIP.
- Utiliser le compte `AX8_MAX\olivi`, avec les droits de lecture/écriture sur `\\ACCUEIL\C$`, les droits WMI/DCOM sur ACCUEIL et l'accès à `\\DS224\CabinetCardio-Dev`.
- Le service applicatif de production doit déjà être disponible, avec protocole 2 et schéma 2, une révision validée et ses données cliniques. Ce script ne crée ni ne migre le serveur, sa base ou les données du NAS.
- Configurer séparément les deux profils applicatifs. Les deux postes doivent viser le même service HTTPS de production et utiliser deux comptes applicatifs distincts.
- Le certificat HTTPS doit être reconnu sur les deux PC. La sécurité des macros Office doit permettre l'ouverture des binaires ; le script ne modifie aucun réglage Office.

À la dernière vérification d'ACCUEIL, `chemin.txt` était présent, mais `service.url`, `service.token` et `poste.ini` étaient absents. Il faut fournir ces configurations de production avant l'installation. Le script s'arrêtera si elles manquent. Il ne copie pas les jetons ni les configurations de recette.

| Fichier dans `%APPDATA%\CabinetCardio` de chaque utilisateur | Contenu requis |
| --- | --- |
| `chemin.txt` | `\\DS224\CabinetCardio` |
| `service.url` | URL HTTPS exacte du service de production, sans retour de ligne final ; le port de recette 8444 est refusé |
| `service.token` | Jeton personnel de production, au moins 32 caractères, sans espace ni retour de ligne ; ne pas le publier ni le partager |
| `poste.ini` sur AX8_MAX | Section `[POSTE]`, ligne `Profil=CabinetMedecin` |
| `poste.ini` sur ACCUEIL | Section `[POSTE]`, ligne `Profil=CabinetSecretariat` |

Le contrôle lit les fichiers d'ACCUEIL via `\\ACCUEIL\C$\Users\accueil\AppData\Roaming\CabinetCardio`. Il vérifie l'authentification des deux comptes depuis AX8_MAX. La connexion HTTPS depuis ACCUEIL et l'exécution réelle des macros seront vérifiées par le lanceur au premier démarrage sur ce PC.

## 1. Contrôler sans installer

Dans Windows PowerShell sur AX8_MAX, depuis le dossier extrait :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy-C5-NP10-AX8-ACCUEIL.ps1
```

Le contrôle affiche la révision du service s'il réussit. Il n'écrit aucun fichier d'installation. Il contrôle les empreintes des binaires, les configurations, les rôles, les dossiers partagés et la fermeture d'Office.

Si l'ancien complément Word n'est pas identifié de façon unique dans le dossier STARTUP par défaut, relancer en ajoutant :

```powershell
-AncienComplementWord 'C:\Users\olivi\AppData\Roaming\Microsoft\Word\STARTUP\NOM-EXACT-DU-MODELE.dotm'
```

Remplacer ce chemin par le fichier réellement chargé au démarrage de Word. Le fichier doit se trouver sous l'APPDATA d'olivi et être un complément `Cabinet...` ou `ModeleCourrierChatGPT...` en `.dot` ou `.dotm`.

Si le Bureau d'ACCUEIL est redirigé, ajouter `-BureauAccueil` avec son chemin local réel sur C:, par exemple le dossier Desktop du profil OneDrive. Le défaut est `C:\Users\accueil\Desktop`.

## 2. Installer sur les deux PC

Après un contrôle réussi, reprendre exactement la révision validée affichée et lancer :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy-C5-NP10-AX8-ACCUEIL.ps1 -Execute -ConfirmerBascule DEPLOYER-C5-NP10 -RevisionServiceAttendue 'REVISION_VALIDEE'
```

Remplacer `REVISION_VALIDEE` par la révision attendue du service de production. Reprendre les options de chemin utilisées au contrôle, si nécessaire. Le script recommence les contrôles avant toute installation.

L'installation crée :

- une sauvegarde de l'ancien complément Word et une copie du paquet sous `\\DS224\CabinetCardio-Dev\Deploy\C5-NP10\IDENTIFIANT` ;
- une version locale sous `%LOCALAPPDATA%\CabinetCardio\Versions\C5-NP10-IDENTIFIANT` sur chaque PC ;
- les raccourcis **Cabinet C5 - Medecin - IDENTIFIANT** sur AX8_MAX et **Cabinet C5 - Secretariat - IDENTIFIANT** sur ACCUEIL ;
- un journal local sous `%LOCALAPPDATA%\CabinetCardio\Deploiements\IDENTIFIANT.json` sur AX8_MAX, également copié avec le paquet NAS.

L'ancien complément Word est renommé avec un suffixe `.desactive-IDENTIFIANT` après copie et vérification des fichiers des deux postes. Les anciens raccourcis Excel sont conservés : utiliser le nouveau raccourci pour C5. Aucun jeu de données ni transaction clinique n'est créé par l'installation.

Les commandes PowerShell et les raccourcis utilisent `ExecutionPolicy Bypass` pour leur processus seulement. Aucune stratégie d'exécution permanente ni stratégie de groupe n'est modifiée.

## 3. Premier lancement

Sur chaque PC, ouvrir son nouveau raccourci depuis le profil utilisateur prévu, avec Office fermé. Le lanceur vérifie de nouveau le service et la révision avant l'ouverture d'Office.

Sur AX8_MAX, il prépare le cache des correspondants, charge le complément Word sous le nom interne requis `CabinetUnifie.dotm`, puis initialise les bases, les raccourcis et la file des arrivées. Ce changement de nom conserve les octets et l'empreinte du modèle C5-NP10.

Sur ACCUEIL, il ouvre la copie locale du classeur en écriture, vérifie la racine lue par VBA, déclenche l'actualisation de l'agenda du jour et démarre la scrutation des courriers. La copie Excel d'origine contient une vue de recette fictive ; cette vue doit être remplacée au lancement et ne constitue pas une importation de données de recette dans le service. Si Excel affiche une erreur d'agenda, arrêter ce lancement : la macro peut afficher son erreur sans la renvoyer au lanceur PowerShell. Vérifier visuellement la semaine affichée et les rendez-vous avant utilisation.

Si le démarrage affiche `ECHEC`, ne pas travailler dans une instance partiellement ouverte. Le lanceur ferme l'instance qu'il a créée. Utiliser le retour arrière si nécessaire. La validation clinique et le fonctionnement Dragon/PowerMic restent à vérifier sur les deux postes après installation.

## 4. Retour arrière

Fermer les nouvelles applications sur les deux postes. Sur AX8_MAX, reprendre le **journal local** dont le chemin a été affiché à l'installation :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy-C5-NP10-AX8-ACCUEIL.ps1 -Restaurer 'CHEMIN_DU_JOURNAL_LOCAL.json' -Execute
```

L'ancien complément Word est restauré sans écraser un fichier différent. Les nouveaux raccourcis sont renommés pour les rendre inactifs ; les versions C5 et les sauvegardes sont conservées. Si ACCUEIL est inaccessible, Word est tout de même restauré sur AX8_MAX et le journal signale la neutralisation distante incomplète. Fermer alors les applications sur ACCUEIL et reprendre cette commande lorsque le partage est accessible.

Ce retour arrière concerne les clients. Il n'annule pas les actions cliniques effectuées entre-temps et ne restaure aucune donnée du NAS.

## Provenance et validation

Sources C5-NP10 : commit `f2e74f52e145f3e07480ea9bbeac3bda3fafe346`, dépôt `Mandagoutolivier/chatgpt`, branche `codex/u2-recette-20260920`. Word provient du modèle compilé C5-NP10 ; Excel est le classeur compilé C3 conservé pour ce jalon Word.

| Binaire fourni | SHA-256 |
| --- | --- |
| `CabinetDragonC5NP10.dotm` | `FF183F42D591BE91CEE2146AF2260439157957C3A5D4FD13D9328B31F19C2C7C` |
| `Cabinet.xlsm` | `454670F1CB00DD2929EA09D9E6AA0243E68C43E6478B279979E836AFB1FECD24` |

Les scripts ont passé le contrôle syntaxique PowerShell et 13 tests locaux ciblant le retour arrière et les configurations incorrectes. Ces tests isolés n'exécutent ni Office ni une installation sur les deux PC. Ils ne remplacent pas une validation de démarrage sur Windows PowerShell 5.1 et Office dans les profils de production.
