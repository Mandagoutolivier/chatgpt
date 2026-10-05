# Binaires C5-NP10 et deploiement AX8_MAX / ACCUEIL

Sources C5-NP10 : commit `f2e74f52e145f3e07480ea9bbeac3bda3fafe346`.
Fichiers recuperes le 5 octobre 2026 dans les profils Windows de recette.

| Fichier | Origine | SHA-256 |
| --- | --- | --- |
| CabinetDragonC5NP10.dotm | AX8_MAX, runtime C5-NP10 du profil CabinetU2Medecin | FF183F42D591BE91CEE2146AF2260439157957C3A5D4FD13D9328B31F19C2C7C |
| Cabinet.xlsm | RDC, runtime C3 du profil CabinetU2Test | 454670F1CB00DD2929EA09D9E6AA0243E68C43E6478B279979E836AFB1FECD24 |

C5 ne modifie aucun module Excel : le classeur C3 accompagne le modele Word C5.
Le hash Excel ci-dessus correspond a la copie actuellement sauvegardee, et differe du
hash de son recu initial C3. Le fichier contient la vue de recette sauvegardee :
patient fictif RECETTEC3 ALPHA dans l'agenda et indicateur de neuf courriers en attente.
Il n'est donc pas un classeur vierge. Le modele Word ne contient pas de texte de courrier.
Les deux binaires ne contiennent aucun module modRecette/modAuditTests ; les recherches
de cles OpenAI/GitHub et d'affectations litterales de secrets dans les sources VBA sont negatives.
Aucun fichier de configuration, jeton, cle API ou recu contenant des empreintes de secrets
n'est inclus dans ce dossier.

## Script fourni

`Deploy-C5-NP10-AX8-ACCUEIL.ps1` utilise les deux fichiers de ce dossier par defaut.
Telecharger le dossier complet, puis lancer PowerShell sous `AX8_MAX\olivi` :

```powershell
.\Deploy-C5-NP10-AX8-ACCUEIL.ps1
```

Le mode par defaut controle seulement les fichiers, les empreintes, Office ferme,
l'acces administratif ACCUEIL et les fichiers de configuration des deux utilisateurs.
Il refuse la racine de recette et l'adresse `https://DS224:8444`.
Il ne configure pas le NAS ni les comptes du service, et ne copie pas les secrets de recette.
La configuration clinique PostgreSQL/HTTPS doit donc deja exister sur les deux profils.

Installation des copies et creation des raccourcis :

```powershell
.\Deploy-C5-NP10-AX8-ACCUEIL.ps1 -Execute -ConfirmerBascule DEPLOYER-C5-NP10 `
  -AncienComplementWord 'CHEMIN_EXACT_DU_COMPLEMENT_WORD_ACTUEL'
```

Remplacer le chemin par celui de l'installation active. Le script controle Office
sur ACCUEIL par WMI/DCOM : une absence de droits ou de connectivite provoque un arret.
Il cree des dossiers et raccourcis horodates, verifie les copies, puis sauvegarde
et renomme l'ancien complement Word. Il ne modifie pas RDC, Normal.dotm, les bases,
la configuration Dragon, les imprimantes ni le serveur. Les anciens raccourcis Excel
sont conserves : ne pas utiliser simultanement les deux applications sur les memes donnees.

## Validation et limites

La syntaxe du script principal et du lanceur Word inclus a ete verifiee par le parseur
PowerShell. Le script de deploiement n'a PAS ete execute sur AX8_MAX/ACCUEIL ; ce depot
de fichiers ne constitue pas une validation de bascule clinique ni une release production.
Le script ne verifie pas l'authentification ni la revision du service de production.
Le serveur, la migration des donnees, les sauvegardes NAS et le parcours materiel
Dragon/PowerMic restent des validations distinctes.

Si une copie ou un acces echoue, il peut rester des dossiers/raccourcis de preparation :
le script n'est pas une transaction a deux postes. L'ancien complement Word est renomme
seulement apres verification des deux copies. Le manifeste du paquet NAS indique les chemins.

Retour arriere local : fermer Office, cesser d'utiliser les nouveaux raccourcis,
remettre le fichier `ancienComplementDesactive` a son chemin `ancienComplement`
consigne dans le manifeste, puis reprendre les anciens raccourcis. Ce retour arriere
ne restaure pas les transactions du NAS et ne doit pas faire coexister deux logiciels
ecrivant de maniere incompatible dans les memes donnees.

La publication est faite a la demande de l'utilisateur ; aucune installation clinique
n'a ete executee pendant cette publication.
