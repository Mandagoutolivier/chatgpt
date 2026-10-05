# C5-NP10 : binaires et déploiement AX8_MAX / ACCUEIL

Extraire le paquet complet sur AX8_MAX, sous `AX8_MAX\olivi`, puis lancer `Lancer-Deploiement-C5.cmd` pour contrôler les prérequis sans installer.

Les commandes d'installation et de retour arrière figurent dans [LIRE-MOI-DEPLOIEMENT-C5.md](LIRE-MOI-DEPLOIEMENT-C5.md). Le script principal est `Deploy-C5-NP10-AX8-ACCUEIL.ps1` ; les deux scripts auxiliaires doivent rester à ses côtés.

L'installation nécessite un service de production et deux comptes applicatifs déjà configurés. À la vérification du 5 octobre 2026, ACCUEIL avait son chemin NAS, mais pas `service.url`, `service.token` ni `poste.ini`. Aucune configuration de recette n'est copiée. Le script vise AX8_MAX et ACCUEIL exclusivement.

## Binaires conservés

Sources C5-NP10 : `f2e74f52e145f3e07480ea9bbeac3bda3fafe346`. Binaires récupérés le 5 octobre 2026 dans les profils de recette. C5 concerne Word ; le classeur Excel C3 accompagne ce jalon.

| Fichier | Origine | SHA-256 |
| --- | --- | --- |
| `CabinetDragonC5NP10.dotm` | AX8_MAX, runtime C5-NP10, CabinetU2Medecin | `FF183F42D591BE91CEE2146AF2260439157957C3A5D4FD13D9328B31F19C2C7C` |
| `Cabinet.xlsm` | RDC, runtime C3, CabinetU2Test | `454670F1CB00DD2929EA09D9E6AA0243E68C43E6478B279979E836AFB1FECD24` |

Excel contient une vue de recette fictive enregistrée ; son empreinte diffère du reçu de construction initial après sauvegarde de cette vue. Le lanceur déclenche l'actualisation de l'agenda avant utilisation ; consulter les limites dans le guide. Word est installé sous le nom interne `CabinetUnifie.dotm`, sans modification de son contenu.

Les binaires ne contiennent pas les modules de recette `modRecette` / `modAuditTests`. Le modèle Word ne contient pas de texte de courrier. Aucun fichier de configuration, jeton ou clé API n'est inclus dans ce dossier.

## Validation

Contrôle syntaxique des trois scripts et 13 tests locaux de configurations incorrectes et de restauration réussis. Le script sauvegarde le complément existant, vérifie les copies des deux postes, crée les raccourcis et fournit un journal local de retour arrière.

Aucune installation ou ouverture Office de production n'a été exécutée lors de la préparation de ce paquet. Le démarrage sous Windows PowerShell 5.1/Office, les données du serveur et le parcours clinique Dragon/PowerMic doivent être vérifiés sur les postes.
