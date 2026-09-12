# Cabinet Cardio — version 2026.09.12, service NAS

Cette version reprend l'audit du commit `44ff5bb` et remplace les écritures concurrentes dans les classeurs par un service transactionnel PostgreSQL sur le Synology. Word reste l'interface de dictée, Excel celle du secrétariat. SQLite est uniquement un cache local des patients arrivés pour le poste médecin.

**Sources et outils livrés ; compilation et recette Office/Dragon/ECG/imprimante à effectuer sur Windows avant utilisation avec des patients.** Les fichiers dans `ModelesSource` sont les originaux de référence, pas les binaires corrigés à installer directement.

## Installation à trois entrées

Ouvrir `Installer.cmd` ou lancer `Installer.ps1` :

| Choix | Profil | Partie installée |
|---|---|---|
| 1 | Domicile | Secrétariat + médecin, accès au NAS par VPN |
| 2 | Secretariat | Patients, agenda, courriers reçus, actes et journal |
| 3 | Cabinet | Poste médecin, Word unifié, cache SQLite et export ECG |

Le service Synology se configure **une fois** avant l'activation des postes. L'installation prépare les binaires, puis active le dossier vérifié sans le reconstruire. Voir [installation des postes](INSTALLATION_MULTI_POSTES.md) et [installation NAS](Serveur/INSTALLATION_NAS.md).

## Parcours

Saisie unique au secrétariat : nom, prénom, date de naissance, **sexe explicite pour l'ECG**, adresse, téléphone et médecin traitant. Une arrivée ouvre une consultation identifiée ; un seul compte médecin peut la réserver.

- **A** : choisir un patient arrivé, préparer le courrier et l'identité ECG, puis dicter le raccourci Dragon du destinataire.
- **B** : formule d'appel et corps du texte.
- **C** : insérer nom, prénom et âge.
- **D** : correction API, lettres annexes et gras Cabinet. **Relire, puis D une seconde fois** pour confirmer le destinataire et transmettre au secrétariat.

Le destinataire est confirmé par identifiant ; les homonymes ou anciennes clés ambiguës ne sont pas départagés automatiquement. Les courriers publiés sont conservés sous empreinte. Une nouvelle version du courrier ne crée pas une seconde séance comptable.

## Documents de livraison

- [Corrections et limites](AUDIT.md) : correspondance avec les recommandations précédentes.
- [Architecture et contrats](INTEGRATION_UNIFIEE.md).
- [Recette Windows](RECETTE_WINDOWS.md) : essais attendus et validation de la préparation.
- [Service NAS](Serveur/INSTALLATION_NAS.md) : comptes, migration, sauvegarde et restauration.
- [Inventaire des sources](Tests/inventaire_sources.json).

La version précise R11/R12/R13 de `PROD(6)` demeure **non démontrée**. L'analyse d'empreinte reste disponible dans l'[audit précédent](../../cabinet-unifie/AUDIT.md) ; le présent numéro de livraison ne prétend pas résoudre cette filiation.
