# U1 — préparation et qualification

Branche de travail issue de la livraison U0 `10dcccbcee2ad926a6dbcc6235f2ad7d2095e84f`.

Le travail porte sur le contrat NAS et la suppression de SQLite dans le flux actif, la migration ACTES, la reprise des courriers, les étapes de facturation et d'impression, l'agenda et la robustesse de l'installation.

Cette livraison est en cours de qualification. Elle ne doit pas être activée au cabinet.

## AX8_Max

Les essais utilisent uniquement `%LOCALAPPDATA%\CabinetCardioTest\U1-20260914`. Une sauvegarde avec empreintes a été faite avant U1 : cinq fichiers de configuration et quarante fichiers Word. Les modèles actifs ne sont pas remplacés et aucune migration du NAS existant n'est exécutée.

Le modèle Word et le classeur Excel de test sont construits. La compilation Word a réussi sur l'essai 4. Le test d'exécution a révélé une erreur dans le calcul d'empreinte Windows ; elle reste à corriger dans ce point de sauvegarde.

## Validation

Avant ce point de sauvegarde, les tests locaux du serveur ont donné 75 réussites et un test de concurrence réservé à PostgreSQL natif. Les suites PowerShell et l'analyse statique ont été exercées. Les résultats GitHub sur ce commit et les nouvelles recettes Office doivent être consultés avant de conclure.

Le test de restauration a été étendu au schéma 2, avec migration ACTES réelle sur données fictives et vérification de l'historique des règlements.

## Conditions avant une activation ultérieure

Compiler Word et Excel, réussir les recettes fictives sur AX8_Max et les tests PostgreSQL natifs, comparer l'installation active à la sauvegarde, puis qualifier le flux NAS/SMB et le calage CERFA avec des données fictives. Les comptes rendus patients actuels, l'ECG, les accès DSM et la configuration de production ne font pas partie des mutations de cette préparation U1.
