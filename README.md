# Cabinet Cardio — version NAS unique

La seule version active est [cabinet-2026.09.12](versions/cabinet-2026.09.12/README.md), avec PostgreSQL comme autorité des données. L’ancienne livraison `cabinet-unifie` est retirée ; ses points d’entrée échouent avant toute écriture. Son code reste dans l’historique Git.

Les sources U0 sont destinées à une **recette isolée**. La clôture U0 nécessite les preuves NAS, Office et Resting12Lead décrites dans [U0_RECETTE.md](versions/cabinet-2026.09.12/U0_RECETTE.md). Les modèles dans `ModelesSource` sont des entrées de construction, pas des compléments corrigés prêts à activer.

Lire le [guide U0](versions/cabinet-2026.09.12/U0_RECETTE.md) avant une migration ou une activation. Le [lanceur](Installateur/Demarrer_Installation_Cabinet.cmd) vise un commit immuable et vérifie les octets de ce commit. Il prépare les modèles et exige toujours la compilation et la recette sur le poste cible.

Les anciennes copies déjà présentes sur les PC ne peuvent pas être neutralisées par une modification GitHub. Les inventorier et retirer aux comptes clients l’écriture SMB sur les bases historiques avant bascule.

Ne jamais déposer de données patients, courriers ou secrets dans ce dépôt.
