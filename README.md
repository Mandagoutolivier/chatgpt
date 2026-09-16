# Cabinet Cardio — correctifs ciblés U2

Le dossier maintenu dans ce dépôt est [cabinet-2026.09.12](versions/cabinet-2026.09.12/README.md), avec PostgreSQL comme autorité des données. L'ancienne livraison `cabinet-unifie` reste dans l'historique Git.

Cette branche ajoute les [correctifs des audits U2](versions/cabinet-2026.09.12/U2_CORRECTIFS_AUDITS.md), révision `2026.09.16-u2b`. Elle sert à une **recette isolée**, sans remplacement de l'installation clinique actuelle. Les modèles de `ModelesSource` sont les originaux de construction.

Lire [U1_RECETTE.md](versions/cabinet-2026.09.12/U1_RECETTE.md) pour les changements, preuves et limites, ainsi que [U0_RECETTE.md](versions/cabinet-2026.09.12/U0_RECETTE.md) pour les conditions NAS, SMB et ECG avant bascule. Le lanceur est épinglé à un commit et ne remplace pas ces validations.

Les copies déjà installées sur les PC ne sont pas modifiées par la publication de cette branche. Aucun patient, courrier ou secret ne doit être déposé dans ce dépôt.

La [qualification antérieure U2](versions/cabinet-2026.09.12/U2_RECETTE.md) comprend 88 tests serveur et 54 contrôles VBA réels réussis dans Office sous Windows. Les constructions de production Word et Excel compilent et leurs sources ont été comparées au manifeste. L'installation active est conservée ; la qualification du parcours complet au cabinet reste nécessaire avant une mise en service.

Les changements VBA de cette branche nécessitent une nouvelle compilation et recette Office. Les résultats antérieurs ne valent pas validation des nouveaux correctifs.
