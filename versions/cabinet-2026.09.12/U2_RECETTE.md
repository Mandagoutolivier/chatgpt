# U2 — qualification technique du 15 septembre 2026

Base : U1 `80b471128113100cd47c6a234fd3ed1ae82e85c8`, PR nº 3. Branche : `codex/urgence-2-20260915`, [PR nº 4 en brouillon](https://github.com/Mandagoutolivier/chatgpt/pull/4).

**Les correctifs U2, la compilation des modèles et les contrôles VBA ci-dessous sont réalisés. La qualification du parcours complet reste nécessaire avant une mise en service.**

Le commit exécuté pour la recette Office est **`f0a45ecc6daa7a5cd9413f1d9e769747388a8964`**. La publication de ce bilan ne modifie pas les sources VBA, Python ni PowerShell qualifiées ; le lanceur est régénéré pour inclure la documentation actualisée. Les journaux détaillés restent conservés localement.

## Changements et limites

| Sujet de l'audit | Réalisation U2 | Qualification complémentaire |
|---|---|---|
| Cycle de courrier | Orchestration dans `modCycleCourrier`, document source privé, état centralisé ; entrée historique conservée et copies temporaires de texte libérées en fin de cycle. Compilation réussie. | Alternance entre deux courriers fictifs, erreur/reprise du parcours complet et commandes vocales. |
| Transport et utilitaires | JSON, UTF-8 et SHA-256 isolés dans `modDonneesTransport`, fonctions historiques conservées comme adaptateurs. Les 34 contrôles Word U1 ont été rejoués avec succès. | Parcours client/serveur réel. |
| Reprise et jetons | Clé stable par serveur, compte, opération et contenu. Dix contrôles VBA U2 réussis, dont migration des fichiers U1, renouvellement simulé et conservation des identifiants contradictoires. | Avant le premier renouvellement réel depuis U1, résoudre/migrer les attentes tant que l'ancien jeton est disponible. |
| Code désactivé | Trois modules déjà exclus déplacés dans `Archives/Vba`. Deux corps de moteurs historiques retirés sans appel trouvé dans les sources actives ; signatures conservées avec message explicite. | Recenser les appels externes des commandes vocales. |
| Production et recette | Modules de recette dans `Tests/Vba`, importés uniquement avec `-InclureRecette`. Les constructions de production et de recette compilent séparément. Les 55 composants Word et 25 Excel de production correspondent au manifeste et aux sources. | Vérifier le ruban et les formulaires dans le parcours complet. |
| Exploitation serveur | Simulation puis compaction explicite des copies de résultats, preuve de déduplication conservée ; renouvellement/révocation auditée et contrôles de concurrence testés sur PostgreSQL. | Exercices fictifs sur serveur de recette et reprises côté clients avant toute opération réelle. |
| Diagnostic et dépendances | Durée, opération autorisée, étape et statut sans contenu clinique ni jeton. Rotation des logs (3 × 10 Mo). Dépendances de test exclues de l'image API. | Mesurer les durées avant toute optimisation. |

L'inventaire reproductible `Tests/inventaire_architecture.py` distingue production, recette et archives. Il ne recense pas les commandes vocales externes. Les fonctions de normalisation dont les contrats diffèrent restent séparées.

## Vérifications automatisées

- Contrôles statiques de production et de recette sans erreur ; quatre tests d'architecture réussis.
- **88 tests serveur réussis, aucun sauté**, sur PostgreSQL 17 natif. Deux avertissements de dépréciation des bibliothèques de test.
- Quatre jobs GitHub réussis : métier PostgreSQL, restauration, image serveur et installateur Windows. Contrôles de l'image sous deux UID non root, PowerShell Windows 5.1, séparation des composants et lanceur vérifié contre les objets Git.

Preuves du commit exécuté : [exécution de la branche](https://github.com/Mandagoutolivier/chatgpt/actions/runs/34942315972) et [exécution de la PR](https://github.com/Mandagoutolivier/chatgpt/actions/runs/34942320296). Les contrôles des commits de documentation ultérieurs restent visibles dans la PR.

## Recette Office réelle

| Contrôle | Résultat |
|---|---|
| Construction et compilation des modèles de recette Word/Excel | Réussies. |
| Tests VBA Word U1 rejoués | **34 réussis**, aucune erreur. |
| Tests VBA U2 | **10 réussis**, aucune erreur. |
| Tests VBA Excel U1 rejoués | **10 réussis**, aucune erreur. |
| Construction de production sans modules de recette | Réussie ; compilation Word et Excel réussie. |
| Composants de production | **55 Word + 25 Excel** ; noms, code source et références vérifiés dans Office. |

Ces contrôles utilisent des données fictives. Ils ne constituent ni une validation du sens médical des courriers, ni une qualification du parcours réseau ou de l'impression physique. Les modèles construits restent destinés à la qualification ; les originaux de `ModelesSource` ne doivent pas être installés directement.

Le contrôle supplémentaire de production a révélé qu'Excel pouvait rester en arrière-plan après sa fermeture COM. La fin du processus de contrôle a libéré cette instance ; la restauration de l'environnement a ensuite été vérifiée séparément. Le succès des compilations et la fin du nettoyage sont donc deux constats distincts.

## Avant la mise en service

Qualifier sur données fictives le parcours complet Word → NAS/SMB → Excel, les pannes et reprises réseau, l'impression CERFA, l'ECG, les commandes vocales, les formulaires et les données ACTES. Tester la sauvegarde/restauration du serveur de recette avant une bascule décidée séparément.

Pour chaque nouveau créneau, refaire l'inventaire et la sauvegarde de l'état courant ; ne pas restaurer une ancienne sauvegarde par-dessus les travaux récents.

Les améliorations d'interface et les optimisations de performance non mesurées relèvent d'U3.
