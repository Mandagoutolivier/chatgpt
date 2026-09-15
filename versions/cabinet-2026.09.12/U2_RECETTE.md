# U2 — développement hors poste, 15 septembre 2026

Base exacte : U1 `80b471128113100cd47c6a234fd3ed1ae82e85c8`, PR nº 3. Branche de travail : `codex/urgence-2-20260915`.

**Cette livraison prépare U2 sans accéder à AX8_Max. Aucune compilation ni recette Office U2, activation clinique, migration NAS, rotation réelle de jeton ou suppression réelle de données n'est déclarée effectuée. L'accès au poste attend le feu vert d'Olivier.**

## Changements

| Sujet de l'audit | Réalisation hors poste | Validation encore nécessaire sur poste |
|---|---|---|
| Cycle de courrier | Orchestration dans `modCycleCourrier`, document source privé, état de correction centralisé ; macro `PR_CorrigerToutEnUnClic` conservée comme point d'entrée. Nettoyage des copies de texte à la fin du cycle. | Compilation, alternance entre deux courriers fictifs, erreur/reprise et touches PowerMic. |
| Transport et utilitaires | JSON, UTF-8 et SHA-256 déplacés dans `modDonneesTransport` ; fonctions historiques conservées comme adaptateurs. Commandes locales isolées dans `modCommandesLocales`. | Réexécuter les tests Word U1 qui exercent réellement ces fonctions Windows. |
| Reprise et jetons | Clé de commande basée sur serveur, compte, opération et contenu. Le jeton n'entre plus dans cette clé. Migration locale de l'ancien nom si trouvé ; deux identifiants contradictoires bloquent en conservant les fichiers. | Dix contrôles VBA U2 préparés, dont migration du fichier et conservation de l'identifiant après renouvellement simulé. |
| Code désactivé | Trois anciens modules déjà exclus du manifeste U0 déplacés vers `Archives/Vba`. Les anciens corps `CorrigerDocument` et `GenererDepuisProfil`, sans appel trouvé dans les sources actives, sont retirés ; leurs signatures renvoient un message vers le parcours unifié. | Recensement des appels externes installés dans Dragon ; aucune suppression aveugle des noms de macros. |
| Production et recette | Modules de recette dans `Tests/Vba`, importés uniquement avec `-InclureRecette`. Validation de préparation refusant tout composant supplémentaire. Formulaires et modèles source conservés. | Compiler séparément les constructions de production et de recette ; contrôler ruban et formulaires. |
| Exploitation serveur | Simulation puis compaction explicite des anciens résultats de commandes, preuve de déduplication conservée ; renouvellement/révocation de jetons avec audit et contrôles de concurrence. | Exercices fictifs sur NAS de recette et essais de reprise côté clients avant toute opération d'exploitation. |
| Diagnostic et dépendances | Durée, opération autorisée, étape et statut journalisés sans contenu clinique ni jeton. Rotation Docker des logs (3 × 10 Mo). Dépendances de test exclues de l'image API, versions inchangées. | Mesurer ensuite les durées sur le réseau du cabinet avant de décider une optimisation. |

L'inventaire est reproductible avec `Tests/inventaire_architecture.py`. Il distingue production, recette et archives, signale les composants non classés et les références vers les archives. Il ne peut pas découvrir les commandes Dragon présentes sur un PC auquel aucun accès n'a été fait. Les fonctions de normalisation dont les contrats diffèrent restent séparées.

## Vérifications hors poste

- Reconstitution des 190 fichiers de la base U1, avec vérification des blobs, de l'arbre et du commit Git exacts.
- Contrôles statiques de production et de recette : aucun défaut détecté ; ces contrôles ne compilent pas VBA.
- Quatre tests d'architecture : cohérence des manifestes, détection d'un composant non classé, refus d'un appel de production vers les tests et traitement des littéraux VBA.
- Première exécution locale : 43 tests serveur réussis, 45 tests réservés à PostgreSQL non exécutés localement. La preuve des tests transactionnels est l'exécution GitHub avec PostgreSQL natif.
- Le workflow contrôle aussi la construction de l'image sans dépendances de test, sa lecture sous deux UID sans root, la sauvegarde/restauration, PowerShell Windows et le lanceur généré depuis les objets du commit source.

Les résultats définitifs sont les contrôles associés au commit de la PR U2 ; un test sauté ou la seule présence d'un script ne constitue pas une réussite. Les 34 tests Word et 10 tests Excel d'U1 doivent être rejoués sur les nouveaux composants, en plus des dix tests U2.

## Prochaine étape demandant l'accès à AX8_Max

Après feu vert explicite, et pendant un créneau sans consultation ni document Office ouvert : nouvel inventaire et sauvegarde de l'état courant (sans rétablir une ancienne sauvegarde par-dessus du travail récent), transfert des sources par empreintes, puis `Build/Tester_U2_Office.ps1 -Sortie DOSSIER_NEUF_DE_RECETTE`.

Ce script utilise les protections U1 de Normal.dotm et AccessVBOM. La validation de la construction de production reste distincte de celle des modèles contenant les tests. Les comparaisons finales doivent partir de la sauvegarde de ce nouveau créneau.

La décision de mise en service reste ultérieure : parcours complet NAS/SMB, panne réseau, impression CERFA, ECG et commandes vocales à qualifier sur données fictives. Les améliorations d'interface et les optimisations de performance non mesurées relèvent d'U3.
