# Suites de l'audit — livraison 2026.09.12

Cette livraison part du commit privé `44ff5bb74a20f93a81b81de0d61d75c8370aef50`. L'ancien dossier `cabinet-unifie` conserve l'audit précédent. Les corrections nouvelles sont isolées dans `versions/cabinet-2026.09.12`.

**« Implémenté » signifie présent dans les sources et soumis aux contrôles indiqués. Ce n'est pas une déclaration de recette clinique ou matérielle achevée.** Les originaux Office sont conservés, les fichiers corrigés sont construits sur Windows à partir du manifeste.

## Correspondance avec les recommandations restantes

| Recommandation de l'audit | Suite apportée | État / limite |
|---|---|---|
| Transaction commune au parcours | Service PostgreSQL, transitions sous transaction, commandes idempotentes, publications immuables | Implémenté ; fichiers non référencés possibles après échec SQL, détectés par rapprochement |
| Unifier les annuaires et supprimer le rapprochement par nom seul | Migration par ID, alias d'examen, destinataire principal explicitement confirmé, résolution ambiguë refusée | Implémenté ; anciennes données ambiguës à corriger après simulation |
| Contrôler le contenu généré et imposer relecture | Schéma JSON strict, refus de sorties incomplètes, comparaison nombres/négations/marqueurs, D en deux étapes | Implémenté ; sens médical et anonymisation exhaustive non garantis |
| Patient / assuré, NIR, identifiants praticien | Champs assuré, saisie occasionnelle, contrôle clé NIR, RPPS/AM, positions requises, calage local | Implémenté ; positions du formulaire et impression à vérifier au cabinet ; quatre lignes maximum conservé |
| Limiter copies de bases et latence VPN | Lecture par ID, recherche et agenda filtrés, journal paginé, cache local des arrivées uniquement | Implémenté ; mesure des temps avec le volume réel à effectuer |
| Remplacer protocoles de verrous multi-fichiers | Verrou PostgreSQL, révision optimiste des fiches, réservation serveur | Implémenté ; anciens clients à arrêter pendant la bascule |
| Schéma API et corpus de régression | `schema_reponse_api.json`, tests des règles et fixtures JSON/GDT VBA | Implémenté ; corpus de lettres du cabinet à enrichir avec cas fictifs après recette |
| Retirer code mort et renommer modules | 47 procédures privées supprimées, 9 composants retirés, 3 modules renommés | Nettoyage effectué ; compatibilités publiques conservées pour les commandes externes non recensées |
| Remplacer les écritures directes dans les classeurs | Patients, agenda, actes, annuaire et dictionnaires passent par le service | Implémenté ; fichiers initiaux conservés comme sources de migration |
| Comptes, rôles et journal central | Jetons individuels hachés, permissions serveur, journal de transitions sans texte médical | Implémenté ; pas de SSO/MFA ni de journal inviolable |
| Sauvegarde et restauration | Dump + fichiers + empreintes, restauration d'essai séparée, restauration vers volumes neufs, rapprochement | Scripts livrés ; exécution sur DSM et essai de reprise à effectuer |
| Séparer métier et Office et automatiser les tests | Règles Python, tests transactionnels, CI PostgreSQL natif ; contrôles PowerShell et validation de préparation Office | Tests locaux exécutés ; compilation et périphériques Windows non disponibles ici |
| Activer les binaires réellement vérifiés | Reçu SHA-256 des sources/binaires, comparaison du code rouvert dans Office, activation sans reconstruction | Implémenté et testé côté empreintes ; partie COM à exécuter sur Windows |

Les défauts déjà corrigés par l'audit précédent sont repris : import Unicode des sources, fusion OOXML sans suppression des relations, manifeste complet, conflits de macros, erreurs API propagées, sauvegardes avant activation, identité ECG contrôlée, dates strictes, facturation idempotente et réimpression des actes enregistrés. La réécriture de `modBaseIO` remplace les mécanismes Excel de verrouillage ; elle conserve ces invariants dans PostgreSQL.

## Contrôles réellement exécutés lors de la livraison

| Vérification | Résultat local |
|---|---|
| Sources VBA déclarées, doublons publics, appels qualifiés, équilibre des procédures | 66 fichiers VBA, 15 399 lignes uniques, 73 composants ; aucun défaut détecté par le contrôle statique ; détails/empreintes dans `Tests/inventaire_sources.json` |
| PowerShell 7.6.6 sous Linux, parsing des scripts et conservation OOXML | 34 contrôles réussis |
| Préparation/activation : profil, source modifiée, binaire modifié, dépendance absente | 5 contrôles réussis |
| Cache SQLite | 4 tests réussis |
| Python : règles, API, transactions, migration, conservation des fichiers | 36 tests réussis avec moteur PostgreSQL PGlite ; 1 test de concurrence réservé à PostgreSQL natif |
| Ressources de migration fournies | Simulation sans erreur ; aucune fiche patient dans le jeu initial ; un alias ancien ambigu signalé |
| CI GitHub PostgreSQL 17.11 natif, commit `08db5da` | **37 tests serveur réussis**, dont concurrence ; 4 tests SQLite et 39 contrôles PowerShell réussis — [exécution 34677038738](https://github.com/Mandagoutolivier/chatgpt/actions/runs/34677038738) |
| Compilation/réouverture Office, PowerShell Windows 5.1, NAS réel, Dragon, ECG, impression | Non exécutés ici |

PGlite expose PostgreSQL via un adaptateur socket, avec une connexion interne multiplexée : il ne démontre pas la concurrence d'un serveur natif. Le test concerné est exécuté dans le workflow GitHub `.github/workflows/cabinet-2026-09-12.yml` avec PostgreSQL 17.11. La preuve de son résultat est l'exécution GitHub liée au commit, pas la simple présence du workflow. Les avertissements de dépréciation du client de test HTTP n'affectent pas les assertions mais restent à suivre lors de la prochaine mise à jour des dépendances.

Les contrôles statiques ne compilent pas VBA et ne valident pas les dessins de formulaires. La recette Windows et la vérification du code rouvert dans Office sont exigées avant activation. Aucun fichier patient réel, compte NAS ou secret du cabinet n'a servi aux tests.

## Ce qui demande encore une intervention au cabinet

- Configuration DSM, volumes et droits du compte de service, certificat et VPN.
- Import après arrêt des anciens clients et contrôle des avertissements ; aucune migration de votre NAS n'a été exécutée à distance.
- Compilation et recette Word/Excel/Dragon, correspondance des signets et raccourcis.
- Paramétrage GDT de Resting12Lead, vérification sur identité fictive.
- Calage de la feuille papier, cas assuré distinct et identification du praticien.
- Essai de sauvegarde/restauration et coupures réseau sur le matériel réel.

Le support FSE/CPS/SESAM-Vitale, les paiements partiels, une intégration propriétaire SQL directe dans Resting12Lead et un transfert automatique des réservations entre comptes ne sont pas inclus. Le présent audit ne leur attribue pas une disponibilité fictive.

## Reproduire les tests

Depuis la racine du dépôt :

```sh
python -m pip install -r versions/cabinet-2026.09.12/Serveur/requirements.txt
PYTHONPATH=versions/cabinet-2026.09.12/Serveur CABINET_TEST_DATABASE_URL=postgresql://compte:motdepasse@localhost/base_essai python -m pytest -q versions/cabinet-2026.09.12/Serveur/tests
python versions/cabinet-2026.09.12/Tests/audit_statique.py
python versions/cabinet-2026.09.12/Tests/test_sqlite.py
pwsh -NoProfile -File versions/cabinet-2026.09.12/Tests/test_construction.ps1
pwsh -NoProfile -File versions/cabinet-2026.09.12/Tests/test_installation.ps1
```

Utiliser une base d'essai. Chaque test transactionnel crée et supprime son schéma `cabinet_test_<UUID>`. Sans `CABINET_TEST_DATABASE_URL`, les tests transactionnels sont sautés : ce résultat ne constitue pas une validation du serveur.
