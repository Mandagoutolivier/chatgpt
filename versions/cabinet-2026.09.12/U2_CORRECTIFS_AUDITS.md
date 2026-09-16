# Correctifs ciblés après les audits — 2026.09.16-u2b

Cette livraison remplace le candidat U2a sans le modifier rétroactivement. Elle prolonge U2 (`8339a7dbeccc903b8a1416fa6aa2d05a1e457446`).

**Sources destinées à une recette isolée. Aucune activation de l'installation clinique.** Le déploiement d'essai utilise le projet Compose `cabinetcardio-test-u2`, le port `8766`, des volumes NAS distincts et l'image `cabinet-maintenance:2026.09.16-u2b`.

## Corrections et comportement obtenu

| Point vérifié | Correction |
|---|---|
| Encaissement d'une séance avec tiers payant | Choix Patient/Organisme pour chaque bundle d'acte ; l'acte associé hérite du principal. Une répartition incomplète est refusée. Seules les lignes impayées du payeur choisi sont soldées ; les autres restent intactes. Montants décimaux exacts, contrôle de concurrence et audit avant/après conservés. |
| Sélection de la séance | Recherche par date, sélection dans une liste avec identité, naissance et identifiant stable, puis relecture de la séance sur le serveur avant confirmation. Plus de saisie manuelle obligatoire du SeanceID. |
| Recherche de patients/correspondants | Même normalisation Unicode des deux côtés ; accents et casse neutralisés. `%`, `_` et `\` sont recherchés littéralement. Pagination après filtrage et lecture par curseur serveur. Ce filtrage peut parcourir la table : aucun gain de performance non mesuré n'est annoncé. |
| Agenda | Action « Remettre PREVU » utilisant la transition serveur existante. Une arrivée déjà annulée reste protégée : créer un nouveau rendez-vous selon le refus serveur. |
| Destinataire et IA | Sélection d'un identifiant actif et validé avant le cycle de correction ; refus sans destinataire, revalidation avant chaque appel et ajout au contexte de masquage. Le cycle mémorise ce destinataire ; un changement demande un nouveau cycle explicite. La confirmation après relecture reste obligatoire. Cela ne garantit pas la détection de tous les noms libres dans une dictée. |
| Orientation CCN | Apostrophes droite et typographiques reconnues ; un antécédent isolé ne déclenche pas une nouvelle orientation. |
| Copie secondaire du courrier | `ExportActif`, `Dossier` et `NomFichier` sont explicites. Une clé absente ou une valeur inconnue produit une erreur au lieu de désactiver silencieusement l'export. Avec `0`, aucune copie. Avec `1`, `PublicationID` reste pseudonymisé ; `IdentitePublication` ajoute une identité normalisée sans tronquer l'identifiant. Empreinte, fichier temporaire et refus d'écrasement divergent sont conservés. |
| Archivage DOCX/PDF | Les deux contenus sont lus et validés avant la première copie ; les mêmes octets sont archivés. Un PDF invalide ne laisse donc plus un DOCX seul. Les chemins UNC sont résolus sans dépendre de la casse ; ambiguïtés et liens sortant du volume sont refusés. |
| Minimisation et réimpression | Les nouvelles publications ne recopient ni NIR, ni coordonnées, ni données d'assuré. La première facturation relit la fiche puis fige identité, assuré, actes, montants et libellés CERFA. Une réimpression relit exclusivement ce cliché, même si la fiche ou la nomenclature est corrigée ensuite ; un autre patient ou un historique incohérent est refusé. |
| Migration historique | Les cellules numériques entières deviennent des identifiants sans `.0` ; les chaînes, notamment leurs zéros initiaux, sont conservées. Collisions, statuts inconnus, patients incomplets et rendez-vous/séances dépendants sont signalés. Les historiques exigent une identité, des règlements et un état d'impression cohérents ; une identité d'assuré partielle est refusée. La fusion de correspondants conserve les restrictions d'inactivité ou de validation. L'import reste atomique : aucune ligne silencieusement omise, aucun sexe déduit. |
| Adresses, déploiement et version | Suppression des lignes vides dans les blocs destinataires générés. Métadonnées harmonisées sur `2026.09.16-u2b`, schéma cible 2. Projet, port, volumes et image de maintenance ont des valeurs de recette distinctes ; leur séparation doit être vérifiée avant tout lancement. Le dossier GDT proposé par les installateurs est `C:\CabinetCardioTestU2\GDT`. |

Le protocole reste 2 : les nouveaux paramètres de paiement sont facultatifs pour un ancien client sans tiers payant. Un ancien client tentant un paiement avec tiers payant reçoit un refus explicite, sans écriture. Les nouveaux postes exigent la révision de service indiquée ci-dessus lors de leur validation.

## Validation reproductible

Les données des tests sont fictives. Les régressions ajoutées couvrent les paiements mixtes et leurs reprises, les refus sans écriture, la sélection par date, les accents et caractères littéraux, la minimisation et le figement comptable, les archives, les liens de migration et la protection du rapport.

```sh
PYTHONPATH=Serveur python -m pytest -q Serveur/tests
python Tests/audit_statique.py
python Tests/audit_statique.py --recette
python Tests/inventaire_architecture.py
PYTHONPATH=Serveur python Tests/test_u2_architecture.py
```

PostgreSQL natif est nécessaire pour les tests transactionnels (`CABINET_TEST_DATABASE_URL`). Sans lui, pytest les marque comme ignorés : ce résultat ne les valide pas. Les workflows GitHub exécutent aussi les contrôles PowerShell Windows, la construction de l'image et une sauvegarde/restauration dans un projet isolé. Les résultats finaux sont consignés dans la PR.

La recette Office U2b exige le socle Word antérieur (au moins **34 contrôles**), puis exactement **29 essais Word U2** et **70 essais Excel**. `Build/Tester_U2_Office.ps1` refuse un échec, un champ manquant ou un total U2 différent du total attendu. Ces suites doivent encore être compilées et exécutées dans les applications Office réelles sur le PC d'essai.

## Préparation avant mise en service

1. Conserver l'installation active et ses fichiers locaux. Préparer le service et les nouveaux modèles dans une recette séparée, avec sauvegarde vérifiée.
2. Renseigner localement les coordonnées du praticien et les chemins ECG/export. Le modèle `config.ini` livré ne contient plus de coordonnées réelles ; l’initialiseur conserve un `config.ini` NAS existant. La signature des annexes utilise le nom et le prénom configurés. Simuler la migration sur une copie des classeurs : `python -m cabinet.migration --racine /chemin/copie --rapport /chemin/prive/rapport-neuf.json`. Le rapport est neuf, de mode 0600, sans fiches complètes ; il contient néanmoins des identifiants et doit rester privé. Corriger les données sources et relancer la simulation avant d'utiliser `--appliquer --empreinte-validee ...`. Une précision déjà perdue dans une ancienne cellule Excel ne peut pas être reconstruite automatiquement.
3. Recompiler Word et Excel, exécuter les recettes puis vérifier le parcours fictif : annulation du choix du destinataire sans appel IA, destinataire inactif refusé, reprise de correction, relecture et publication, CERFA et paiement Patient puis Organisme, rafraîchissement concurrent et historique des règlements.
4. Définir explicitement `[SORTIE] ExportActif`, `Dossier` et `NomFichier`. Commencer avec `ExportActif=0`. En cas d'activation, vérifier les droits SMB, la reprise identique, le refus d'un fichier divergent et choisir consciemment entre `PublicationID` et `IdentitePublication`.
5. Vérifier sur le NAS que les comptes Office lisent les archives mais ne les modifient pas, et tester le parcours imprimante/ECG/Dragon avant décision de bascule.

La validation préalable des deux fichiers n'est pas une transaction commune au système de fichiers et à PostgreSQL : une panne d'E/S ou SQL peut encore laisser une archive non référencée. La réconciliation et la sauvegarde restent nécessaires ; aucune purge automatique n'est ajoutée.

## Retour arrière de la recette

Arrêter uniquement le projet `cabinetcardio-test-u2` et conserver ses volumes pour analyse. Sur le PC d'essai, restaurer la sauvegarde locale créée par l'installateur si U2b a été activé. Cette procédure ne modifie ni le projet clinique, ni son port `8765`, ni ses données. Toute bascule clinique ultérieure fera l'objet d'une décision et d'une sauvegarde vérifiée séparées.

La facturation sans courrier, les acomptes, un relais IA central et une refonte de la gestion des secrets ne font pas partie de ces correctifs. La clé OpenAI et les accès locaux conservent les limites documentées dans U2. Les modifications d'architecture et de performance restent à instruire séparément, à partir de mesures et du besoin métier.
