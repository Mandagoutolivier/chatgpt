# Correctifs ciblés après les audits — 2026.09.16-u2a

Cette livraison prolonge U2 (`8339a7dbeccc903b8a1416fa6aa2d05a1e457446`). Les deux audits transmis portaient sur l'ancien commit `71e72fa` ; les points déjà corrigés dans U0/U1/U2 ne sont pas présentés comme des travaux nouveaux.

**Sources destinées à une recette isolée. Aucune activation de l'installation clinique. Les contrôles Office réussis sur U2 ne qualifient pas ces nouveaux changements VBA.**

## Corrections et comportement obtenu

| Point vérifié | Correction |
|---|---|
| Encaissement d'une séance avec tiers payant | Choix explicite Patient/Organisme et montant réellement reçu. Seules les lignes impayées du payeur choisi sont soldées ; les autres restent intactes. Montants décimaux exacts, contrôle de concurrence et audit avant/après conservés. Un paiement partiel ou un montant incohérent est refusé intégralement. |
| Sélection de la séance | Recherche par date, sélection dans une liste avec identité, naissance et identifiant stable, puis relecture de la séance sur le serveur avant confirmation. Plus de saisie manuelle obligatoire du SeanceID. |
| Recherche de patients/correspondants | Même normalisation Unicode des deux côtés ; accents et casse neutralisés. `%`, `_` et `\` sont recherchés littéralement. Pagination après filtrage et lecture par curseur serveur. Ce filtrage peut parcourir la table : aucun gain de performance non mesuré n'est annoncé. |
| Agenda | Action « Remettre PREVU » utilisant la transition serveur existante. Une arrivée déjà annulée reste protégée : créer un nouveau rendez-vous selon le refus serveur. |
| Destinataire et IA | Sélection d'un identifiant actif et validé avant le cycle de correction ; refus sans destinataire, revalidation avant chaque appel et ajout au contexte de masquage. Le cycle mémorise ce destinataire ; un changement demande un nouveau cycle explicite. La confirmation après relecture reste obligatoire. Cela ne garantit pas la détection de tous les noms libres dans une dictée. |
| Orientation CCN | Apostrophes droite et typographiques reconnues ; un antécédent isolé ne déclenche pas une nouvelle orientation. |
| Copie secondaire du courrier | Le parcours de publication applique `[SORTIE] ExportActif=0` par défaut, y compris si la clé manque. Avec `1`, copie sous le seul PublicationID, empreinte SHA-256 vérifiée, fichier temporaire puis renommage, aucun écrasement d'un contenu différent. Les anciennes macros publiques externes sont conservées et leurs éventuels appels Dragon restent à inventorier. |
| Archivage DOCX/PDF | Les deux contenus sont lus et validés avant la première copie ; les mêmes octets sont archivés. Un PDF invalide ne laisse donc plus un DOCX seul. Les chemins UNC sont résolus sans dépendre de la casse ; ambiguïtés et liens sortant du volume sont refusés. |
| Minimisation des publications | Les nouvelles publications conservent l'identité utile au courrier, sans recopier NIR, coordonnées ni données d'assuré. Le contrôle CERFA relit la fiche, et l'identité de facturation est ensuite figée dans la séance pour les réimpressions. Un changement de nom/prénom/naissance depuis le courrier bloque la première facturation. Les historiques existants ne sont pas effacés. |
| Migration historique | Les cellules numériques entières deviennent des identifiants sans `.0` ; les chaînes, notamment leurs zéros initiaux, sont conservées. Collisions, statuts inconnus, patients incomplets et rendez-vous/séances dépendants sont signalés. L'import reste atomique : aucune ligne silencieusement omise, aucun sexe déduit. |
| Adresses et version | Suppression des lignes vides dans les blocs destinataires générés. Métadonnées de livraison et contrôle du service harmonisés sur `2026.09.16-u2a`, schéma cible 2. |

Le protocole reste 2 : les nouveaux paramètres de paiement sont facultatifs pour un ancien client sans tiers payant. Un ancien client tentant un paiement avec tiers payant reçoit un refus explicite, sans écriture. Les nouveaux postes exigent la révision de service indiquée ci-dessus lors de leur validation.

## Validation reproductible

Les données des tests sont fictives. Les régressions ajoutées couvrent les paiements mixtes et leurs reprises, les refus sans écriture, la sélection par date, les accents et caractères littéraux, la minimisation et le figement comptable, les archives, les liens de migration et la protection du rapport.

```sh
PYTHONPATH=Serveur python -m pytest -q Serveur/tests
python Tests/audit_statique.py
python Tests/audit_statique.py --recette
python Tests/inventaire_architecture.py
python Tests/test_u2_architecture.py
```

PostgreSQL natif est nécessaire pour les tests transactionnels (`CABINET_TEST_DATABASE_URL`). Sans lui, pytest les marque comme ignorés : ce résultat ne les valide pas. Les workflows GitHub exécutent aussi les contrôles PowerShell Windows, la construction de l'image et une sauvegarde/restauration dans un projet isolé. Les résultats finaux sont consignés dans la PR.

Les recettes Word ajoutent cinq cas CCN et deux empreintes de fichiers. La recette Excel ajoute les conversions monétaires exactes et sept entrées invalides. Elles sont fournies pour exécution avec `Build/Tester_U2_Office.ps1` sur un créneau autorisé ; elles n'ont pas été exécutées dans Office pour cette livraison.

## Préparation avant mise en service

1. Conserver l'installation active et ses fichiers locaux. Préparer le service et les nouveaux modèles dans une recette séparée, avec sauvegarde vérifiée.
2. Renseigner localement les coordonnées du praticien et les chemins ECG/export. Le modèle `config.ini` livré ne contient plus de coordonnées réelles ; l’initialiseur conserve un `config.ini` NAS existant. La signature des annexes utilise le nom et le prénom configurés. Simuler la migration sur une copie des classeurs : `python -m cabinet.migration --racine /chemin/copie --rapport /chemin/prive/rapport-neuf.json`. Le rapport est neuf, de mode 0600, sans fiches complètes ; il contient néanmoins des identifiants et doit rester privé. Corriger les données sources et relancer la simulation avant d'utiliser `--appliquer --empreinte-validee ...`. Une précision déjà perdue dans une ancienne cellule Excel ne peut pas être reconstruite automatiquement.
3. Recompiler Word et Excel, exécuter les recettes puis vérifier le parcours fictif : annulation du choix du destinataire sans appel IA, destinataire inactif refusé, reprise de correction, relecture et publication, CERFA et paiement Patient puis Organisme, rafraîchissement concurrent et historique des règlements.
4. Choisir explicitement si la copie secondaire est nécessaire. En cas d'activation de `ExportActif=1`, vérifier les droits réels du dossier SMB, la copie identique lors d'une reprise, le refus d'un fichier divergent et le besoin éventuel d'adapter un consommateur à son nom PublicationID. Le document contient toujours les informations du courrier, même si son nom ne porte plus l'identité.
5. Vérifier sur le NAS que les comptes Office lisent les archives mais ne les modifient pas, et tester le parcours imprimante/ECG/Dragon avant décision de bascule.

La validation préalable des deux fichiers n'est pas une transaction commune au système de fichiers et à PostgreSQL : une panne d'E/S ou SQL peut encore laisser une archive non référencée. La réconciliation et la sauvegarde restent nécessaires ; aucune purge automatique n'est ajoutée.

La facturation sans courrier, les acomptes, un relais IA central et une refonte de la gestion des secrets ne font pas partie de ces correctifs. La clé OpenAI et les accès locaux conservent les limites documentées dans U2. Les modifications d'architecture et de performance restent à instruire séparément, à partir de mesures et du besoin métier.
