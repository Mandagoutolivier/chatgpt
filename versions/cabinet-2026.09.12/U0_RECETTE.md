# Urgence 0 — livraison et recette

Référence : réanalyse du 14 septembre 2026, base Git `71e72fa4ee22f5595f557cbee0e71cb0c69cfd42`. Révision proposée : `2026.09.14-u0`. Ce dossier contient des sources à construire. **Une validation logicielle ne clôt pas les contrôles sur les postes et le Synology. Aucune migration ni activation du cabinet n'est effectuée par cette proposition.**

## Ce qui est corrigé

- Une seule livraison active. Le contenu ancien de `cabinet-unifie` reste dans Git ; ses anciens scripts sont des arrêts immédiats, y compris lorsqu'ils sont appelés directement. Les anciens lanceurs datés sont neutralisés.
- Le générateur du lanceur lit les objets du commit choisi : une modification locale ou un fichier supplémentaire ne change pas ses empreintes. Chaque préparation conserve les empreintes des sources et des modèles ; l'installation conserve un reçu par profil et vérifie la révision du service, le protocole et le schéma.
- La sauvegarde capture PostgreSQL, les fichiers du partage, les brouillons et la configuration du serveur, y compris ses secrets dans une archive privée. Elle vérifie les empreintes, les références des publications et la stabilité des fichiers. Un jeu incomplet n'a pas de marqueur `TERMINE`.
- La vérification restaure dans un projet Compose distinct, avec cluster, réseau interne et volumes dédiés. Elle supprime uniquement ces volumes à sa sortie, succès ou échec. La restauration persistante refuse une base ou des dossiers non vides, sans arrêter une API active.
- Les chemins documentaires connus sont adaptés au nouveau partage dans les publications, consultations, ressources et résultats de commandes mémorisés. Les identifiants et empreintes documentaires restent inchangés. Un ancien chemin restant dans un champ non prévu bloque la restauration.
- Un ID de correspondant explicite n'autorise plus un autre correspondant par repli. Le contexte de masquage inclut le destinataire connu, le médecin traitant et l'auteur configuré. Les requêtes préparées peuvent être inspectées sur données fictives, sans clé ni envoi API. Les noms courts sont recherchés en mots entiers ; les téléphones internationaux français et emails résiduels sont bloqués.
- Les nouvelles archives restent privées en l'absence de groupe configuré. Avec un groupe de lecteurs dédié : fichiers `0640`, dossier `0750`. Aucun passage général à `0644`. Les droits SMB restent à contrôler depuis chaque compte Windows.

## 1. Inventorier les trois postes avant de modifier une installation

Télécharger la proposition et extraire l'archive. Dans `versions/cabinet-2026.09.12`, exécuter sur domicile, secrétariat et AX8_MAX :

```powershell
.\Build\Inventorier_U0.ps1 -Sortie "$env:USERPROFILE\Desktop\U0-$env:COMPUTERNAME.json" -TesterService
```

Le script relève les modèles du dossier de démarrage Word par défaut, le classeur dans Documents, leurs empreintes, les reçus d'installation disponibles et le contrat du service. Il ne lit ni base patients ni contenu des lettres, ne transmet aucun secret dans le rapport et ne lance pas Office. Si Word utilise un autre dossier de démarrage, fournir `-DossierDemarrageWord 'CHEMIN_REEL'` ; l'inventaire le signale comme point à confirmer. Un ancien modèle sans reçu ne devient pas une version connue par sa seule appellation.

Relever sur DSM, sans publier `.env` ou les secrets : projet et image actifs, `CABINET_DATA_VOLUME`, `CABINET_DB_VOLUME`, `CABINET_BACKUP_VOLUME`, `CABINET_UNC`, propriétaires et groupes, partages des brouillons et archives, tâche de sauvegarde existante. Vérifier que `home` désigne le même dossier pour les comptes concernés. Les chemins du fichier `.env.example` sont des exemples.

**Les copies anciennes déjà téléchargées ne sont pas corrigées à distance par GitHub.** Avant une migration, fermer leurs applications et retirer aux comptes clients l'écriture sur les bases historiques (`Base`, `Actes`, anciens fichiers d'agenda et d'échange concernés). Les comptes du service conservent leurs droits nécessaires. Retirer les anciens raccourcis/modèles après les avoir inventoriés et sauvegardés. Ne pas se contenter du marqueur ou du nom de version dans le dépôt.

## 2. Sauvegarde et restauration technique

Prérequis : Container Manager/Docker Compose avec `up --wait`, espace libre pour une sauvegarde et une restauration supplémentaires. Construire les images avant la fenêtre d'arrêt. Les données PostgreSQL sont sur le stockage local du NAS.

**Suspendre réellement les écritures SMB** : fermer les applications de tous les postes et rendre temporairement les données concernées non modifiables par leurs comptes, selon les ACL/partages DSM. Une simple déclaration ne change aucun droit. Les contrôles avant/après détectent les modifications observées ; ils ne garantissent pas l'absence d'un client ancien encore actif. Ne pas programmer une tâche qui passe ce drapeau alors que les écritures restent possibles.

Depuis le dossier `Serveur` du projet source :

```sh
sh sauvegarder.sh --ecritures-smb-suspendues
sh verifier_sauvegarde.sh cabinet-HORODATAGE-IDENTIFIANT
```

Reprendre le nom exact imprimé par la première commande. La pause API couvre toute la capture ; l'API est redémarrée seulement si elle était active au départ. Le verrou de mutations SQL est également conservé pendant la capture. Ne rétablir les droits SMB qu'après la fin. En cas de coupure brutale laissant `.cabinet-backup.lock`, vérifier l'absence de sauvegarde active avant de retirer ce seul verrou ; ne jamais supprimer un jeu pour le faire paraître réussi.

Un dump ne sauvegarde pas les ACL Synology, les comptes DSM, les certificats ou le reverse proxy. Conserver aussi leur sauvegarde via les mécanismes DSM disponibles et vérifier leur restauration. L'archive `configuration.tar.gz` contient les fichiers du dossier serveur et peut inclure des secrets : la garder dans l'espace administrateur, jamais dans un partage ouvert ou dans GitHub.

Le vérificateur ne publie aucun port PostgreSQL, n'utilise pas les secrets du cabinet et ne crée pas de base de contrôle dans le cluster de production. Sa réussite contrôle les octets et les références, **pas l'ouverture dans Word ni les autorisations SMB**.

## 3. Conserver une restauration pour la recette Windows

Créer un nouveau projet Compose avec un nouveau dossier DB, un nouveau partage vide, un nouveau nom de projet et un nouveau dossier privé vide `CABINET_RESTORE_CONFIG`. Adapter le `.env` du nouveau projet ; `CABINET_BACKUP_VOLUME` désigne la sauvegarde source, `CABINET_DATA_VOLUME` le nouveau partage et `CABINET_UNC` son nouveau chemin UNC. Générer les secrets du nouveau projet conformément au guide NAS. Ne pas copier les anciens chemins de volumes dans la configuration active.

Depuis **ce nouveau projet**, démarrer seulement la base puis restaurer :

```sh
docker compose up -d db
sh restaurer_nas_vide.sh cabinet-HORODATAGE-IDENTIFIANT
sh restaurer_nas_vide.sh cabinet-HORODATAGE-IDENTIFIANT --appliquer
```

La première commande de restauration décrit la cible ; la seconde effectue les contrôles puis restaure. Si elle échoue partiellement, la cible est conservée pour diagnostic et aucune bascule n'est effectuée. Repartir ensuite sur une nouvelle cible vide. La configuration récupérée reste séparée de la configuration active.

Restaurer ou adapter les UID/GID et ACL DSM de la cible. Le script conserve les modes et propriétaires POSIX, ce qui ne suffit pas si les comptes ou groupes du nouveau NAS diffèrent. Vérifier les chemins internes aux documents Office, les paramètres locaux et les liens éventuels : seul le rebasculement des champs de chemins du service est automatisé. N'activer aucun client du cabinet sur cette cible avant sa recette.

## 4. ECG et contenu API : données exclusivement fictives

Construire les modèles depuis cette révision dans un environnement de recette, compiler le VBA dans Office et ouvrir le modèle préparé. Exécuter par `Alt+F8` **`Unifie_U0_EssaisLocaux`**. La macro n'appelle pas `Unifie_A_NouvelleLettre`, ne réserve aucun rendez-vous, ne contacte pas l'API et ne touche pas au dossier surveillé par Resting12Lead. Elle crée un dossier temporaire dont le chemin est affiché.

Inspecter les deux JSON : nom, prénom, nom de naissance, DDN, adresse, téléphone, email et professionnels fictifs doivent être masqués ; les données médicales de l'essai doivent rester présentes. La réinjection et la perte de balise sont testées. Le nom `PAGE` ne doit pas altérer `dopage`. Les JSON empruntent exactement le préparateur de requête utilisé avant l'envoi HTTP. La construction du contexte depuis un document réel et ses correspondants reste à essayer avec deux courriers fictifs distincts. Il s'agit de pseudonymisation : un nom inconnu du contexte ou un récit indirectement identifiant n'est pas garanti détecté.

Les fichiers `ECG_A.gdt` et `ECG_B.gdt` contiennent des IDs longs partageant leurs dix premiers caractères. **Aucun alias ni aucune troncature n'est ajouté tant que Resting12Lead n'a pas démontré cette nécessité.** Sur une installation/base ECG d'essai séparée :

1. Configurer le dossier GDT d'essai ; importer A, contrôler nom accentué, prénom, DDN `25/01/1946`, sexe féminin et ID complet.
2. Importer B, contrôler DDN `17/02/1952`, sexe masculin et l'existence de deux fiches distinctes.
3. Réimporter A : retour sur la fiche A, sans créer une troisième identité ni modifier B.
4. Avec un examen d'essai, vérifier son rattachement, puis répéter après restauration de la base ECG d'essai.

En cas de troncature/collision, ne pas utiliser ce flux sur des patients ; un alias serveur unique et durable fera l'objet du correctif correspondant. Ces quatre essais ne peuvent pas être remplacés par la lecture du fichier GDT.

## 5. Archives : comptes réels, fichiers fictifs

Dans le partage de recette, créer un fichier vide `U0-RECETTE-SEULEMENT.txt`. Publier un DOCX et un PDF fictifs via le service de recette. Configurer un **groupe DSM dédié aux lecteurs d'archives**, puis `CABINET_ARCHIVE_GID` et les droits du dossier `Documents`. Pour l'API, ce GID doit être effectivement accessible. Les archives existantes ne sont pas rendues lisibles automatiquement : régler leurs droits dans le contexte du partage restauré, après avoir vérifié le groupe.

Depuis chaque compte Windows médecin, secrétariat et domicile :

```powershell
.\Build\Tester_Droits_U0.ps1 -RacineRecette '\\NAS-RECETTE\Partage' `
  -Docx '\\NAS-RECETTE\Partage\Documents\EMPREINTE.docx' `
  -Pdf '\\NAS-RECETTE\Partage\Documents\EMPREINTE.pdf' `
  -Sortie "$env:USERPROFILE\Desktop\U0-droits-$env:COMPUTERNAME.json"
```

Le contrôle demande des droits d'ouverture sans écrire ni supprimer : lecture doit être accordée ; écriture, suppression et remplacement via les dossiers parents doivent être refusés. Une erreur autre qu'accès refusé est indéterminée. Ouvrir ensuite les deux fichiers depuis Word et le lecteur PDF, vérifier le patient fictif et le contenu. Vérifier aussi qu'un compte non autorisé ne lit pas les archives. Les droits des fichiers seuls ne suffisent pas si le compte peut supprimer le dossier parent.

## Preuves à conserver et clôture

| Volet | Preuve nécessaire | État avant intervention sur les postes |
|---|---|---|
| Livraison | Commit du lanceur, sources vérifiées, reçus et modèles actifs sur chaque poste | Correctifs fournis ; inventaires à recueillir |
| Sauvegarde | Jeu complet privé, suspension SMB effective, seconde copie indépendante du NAS | À exécuter sur l'installation constatée |
| Restauration | Rapport technique, nouvelle cible, chemins corrects, fichiers ouvrables | Automatisation fournie ; recette NAS/Office à exécuter |
| Identité ECG | A/B/A sans collision, sexe/DDN corrects, examen rattaché après restauration | À réaliser dans Resting12Lead |
| API | JSON fictifs inspectés, essais de documents distincts, compilation VBA | Macro fournie ; exécution Office à réaliser |
| Permissions | Résultat par compte, lecture autorisée et modification impossible, refus hors rôle | Script fourni ; droits DSM à vérifier |

Fixer avec l'organisation du cabinet la **perte maximale de travail acceptable (RPO)** et le **délai de remise en service (RTO)** ; aucune valeur n'est présumée ici. Après une restauration chronométrée réussie, configurer dans DSM la fréquence des sauvegardes et de leur contrôle, la méthode effective de suspension des écritures, la conservation et la seconde copie. Le programme ne choisit pas des objectifs à votre place et aucune tâche NAS n'est créée depuis cette session.

Les chantiers U1, notamment la boucle SQLite, la reprise des commandes et la facturation, ne sont pas réputés corrigés par cette livraison U0.
