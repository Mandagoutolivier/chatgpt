# Préparer U2c sur le Synology

## Recette U2c isolée

Ce guide prépare la révision `2026.09.21-u2c` dans un environnement d'essai distinct. Il ne constate ni un déploiement NAS réussi ni une validation Office. L'installation U0 du cabinet reste en service et ne doit être ni remplacée ni arrêtée pendant cette recette.

| Élément | Valeur U2c de recette |
|---|---|
| Projet Compose | `cabinetcardio-test-u2` |
| Port local du NAS | `127.0.0.1:8766` |
| Image de maintenance | `cabinet-maintenance:2026.09.21-u2c` |
| Code | `/volume1/docker/cabinetcardio-test-u2/code/versions/cabinet-2026.09.12` |
| Données | `/volume1/CabinetCardioTestU2` |
| PostgreSQL | `/volume1/docker/cabinetcardio-test-u2/postgres` |
| Sauvegardes | `/volume1/docker/cabinetcardio-test-u2/sauvegardes` |
| Chemin Windows (exemple) | `\\NAS-RECETTE\CabinetCardioTestU2` |

Le projet clinique, son port `8765`, ses volumes, ses secrets et son reverse proxy ne doivent pas être modifiés. PostgreSQL reste sur un volume local du NAS, jamais sur un montage SMB.

## Organisation et prérequis

Déployer d'abord un projet d'essai distinct. Il faut Container Manager/Docker Compose sur le Synology, un dossier local pour PostgreSQL, un partage de documents et un dossier de sauvegarde. Les chemins `/volume1/...` sont des exemples à adapter. Les fichiers de PostgreSQL doivent être sur le volume **local du NAS**, jamais sur un montage SMB.

Le service de recette écoute uniquement `127.0.0.1:8766` sur l'hôte NAS. Configurer dans DSM un second reverse proxy HTTPS, avec un nom distinct, vers ce port. Ne pas remplacer la règle du service clinique. Utiliser un nom résolu uniquement depuis le cabinet et la liaison privée sécurisée. Aucun port PostgreSQL n'est publié. Ne pas ouvrir le service directement sur Internet.

L'image retenue est `postgres:17.11-bookworm`. Son [Dockerfile officiel](https://github.com/docker-library/postgres/blob/master/17/bookworm/Dockerfile) définit l'UID PostgreSQL 999. L'API a un compte de base non superutilisateur et des comptes applicatifs distincts.

## Commandes limitées à la recette

Après avoir copié les sources dans le dossier indiqué, définir cette fonction dans **chaque nouvelle session shell NAS**. Adapter ensemble les deux chemins absolus si le code est rangé ailleurs. Le projet, le fichier Compose et le fichier d'environnement sont désignés explicitement, indépendamment du dossier courant :

```sh
u2b_compose() {
    docker compose --project-name cabinetcardio-test-u2 \
        --file /volume1/docker/cabinetcardio-test-u2/code/versions/cabinet-2026.09.12/Serveur/compose.yaml \
        --env-file /volume1/docker/cabinetcardio-test-u2/code/versions/cabinet-2026.09.12/Serveur/.env \
        "$@"
}
```

Un nom de projet distinct n'isole pas à lui seul les dossiers montés : contrôler la sortie de `u2b_compose config`, notamment si le shell contient déjà des variables `CABINET_*`, qui peuvent prendre le pas sur `.env`. Aucun montage ne doit désigner un emplacement U0, directement ou par un lien symbolique. Le nom `NAS-RECETTE` est un exemple à remplacer par le nom réel du NAS. Identifier séparément chaque poste de recette ; le nom AX8_MAX ne détermine ni son emplacement ni son rôle. RDC peut rester éteint ; aucune étape ne prévoit de le réveiller.

## Mise en place unique

1. Copier le dossier de version complet dans le nouvel emplacement de code de recette indiqué dans le tableau, distinct du code clinique. Ne pas placer les secrets dans un partage accessible aux postes.
2. Copier `Serveur/.env.u2-test.example` en `Serveur/.env`. Vérifier le projet `cabinetcardio-test-u2`, le port `8766`, les trois volumes isolés, le chemin UNC et l'UID/GID. Ne pas copier le `.env` ni les secrets cliniques.
3. Créer les trois répertoires et accorder les droits nécessaires. Réserver le sous-dossier `Documents` à l'écriture de l'API ; les postes du cabinet doivent seulement pouvoir le lire. Ils doivent pouvoir écrire dans `Patients` pour les brouillons. Éviter un partage donnant l'écriture globale aux archives.
4. Générer les secrets dans le seul dossier de code U2c :

```sh
cd /volume1/docker/cabinetcardio-test-u2/code/versions/cabinet-2026.09.12/Serveur
python3 preparer_secrets.py
sudo chown 999:999 secrets/admin_password.txt
sudo chmod 600 secrets/admin_password.txt
sudo chown 999:GID_API_RECETTE secrets/db_password.txt
sudo chmod 640 secrets/db_password.txt
sudo chmod 700 secrets
```

Remplacer `GID_API_RECETTE` par le GID numérique configuré dans le `.env` de recette avant exécution. Le dossier parent des secrets doit appartenir à l'administrateur qui lance Compose. Le fichier DB est lisible par PostgreSQL et par le groupe de l'API ; seul ce secret est monté dans l'API. Le générateur ne remplace jamais un secret existant.

5. Depuis le dossier de version U2c d’un PC Windows disponible, initialiser les fichiers de support du seul partage de recette :

```powershell
.\Build\initialiser_nas.ps1 -RacineNas '\\NAS-RECETTE\CabinetCardioTestU2'
```

Pour une base d'essai, utiliser le partage d'essai, jamais le partage réel. Le fichier `Config\config.ini` doit préciser explicitement :

```ini
[SORTIE]
ExportActif=0
Dossier=Sorties
NomFichier=PublicationID
```

Une clé absente ou une valeur inconnue arrête l'initialisation. `ExportActif=1` ne doit être choisi qu'après vérification des droits du dossier. `NomFichier` vaut `PublicationID` ou `IdentitePublication`.

6. Afficher et examiner d'abord la configuration résolue :

```sh
u2b_compose config
```

Vérifier les chemins de chaque montage, l'image de maintenance et le port `127.0.0.1:8766`. Si un chemin clinique apparaît, corriger la configuration avant de continuer. Démarrer ensuite la seule recette :

```sh
u2b_compose up -d --build
u2b_compose ps
curl -fsS http://127.0.0.1:8766/health
```

Configurer le reverse proxy DSM distinct ; vérifier `https://adresse-de-recette/health`. La réponse attendue contient `status: ok`, `version` et `protocole: 2`. **`/health` n’annonce ni le schéma ni la révision U2c.** Ces deux valeurs sont contrôlées par le RPC authentifié `whoami` lors de la validation de l’installateur, après migration et création d’un compte de recette : `schema: 2` et `revision: 2026.09.21-u2c`. Si `init-db.sh` échoue sur un nouveau volume, corriger la cause et recréer uniquement ce volume neuf, sans données. L'initialisation PostgreSQL ne se rejoue pas automatiquement sur un volume déjà initialisé.

## Migration des données dans la copie isolée

Utiliser des données fictives ou une copie cohérente déjà préparée pour la recette. Arrêter uniquement les clients de cette copie et terminer ou annuler leurs files d'arrivée et de courriers. Les commandes de ce guide n'organisent aucun arrêt des clients U0. La simulation vérifie les identités, dates, références, identifiants et files encore actives.

```sh
u2b_compose exec -T api python -m cabinet.migration
```

Lire le rapport et conserver son empreinte. Les erreurs empêchent l'import ; les avertissements signalent notamment les anciennes clés de destinataire ambiguës, à corriger dans le nouvel annuaire avant de les utiliser. Les ressources initiales livrées ne contiennent aucun patient : elles produisent 259 correspondants, 6 actes, 172 médicaments et 53 expressions. Les tarifs sont repris de vos fichiers et doivent être contrôlés par le cabinet.

```sh
u2b_compose exec -T api python -m cabinet.migration --appliquer --empreinte-validee EMPREINTE_DE_LA_SIMULATION
```

L'import est transactionnel et refusé si la base cible contient déjà des données. Rejouer le même import renvoie « déjà importé ». Les fichiers Excel restent intacts ; les changements ultérieurs passent par les applications et le service. Les séances historiques sont identifiées par année et conservent leurs montants. Cette commande ne fusionne pas deux bases ayant divergé.

Les formes `Specialistes_ParType` enrichissent les spécialistes existants via `ID_Specialiste`. Plusieurs clés d'examen peuvent être des alias de la même fiche. Une ancienne clé partagée par plusieurs personnes reste bloquée ; le choix explicite d'une fiche utilise son ID unique.

## Mise à niveau du schéma 2 pour U2c

Une base nouvellement initialisée commence au schéma 1. Une copie restaurée peut également nécessiter cette migration. D'abord sauvegarder et restaurer une copie isolée ; arrêter les clients de cette copie pendant la mise à niveau. Déployer U2c sur cette copie, puis simuler :

```sh
u2b_compose exec -T api python -m cabinet.migration_u1
```

Lire les conflits et conserver l'empreinte du plan. Lorsque le plan est approuvé :

```sh
u2b_compose exec -T api python -m cabinet.migration_u1 --appliquer EMPREINTE_DU_PLAN
```

Cette migration transactionnelle conserve les valeurs antérieures dans `migrations_ressources`, unifie `Libelle` et `LibelleCourt`, conserve `Depassement` et passe le schéma à 2. Si les données ont changé depuis la simulation, elle refuse l'application. L'installateur U2c exige le schéma 2 et la révision `2026.09.21-u2c`. La sauvegarde et la restauration acceptent les schémas 1 et 2.

Aucune de ces commandes de migration n'est à exécuter sur les données cliniques pendant la recette.

## Comptes et activation des postes de recette

```sh
u2b_compose exec -T api python -m cabinet.admin compte secretariat-recette --roles secretariat
u2b_compose exec -T api python -m cabinet.admin compte medecin-recette --roles medecin
u2b_compose exec -T api python -m cabinet.admin compte domicile-recette --roles medecin secretariat
```

Chaque commande affiche **une fois** le jeton à fournir à l'installateur du poste concerné. Ne pas l'enregistrer dans un transcript, une capture ou un dossier partagé. La base conserve une empreinte SHA-256 du jeton aléatoire, pas sa valeur. Révocation :

```sh
u2b_compose exec -T api python -m cabinet.admin revoquer domicile-recette
```

Les reprises de consultation sont limitées au compte qui a réservé la consultation. Avec des comptes distincts domicile/cabinet, terminer la consultation sur son poste d'origine ; le transfert entre comptes n'est pas automatisé. Un même compte nominatif peut être utilisé sur les deux postes du même médecin si ses rôles conviennent ; ne pas partager son jeton avec le secrétariat.

Préparer les seuls postes de recette selon [INSTALLATION_MULTI_POSTES.md](../INSTALLATION_MULTI_POSTES.md), en vérifiant leur identité et leurs chemins locaux. Ne pas activer U2c dans les modèles ou applications cliniques pendant cette phase. Le passage de tous les postes cliniques relève d’une mise en service ultérieure, après recette complète.

## Sauvegarde et reprise

Les exigences de cohérence sont décrites dans [U0_RECETTE.md](../U0_RECETTE.md) : suspension effective des écritures SMB, capture de la base et des fichiers/configuration, vérification dans un cluster isolé, puis restauration persistante sur une cible vide pour la recette Windows. Pour U2c, appliquer ces exigences aux seuls emplacements de recette ; ne pas reprendre les chemins ni les commandes de production U0 sans les adapter et les vérifier. Les anciens jeux sans manifeste U0 restent conservés, mais ne sont pas acceptés silencieusement par ce nouveau vérificateur. Ne pas fabriquer de marqueur `TERMINE` pour les convertir.

## Limites d'exploitation

Les correctifs U2c ne sont pas qualifiés pour le service clinique tant que la recette réelle n'est pas terminée. Les UID, droits SMB, certificat, reverse proxy, volume libre et restauration physique doivent être contrôlés sur le NAS. Les parcours Word/Excel restent à qualifier dans l’environnement de recette. Les comptes sont des jetons applicatifs ; l'authentification SSO/MFA n'est pas implémentée. Le service conserve les ressources métier en JSONB versionné ; une évolution de schéma exige une migration explicite, pas une modification manuelle des tables.

## Arrêt de la recette et retour à l'installation actuelle

Avec la fonction `u2b_compose` définie plus haut :

```sh
u2b_compose stop
u2b_compose ps
```

Cette commande arrête uniquement `cabinetcardio-test-u2`. Elle laisse le service clinique, son port `8765` et ses données inchangés. Conserver les volumes de recette jusqu'à l'analyse ; leur suppression éventuelle fera l'objet d'une opération séparée et explicite.
