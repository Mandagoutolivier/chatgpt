# Déployer le service sur le Synology

## Organisation et prérequis

Déployer d'abord un projet d'essai distinct. Il faut Container Manager/Docker Compose sur le Synology, un dossier local pour PostgreSQL, un partage de documents et un dossier de sauvegarde. Les chemins `/volume1/...` sont des exemples à adapter. Les fichiers de PostgreSQL doivent être sur le volume **local du NAS**, jamais sur un montage SMB.

Le service écoute uniquement `127.0.0.1:8765` sur l'hôte NAS. Configurer dans DSM un reverse proxy HTTPS avec un certificat reconnu par les PC, vers ce port. Utiliser un nom résolu au cabinet et via le VPN domicile. Aucun port PostgreSQL n'est publié. Ne pas ouvrir le service sur Internet pour contourner le VPN.

L'image retenue est `postgres:17.11-bookworm`. Son [Dockerfile officiel](https://github.com/docker-library/postgres/blob/master/17/bookworm/Dockerfile) définit l'UID PostgreSQL 999. L'API a un compte de base non superutilisateur et des comptes applicatifs distincts.

## Mise en place unique

1. Copier le dossier de version complet dans un emplacement administré du NAS, par exemple `/volume1/docker/cabinet-code/versions/cabinet-2026.09.12`. Ne pas placer les secrets dans un partage accessible aux postes.
2. Copier `Serveur/.env.example` en `Serveur/.env`. Adapter les trois volumes, le chemin UNC et l'UID/GID d'un compte de service NAS. Ce compte doit lire les ressources et écrire les brouillons/archives du partage. PostgreSQL et son dossier de données restent réservés à son UID.
3. Créer les trois répertoires et accorder les droits nécessaires. Réserver le sous-dossier `Documents` à l'écriture de l'API ; les postes du cabinet doivent seulement pouvoir le lire. Ils doivent pouvoir écrire dans `Patients` pour les brouillons. Éviter un partage donnant l'écriture globale aux archives.
4. Depuis `Serveur`, générer les secrets sur le NAS :

```sh
python3 preparer_secrets.py
sudo chown 999:999 secrets/admin_password.txt
sudo chmod 600 secrets/admin_password.txt
sudo chown 999:100 secrets/db_password.txt
sudo chmod 640 secrets/db_password.txt
sudo chmod 700 secrets
```

Remplacer `100` par le GID API configuré. Le dossier parent des secrets doit appartenir à l'administrateur qui lance Compose. Le fichier DB est lisible par PostgreSQL et par le groupe de l'API ; seul ce secret est monté dans l'API. Le générateur ne remplace jamais un secret existant.

5. Initialiser les fichiers de support depuis un PC Windows, sans toucher aux classeurs existants :

```powershell
.\Build\initialiser_nas.ps1 -RacineNas '\\DS224\CabinetCardio'
```

Pour une base d'essai, utiliser le partage d'essai, jamais le partage réel. Les copies initiales incluent l'annuaire et les dictionnaires qui serviront à la migration.

6. Démarrer depuis `Serveur` :

```sh
docker compose up -d --build
```

Configurer le reverse proxy DSM ; vérifier `https://adresse-du-service/health`. Il doit indiquer `status: ok` et `protocole: 2`. Si `init-db.sh` échoue sur un nouveau volume, corriger la cause et recréer uniquement ce volume neuf, sans données. L'initialisation PostgreSQL ne se rejoue pas automatiquement sur un volume déjà initialisé.

## Migration des données

Arrêter les anciens clients, terminer ou annuler leurs files d'arrivée et de courriers, puis sauvegarder le partage original. La simulation vérifie les identités, dates, références, identifiants et files encore actives.

```sh
docker compose exec -T api python -m cabinet.migration
```

Lire le rapport et conserver son empreinte. Les erreurs empêchent l'import ; les avertissements signalent notamment les anciennes clés de destinataire ambiguës, à corriger dans le nouvel annuaire avant de les utiliser. Les ressources initiales livrées ne contiennent aucun patient : elles produisent 259 correspondants, 6 actes, 172 médicaments et 53 expressions. Les tarifs sont repris de vos fichiers et doivent être contrôlés par le cabinet.

```sh
docker compose exec -T api python -m cabinet.migration --appliquer --empreinte-validee EMPREINTE_DE_LA_SIMULATION
```

L'import est transactionnel et refusé si la base cible contient déjà des données. Rejouer le même import renvoie « déjà importé ». Les fichiers Excel restent intacts ; les changements ultérieurs passent par les applications et le service. Les séances historiques sont identifiées par année et conservent leurs montants. Cette commande ne fusionne pas deux bases ayant divergé.

Les formes `Specialistes_ParType` enrichissent les spécialistes existants via `ID_Specialiste`. Plusieurs clés d'examen peuvent être des alias de la même fiche. Une ancienne clé partagée par plusieurs personnes reste bloquée ; le choix explicite d'une fiche utilise son ID unique.

## Comptes et activation des postes

```sh
docker compose exec -T api python -m cabinet.admin compte secretariat --roles secretariat
docker compose exec -T api python -m cabinet.admin compte medecin-cabinet --roles medecin
docker compose exec -T api python -m cabinet.admin compte domicile --roles medecin secretariat
```

Chaque commande affiche **une fois** le jeton à fournir à l'installateur du poste concerné. Ne pas l'enregistrer dans un transcript, une capture ou un dossier partagé. La base conserve une empreinte SHA-256 du jeton aléatoire, pas sa valeur. Révocation :

```sh
docker compose exec -T api python -m cabinet.admin revoquer domicile
```

Les reprises de consultation sont limitées au compte qui a réservé la consultation. Avec des comptes distincts domicile/cabinet, terminer la consultation sur son poste d'origine ; le transfert entre comptes n'est pas automatisé. Un même compte nominatif peut être utilisé sur les deux postes du même médecin si ses rôles conviennent ; ne pas partager son jeton avec le secrétariat.

Installer ensuite les postes selon [INSTALLATION_MULTI_POSTES.md](../INSTALLATION_MULTI_POSTES.md). Migrer tous les postes lors du même arrêt ; les anciennes écritures directes dans Excel ne doivent pas continuer.

## Sauvegarde et reprise

La procédure U0 remplace les anciens scripts de contrôle et de restauration. Suivre [U0_RECETTE.md](../U0_RECETTE.md) : suspension effective des écritures SMB, capture de la base et des fichiers/configuration, vérification dans un cluster isolé, puis restauration persistante sur une cible vide pour la recette Windows. Les anciens jeux sans manifeste U0 restent conservés, mais ne sont pas acceptés silencieusement par ce nouveau vérificateur. Ne pas fabriquer de marqueur `TERMINE` pour les convertir.

## Limites d'exploitation

Ces scripts n'ont pas été exécutés sur votre DSM. Les UID, droits SMB, certificat, reverse proxy, volume libre et restauration physique doivent être contrôlés sur le NAS. Les comptes sont des jetons applicatifs ; l'authentification SSO/MFA n'est pas implémentée. Le service conserve les ressources métier en JSONB versionné ; une évolution de schéma exige une migration explicite, pas une modification manuelle des tables.
