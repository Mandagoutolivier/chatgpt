# Maintenance U2c — opérations sur la recette isolée

Ce guide décrit les opérations à qualifier sur des données fictives dans le projet `cabinetcardio-test-u2`. Il ne constate ni déploiement NAS ni validation Office réussis. Le serveur et les clients U2c doivent être qualifiés ensemble ; le RPC authentifié `whoami` annonce la révision `2026.09.21-u2c` et l’installateur exige cette révision, sur schéma 2. `/health` vérifie la disponibilité et le protocole ; il ne donne ni le schéma ni cette révision.

L’installation U0, ses comptes, ses volumes et ses postes restent inchangés. RDC peut rester éteint. Identifier explicitement le poste et le service utilisés pour la recette ; le nom AX8_MAX n’indique pas à lui seul son emplacement ou son rôle.

## Cible explicite des commandes

Suivre d’abord [INSTALLATION_NAS.md](INSTALLATION_NAS.md). Dans chaque nouvelle session shell NAS, définir la fonction suivante en adaptant ensemble les deux chemins absolus si nécessaire :

```sh
u2b_compose() {
    docker compose --project-name cabinetcardio-test-u2 \
        --file /volume1/docker/cabinetcardio-test-u2/code/versions/cabinet-2026.09.12/Serveur/compose.yaml \
        --env-file /volume1/docker/cabinetcardio-test-u2/code/versions/cabinet-2026.09.12/Serveur/.env \
        "$@"
}
```

Afficher `u2b_compose config` et `u2b_compose ps`, puis examiner les montages résolus et le port `127.0.0.1:8766` avant de continuer. Aucun chemin U0 ne doit apparaître. Les variables `CABINET_*` déjà présentes dans le shell peuvent remplacer les valeurs du `.env` ; le nom du projet ne suffit pas à isoler les dossiers montés. Les identifiants et empreintes en capitales ci-dessous sont à remplacer par ceux de la recette.

## État et accès

Dans le conteneur API de recette :

```sh
u2b_compose exec -T api python -m cabinet.admin etat
u2b_compose exec -T api python -m cabinet.admin etat-compte IDENTIFIANT_RECETTE
```

Ces commandes ne créent pas de schéma et ne modifient pas la base. L'état global ne présente que des comptages. L'état du compte donne les rôles, son activation et une empreinte d'approbation, jamais le jeton.

Avant le premier renouvellement après U1, terminer ou diagnostiquer les commandes encore en attente sous les anciens clients. Le client U2 reprend les noms de fichiers U1 tant que l'ancien jeton est encore disponible. Un ancien fichier jamais migré ne peut pas être renommé à partir d'un secret déjà remplacé. Ne pas effacer les fichiers `.pending` pour forcer une reprise.

Après vérification du compte et préparation des postes de recette, créer sur le NAS le dossier privé `/volume1/docker/cabinetcardio-test-u2/jetons-prives`, réservé à l’UID/GID de l’API de recette et accessible en écriture à cet UID. Il ne doit pas se trouver dans un partage de documents. Pour renouveler un jeton, monter ce dossier uniquement dans le conteneur temporaire de l’opération :

```sh
u2b_compose run --rm --no-deps -T \
    --volume /volume1/docker/cabinetcardio-test-u2/jetons-prives:/jetons-prives \
    api python -m cabinet.admin renouveler IDENTIFIANT_RECETTE \
    --empreinte EMPREINTE_LUE --sortie /jetons-prives/nouveau-jeton.token
```

Le chemin de sortie doit être dans un dossier privé monté en écriture pour cette opération, jamais dans un partage de documents cliniques. Le fichier doit être absent ; il est créé avec droits `0600`. Le nouveau secret est écrit et synchronisé avant le commit. Un échec d'écriture ou de transaction conserve l'ancien jeton et retire le fichier créé par l'essai. Un fichier préexistant n'est jamais remplacé.

Après commit, l'ancien jeton devient invalide. Transmettre le nouveau par le mécanisme protégé prévu pour le poste ; ne pas le coller dans un rapport ou une conversation. Les nouvelles commandes locales restent rattachées au même compte lors des renouvellements ultérieurs. Un compte révoqué n'est pas réactivé par renouvellement ; les mutations revérifient son activation sous verrou. Les opérations déjà commencées restent à examiner en cas de révocation urgente.

La révocation est une opération distincte ; ne pas l’exécuter à la suite d’un renouvellement destiné à conserver l’accès :

```sh
u2b_compose exec -T api python -m cabinet.admin revoquer IDENTIFIANT_RECETTE
```

## Conservation des résultats de commandes

Aucune durée automatique n'est imposée et aucune tâche planifiée n'est créée. Définir d'abord la durée applicable aux copies de résultats et qualifier les reprises, puis conserver une sauvegarde cohérente avant toute compaction. Cette fonction ne s'applique pas aux dossiers patients, aux archives de courriers ni à la comptabilité.

Une simulation est obligatoire ; choisir explicitement une date passée avec fuseau :

```sh
u2b_compose exec -T api python -m cabinet.admin compacter-commandes --avant DATE_ISO_AVEC_FUSEAU --limite 100
u2b_compose exec -T api python -m cabinet.admin compacter-commandes --avant DATE_ISO_AVEC_FUSEAU --limite 100 --appliquer EMPREINTE_DU_PLAN
```

Le lot est borné (1 à 500 commandes). La simulation ne présente ni identité ni contenu. L'application recalcule le plan sous le verrou des mutations ; une différence annule tout le lot. Il faut simuler de nouveau pour le lot suivant.

Seule la copie `commandes.resultat` est remplacée par un reçu contenant son empreinte. Le compte, l'identifiant de commande, son empreinte d'entrée et sa date sont conservés sans purge automatique. Une répétition ne peut donc pas exécuter une deuxième création : le serveur répond `409 / resultat_archive` et demande une vérification de reprise. Une recherche depuis un autre compte ne révèle pas cette commande. Les données métier courantes restent lisibles par les opérations habituelles.

**Après compaction, ne pas rétrograder le serveur vers U1**, qui ne reconnaît pas ces reçus. Une récupération doit utiliser U2 ou une restauration cohérente préalablement testée. Les sauvegardes existantes restent soumises à leur propre politique de conservation ; la compaction n'y efface aucun contenu.

## Diagnostics

Chaque RPC journalise une durée en millisecondes, une catégorie d'opération autorisée, l'étape, le statut et des identifiants de corrélation. Les paramètres, chemins patients, textes, jetons et identifiants de commande bruts sont exclus. Les noms d'opération inconnus sont remplacés par `inconnue`. Le service Compose limite ses logs à trois fichiers de 10 Mo.

Ces mesures servent à qualifier la latence réelle. Aucun pool de connexions ni déplacement de copies hors transaction n'est introduit sans mesure et validation du protocole de reprise.
