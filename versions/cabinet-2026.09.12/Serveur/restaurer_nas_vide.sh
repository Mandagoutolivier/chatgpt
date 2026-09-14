#!/bin/sh
# Nouveau projet Compose, db seule demarree, api arretee.
set -eu
cd "$(dirname "$0")"
[ "$#" -ge 1 ] && [ "$#" -le 2 ] || { echo 'Usage : sh restaurer_nas_vide.sh cabinet-HORODATAGE [--appliquer]' >&2; exit 2; }
case "$1" in cabinet-*[!a-zA-Z0-9_-]*|'') exit 2;; cabinet-*) ;; *) exit 2;; esac
if [ "$#" -eq 1 ]; then
    echo 'Simulation : nouveau projet requis ; CABINET_DATA_VOLUME et base doivent etre vides.'
    echo 'CABINET_UNC definit le nouveau partage. Configuration source recuperee a part.'
    exit 0
fi
[ "$2" = '--appliquer' ] || exit 2
[ -z "$(docker compose ps --status running -q api)" ] || { echo 'API active : cible refusee sans arreter aucun service.' >&2; exit 1; }
docker compose build maintenance
# Meme montage que l API : impossible de restaurer un autre /restore par erreur.
docker compose -f compose.yaml -f compose.restauration.yaml run --rm -T maintenance \
    restore --bundle "/backups/$1" --root /data --config /configuration-restauree
echo 'Controle technique termine. Restaurer les ACL DSM et effectuer la recette Office/SMB avant toute bascule.'
