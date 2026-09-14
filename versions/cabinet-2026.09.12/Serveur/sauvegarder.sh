#!/bin/sh
# Suspension SMB obligatoire : voir U0_RECETTE.md. Aucun ancien backup efface.
set -eu
cd "$(dirname "$0")"
[ "$#" -eq 1 ] && [ "$1" = '--ecritures-smb-suspendues' ] || {
    echo 'Suspendre les ecritures SMB puis : sh sauvegarder.sh --ecritures-smb-suspendues' >&2
    exit 2
}
docker compose build maintenance
mkdir .cabinet-backup.lock || { echo 'Sauvegarde deja active ou verrou a examiner.' >&2; exit 1; }
api_was_running=''
cleanup() {
    result=$?
    trap - 0 HUP INT TERM
    if [ -n "$api_was_running" ]; then
        docker compose start api >/dev/null || { echo 'Redemarrage API echoue.' >&2; result=1; }
    fi
    rmdir .cabinet-backup.lock || result=1
    exit "$result"
}
trap cleanup 0
trap 'exit 130' HUP INT TERM
api_was_running=$(docker compose ps --status running -q api)
if [ -n "$api_was_running" ]; then docker compose stop api; fi
docker compose run --rm -T maintenance backup --ecritures-smb-suspendues
