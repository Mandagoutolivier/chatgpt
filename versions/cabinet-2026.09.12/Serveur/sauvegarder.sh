#!/bin/sh
# Depuis Serveur/ sur le NAS. Clients arretes. Aucun ancien backup efface.
set -eu
cd "$(dirname "$0")"
stamp="cabinet-$(date -u +%Y%m%dT%H%M%SZ)-$$"
docker compose stop api
trap 'docker compose start api >/dev/null' EXIT HUP INT TERM
docker compose run --rm -T maintenance sh -eu -c '
    umask 077
    export PGPASSWORD="$(cat /run/secrets/admin_password)"
    mkdir "/backups/$1"
    pg_dump --format=custom --no-owner --file="/backups/$1/base.dump"
    tar -C /data -czf "/backups/$1/fichiers.tar.gz" .
    cd "/backups/$1"
    sha256sum base.dump fichiers.tar.gz > SHA256SUMS
    pg_restore --list base.dump >/dev/null
    printf "%s\n" "2026.09.12-service-nas" > TERMINE
' sh "$stamp"
printf 'Sauvegarde terminee : %s (dans CABINET_BACKUP_VOLUME).\n' "$stamp"
