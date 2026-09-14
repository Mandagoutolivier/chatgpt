#!/bin/sh
# Nouveau cluster, reseau interne et volumes ephemeres ; jamais le cluster du cabinet.
set -eu
cd "$(dirname "$0")"
[ "$#" -eq 1 ] || { echo 'Usage : sh verifier_sauvegarde.sh cabinet-HORODATAGE' >&2; exit 2; }
case "$1" in cabinet-*[!a-zA-Z0-9_-]*|'') exit 2;; cabinet-*) ;; *) exit 2;; esac
docker compose build maintenance
project="cabinet-verif-$(date -u +%Y%m%d%H%M%S)-$$"
verification() { docker compose -f compose.verification.yaml -p "$project" "$@"; }
cleanup() {
    result=$?
    trap - 0 HUP INT TERM
    verification down --volumes --remove-orphans >/dev/null || result=1
    exit "$result"
}
trap cleanup 0
trap 'exit 130' HUP INT TERM
verification up -d --wait verification-db
verification run --rm -T verifier restore --bundle "/backups/$1" --root /restore \
    --config /configuration-restauree --unc '\\NAS-RECETTE\CabinetRestaure' --role postgres
