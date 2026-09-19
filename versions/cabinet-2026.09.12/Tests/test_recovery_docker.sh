#!/bin/sh
set -eu
cd "$(dirname "$0")/../Serveur"
project="cabinet-u2b-ci-$(date +%s)-$$"
export CABINET_BACKUP_VOLUME=/tmp/cabinet-u2b-unused
export CABINET_MAINTENANCE_IMAGE=cabinet-maintenance:2026.09.16-u2b
compose() { docker compose -f compose.verification.yaml -f compose.recette.yaml -p "$project" "$@"; }
cleanup() {
    result=$?
    trap - 0 HUP INT TERM
    compose down --volumes --remove-orphans >/dev/null || result=1
    exit "$result"
}
trap cleanup 0
trap 'exit 130' HUP INT TERM
docker build -t "$CABINET_MAINTENANCE_IMAGE" -f Dockerfile.maintenance ..
compose up -d --wait verification-db
compose run --rm -T recette
