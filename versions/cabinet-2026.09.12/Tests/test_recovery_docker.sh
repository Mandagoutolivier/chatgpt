#!/bin/sh
set -eu
cd "$(dirname "$0")/../Serveur"
project="cabinet-u0-ci-$(date +%s)-$$"
export CABINET_BACKUP_VOLUME=/tmp/cabinet-u0-unused
compose() { docker compose -f compose.verification.yaml -f compose.recette.yaml -p "$project" "$@"; }
cleanup() {
    result=$?
    trap - 0 HUP INT TERM
    compose down --volumes --remove-orphans >/dev/null || result=1
    exit "$result"
}
trap cleanup 0
trap 'exit 130' HUP INT TERM
docker build -t cabinet-maintenance:2026.09.14-u0 -f Dockerfile.maintenance ..
compose up -d --wait verification-db
compose run --rm -T recette
