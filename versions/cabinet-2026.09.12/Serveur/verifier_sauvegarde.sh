#!/bin/sh
# Verification et restauration d ESSAI dans une NOUVELLE base, jamais dans cabinet.
set -eu
cd "$(dirname "$0")"
[ "$#" -eq 1 ] || { echo 'Usage : sh verifier_sauvegarde.sh cabinet-HORODATAGE'; exit 2; }
case "$1" in cabinet-*[!a-zA-Z0-9_-]*|'') echo 'Nom invalide'; exit 2;; cabinet-*) ;; *) exit 2;; esac
docker compose run --rm -T maintenance sh -eu -c '
    export PGPASSWORD="$(cat /run/secrets/admin_password)"
    cd "/backups/$1"
    test -f TERMINE
    sha256sum -c SHA256SUMS
    tar -tzf fichiers.tar.gz >/dev/null
    cabinet_test_db="restauration_$(date -u +%Y%m%d%H%M%S)_$$"
    createdb "$cabinet_test_db"
    pg_restore --exit-on-error --no-owner --dbname="$cabinet_test_db" base.dump
    psql --dbname="$cabinet_test_db" -v ON_ERROR_STOP=1 -c "SELECT count(*) AS ressources FROM ressources" -c "SELECT count(*) AS publications FROM publications"
    printf "Base de controle conservee : %s. Ne pas y connecter les postes.\n" "$cabinet_test_db"
' sh "$1"
