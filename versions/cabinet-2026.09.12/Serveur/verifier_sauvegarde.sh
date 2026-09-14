#!/bin/sh
# Verification et restauration d ESSAI dans une NOUVELLE base, jamais dans cabinet.
set -eu
cd "$(dirname "$0")"
[ "$#" -ge 1 ] && [ "$#" -le 2 ] || { echo 'Usage : sh verifier_sauvegarde.sh cabinet-HORODATAGE [--conserver]'; exit 2; }
conserver=0
if [ "${2:-}" = "--conserver" ]; then conserver=1; elif [ "$#" -eq 2 ]; then echo 'Option invalide'; exit 2; fi
case "$1" in cabinet-*[!a-zA-Z0-9_-]*|'') echo 'Nom invalide'; exit 2;; cabinet-*) ;; *) exit 2;; esac
docker compose run --rm -T maintenance sh -eu -c '
    export PGPASSWORD="$(cat /run/secrets/admin_password)"
    cd "/backups/$1"
    test -f TERMINE
    sha256sum -c SHA256SUMS
    tar -tzf fichiers.tar.gz >/dev/null
    cabinet_test_db="restauration_$(date -u +%Y%m%d%H%M%S)_$$"
    keep_test_db="$2"
    cleanup() {
        if [ "$keep_test_db" = "1" ]; then
            printf "Base de controle conservee sur demande : %s. Ne pas y connecter les postes.\n" "$cabinet_test_db"
        else
            dropdb --if-exists --force "$cabinet_test_db"
        fi
    }
    trap cleanup EXIT
    trap 'exit 130' HUP INT TERM
    createdb "$cabinet_test_db"
    psql --dbname="$cabinet_test_db" -v ON_ERROR_STOP=1 -c "REVOKE CONNECT ON DATABASE \"$cabinet_test_db\" FROM PUBLIC"
    pg_restore --exit-on-error --no-owner --dbname="$cabinet_test_db" base.dump
    psql --dbname="$cabinet_test_db" -v ON_ERROR_STOP=1 -c "SELECT count(*) AS ressources FROM ressources" -c "SELECT count(*) AS publications FROM publications"
    printf "Verification terminee pour : %s.\n" "$cabinet_test_db"
' sh "$1" "$conserver"
