#!/bin/sh
# Restaure uniquement vers un NOUVEAU projet NAS et des volumes VIDES.
set -eu
cd "$(dirname "$0")"
[ "$#" -ge 1 ] || { echo 'Usage : sh restaurer_nas_vide.sh cabinet-HORODATAGE [--appliquer]'; exit 2; }
case "$1" in cabinet-*[!a-zA-Z0-9_-]*|'') echo 'Nom invalide'; exit 2;; cabinet-*) ;; *) exit 2;; esac
if [ "${2:-}" != '--appliquer' ]; then
    echo 'Simulation : cible definie dans .env. Elle doit etre un NOUVEAU projet, base et partage vides.'
    echo 'Les volumes existants ne doivent pas etre reutilises. Lire INSTALLATION_NAS.md.'
    exit 0
fi
docker compose stop api
# Aucun redemarrage automatique si une etape echoue.
docker compose run --rm -T -v "${CABINET_RESTORE_DATA:?Chemin ABSOLU du NOUVEAU volume de donnees}:/restore" maintenance sh -eu -c '
    export PGPASSWORD="$(cat /run/secrets/admin_password)"
    test -z "$(find /restore -mindepth 1 -maxdepth 1 -print -quit)"
    test "$(psql -At -c "SELECT count(*) FROM pg_tables WHERE schemaname=\$\$public\$\$")" = 0
    cd "/backups/$1"
    test -f TERMINE
    sha256sum -c SHA256SUMS
    pg_restore --exit-on-error --no-owner --role=cabinet --dbname=cabinet base.dump
    tar -xzf fichiers.tar.gz -C /restore
' sh "$1"
echo 'Restauration terminee. Verifier les droits du volume et reconcilier avant de reconnecter les postes.'
