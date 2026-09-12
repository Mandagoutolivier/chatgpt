#!/bin/sh
set -eu
# Execute uniquement a la creation d un volume PostgreSQL vide.
# Le compte utilise par l API ne peut creer ni role ni base et n est pas superutilisateur.
cabinet_password=$(cat /run/secrets/db_password)
psql -v ON_ERROR_STOP=1 --username "$POSTGRES_USER" --dbname cabinet --set=cabinet_password="$cabinet_password" <<'SQL'
CREATE ROLE cabinet LOGIN PASSWORD :'cabinet_password' NOSUPERUSER NOCREATEDB NOCREATEROLE NOREPLICATION;
REVOKE ALL ON DATABASE cabinet FROM PUBLIC;
GRANT CONNECT ON DATABASE cabinet TO cabinet;
REVOKE CREATE ON SCHEMA public FROM PUBLIC;
GRANT USAGE, CREATE ON SCHEMA public TO cabinet;
SQL
unset cabinet_password
