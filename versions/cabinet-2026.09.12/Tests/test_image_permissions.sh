#!/usr/bin/env bash
# Reproduit une construction depuis des sources NAS accessibles au seul proprietaire.
# L'import et la lecture SQL restent hors connexion, sans initialiser de base.
set -euo pipefail

racine=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)
contexte=$(mktemp -d)
image_test="cabinet-permissions-test:${GITHUB_RUN_ID:-local}-$$"
nettoyer() {
    rm -rf -- "$contexte"
    docker image rm "$image_test" >/dev/null 2>&1 || true
}
trap nettoyer EXIT

mkdir -p "$contexte/Serveur" "$contexte/Build"
cp "$racine/Serveur/Dockerfile" "$racine/Serveur/requirements.txt" "$contexte/Serveur/"
cp -R "$racine/Serveur/cabinet" "$contexte/Serveur/"
cp "$racine/Build/schemas.json" "$contexte/Build/"
find "$contexte" -type d -exec chmod 700 {} +
find "$contexte" -type f -exec chmod 600 {} +

docker build --file "$contexte/Serveur/Dockerfile" --tag "$image_test" "$contexte"
# Compte par defaut puis compte distinct utilise sur le NAS de test.
for uid_test in 1026 1027; do
    docker run --rm -i --read-only --network none --user "$uid_test:100" \
        "$image_test" python - "$uid_test" <<'PY'
import os
from pathlib import Path
import sys

assert (os.getuid(), os.getgid()) == (int(sys.argv[1]), 100)
import cabinet.api

assert cabinet.api.app.title == "Cabinet NAS"
schema = Path(cabinet.api.__file__).with_name("schema.sql").read_text(encoding="utf-8")
assert "CREATE TABLE" in schema.upper(), "Schema SQL inaccessible ou vide"
print(f"OK : import API et lecture du schema SQL sous {os.getuid()}:{os.getgid()}")
PY
done
