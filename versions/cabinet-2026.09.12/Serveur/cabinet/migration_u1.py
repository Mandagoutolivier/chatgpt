"""Migration des ACTES deja importes : simulation, empreinte, transaction.

python -m cabinet.migration_u1 [--appliquer EMPREINTE]
Les valeurs precedentes restent dans migrations_ressources (sauvegarde DB).
"""
import argparse
import json
from psycopg.types.json import Jsonb
from .actes import normaliser
from .domain import Refus, empreinte


def planifier(db):
    changes, conflicts = [], []
    for row in db.execute("SELECT id,revision,donnees FROM ressources WHERE genre='ACTES' ORDER BY id"):
        try: new = normaliser(row['donnees'], row['id'])
        except Refus:
            conflicts.append({'id': row['id'], 'categorie': 'contrat_actes'}); continue
        if new != row['donnees']:
            changes.append({'id': row['id'], 'revision': row['revision'], 'avant': row['donnees'], 'apres': new})
    return {'version': 2, 'changements': changes, 'conflits': conflicts}


def migrer(service, approved=None):
    with service.connexion() as db:
        db.execute('SELECT pg_advisory_xact_lock(20260912)')
        plan = planifier(db)
        digest = empreinte(plan)
        if approved is not None:
            if approved != digest or plan['conflits']:
                raise Refus('Simulation modifiee ou conflits non resolus : aucune migration.')
            for item in plan['changements']:
                db.execute("INSERT INTO migrations_ressources(version,genre,id,revision,avant,apres) VALUES (2,'ACTES',%s,%s,%s,%s)",
                           (item['id'], item['revision'], Jsonb(item['avant']), Jsonb(item['apres'])))
                db.execute("UPDATE ressources SET donnees=%s,revision=revision+1 WHERE genre='ACTES' AND id=%s",
                           (Jsonb(item['apres']), item['id']))
            db.execute('INSERT INTO schema_version VALUES (2) ON CONFLICT DO NOTHING')
        return {'empreinte': digest, 'modifications': len(plan['changements']), 'conflits': plan['conflits'],
                'applique': approved is not None}


def main():
    from .api import service_environnement
    parser = argparse.ArgumentParser(); parser.add_argument('--appliquer')
    args = parser.parse_args(); service = service_environnement(); service.initialiser()
    print(json.dumps(migrer(service, args.appliquer), ensure_ascii=False))


if __name__ == '__main__': main()
