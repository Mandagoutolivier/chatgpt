"""Commandes executees localement dans le conteneur ; aucune API admin distante."""
from pathlib import Path
import argparse
import hashlib
import json
import secrets
from psycopg.types.json import Jsonb
from .api import service_environnement


def main():
    parser=argparse.ArgumentParser()
    sub=parser.add_subparsers(dest='action',required=True)
    add=sub.add_parser('compte');add.add_argument('identifiant');add.add_argument('--roles',nargs='+',choices=['medecin','secretariat'],required=True)
    revoke=sub.add_parser('revoquer');revoke.add_argument('identifiant')
    sub.add_parser('reconcilier')
    args=parser.parse_args();service=service_environnement();service.initialiser()
    with service.connexion() as db:
        if args.action=='compte':
            token=secrets.token_urlsafe(48)
            db.execute('INSERT INTO comptes(identifiant,token_sha256,roles) VALUES (%s,%s,%s)',
                       (args.identifiant,hashlib.sha256(token.encode()).hexdigest(),Jsonb(args.roles)))
            print(token)  # Affiche une seule fois a l'operateur ; ne pas rediriger vers un journal partage.
        elif args.action=='revoquer':
            db.execute('UPDATE comptes SET actif=false WHERE identifiant=%s',(args.identifiant,))
            print('Compte desactive.')
        elif args.action=='reconcilier':
            referenced=set();missing=[]
            for row in db.execute('SELECT id,donnees FROM publications'):
                for key in ('CheminDocx','CheminPdf'):
                    path=service.documents.resoudre(row['donnees'][key]);referenced.add(path)
                    if not path.is_file():missing.append({'publication':row['id'],'type':key})
            orphan=[p.name for p in service.documents.objects.glob('*') if p.is_file() and p not in referenced]
            pending=db.execute("SELECT id,etat FROM consultations WHERE etat='encours' ORDER BY modifie_le").fetchall()
            print(json.dumps({'archives_manquantes':missing,'fichiers_non_references':orphan,
                              'consultations_en_cours':pending,'suppression_effectuee':False},ensure_ascii=False,indent=2))

if __name__=='__main__':main()
