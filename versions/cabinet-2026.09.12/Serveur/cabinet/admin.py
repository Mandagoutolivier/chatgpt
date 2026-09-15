"""Commandes executees localement dans le conteneur ; aucune API admin distante."""
from pathlib import Path
import argparse
import hashlib
import json
import secrets
from datetime import datetime
from .maintenance import etat_compte, etat_exploitation, renouveler_jeton, revoquer_compte, compacter_commandes
from psycopg.types.json import Jsonb
from .api import service_environnement


def main():
    parser=argparse.ArgumentParser()
    sub=parser.add_subparsers(dest='action',required=True)
    add=sub.add_parser('compte');add.add_argument('identifiant');add.add_argument('--roles',nargs='+',choices=['medecin','secretariat'],required=True)
    revoke=sub.add_parser('revoquer');revoke.add_argument('identifiant')
    sub.add_parser('reconcilier')
    sub.add_parser('etat')
    account=sub.add_parser('etat-compte');account.add_argument('identifiant')
    rotate=sub.add_parser('renouveler');rotate.add_argument('identifiant');rotate.add_argument('--empreinte',required=True);rotate.add_argument('--sortie',required=True)
    compact=sub.add_parser('compacter-commandes');compact.add_argument('--avant',type=datetime.fromisoformat,required=True);compact.add_argument('--limite',type=int,default=100);compact.add_argument('--appliquer')
    args=parser.parse_args();service=service_environnement()
    maintenance = {
        'etat': lambda: etat_exploitation(service),
        'etat-compte': lambda: etat_compte(service,args.identifiant),
        'renouveler': lambda: renouveler_jeton(service,args.identifiant,args.empreinte,args.sortie),
        'revoquer': lambda: revoquer_compte(service,args.identifiant),
        'compacter-commandes': lambda: compacter_commandes(service,args.avant,args.appliquer,args.limite),
    }
    if args.action in maintenance:
        print(json.dumps(maintenance[args.action](),ensure_ascii=False,indent=2));return
    service.initialiser()
    with service.connexion() as db:
        if args.action=='compte':
            token=secrets.token_urlsafe(48)
            db.execute('INSERT INTO comptes(identifiant,token_sha256,roles) VALUES (%s,%s,%s)',
                       (args.identifiant,hashlib.sha256(token.encode()).hexdigest(),Jsonb(args.roles)))
            print(token)  # Affiche une seule fois a l'operateur ; ne pas rediriger vers un journal partage.
        elif args.action=='reconcilier':
            referenced=set();missing=[];altered=[]
            for row in db.execute('SELECT id,donnees FROM publications'):
                for key in ('CheminDocx','CheminPdf'):
                    path=service.documents.resoudre(row['donnees'][key]);referenced.add(path)
                    if not path.is_file():missing.append({'publication':row['id'],'type':key})
                    else:
                        expected=row['donnees'].get('sha_docx' if key=='CheminDocx' else 'sha_pdf')
                        with path.open('rb') as f:actual=hashlib.file_digest(f,'sha256').hexdigest()
                        if actual!=expected:altered.append({'publication':row['id'],'type':key})
            orphan=[p.name for p in service.documents.objects.glob('*') if p.is_file() and p not in referenced]
            pending=db.execute("SELECT id,etat FROM consultations WHERE etat='encours' ORDER BY modifie_le").fetchall()
            print(json.dumps({'archives_manquantes':missing,'archives_alterees':altered,'fichiers_non_references':orphan,
                              'consultations_en_cours':pending,'suppression_effectuee':False},ensure_ascii=False,indent=2))

if __name__=='__main__':main()
