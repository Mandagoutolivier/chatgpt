"""Migration controlee des classeurs : simulation par defaut, aucun effacement."""
from __future__ import annotations
from pathlib import Path
from datetime import datetime,date,time
from collections import defaultdict
import argparse
import hashlib
import json
import re
import openpyxl
from psycopg.types.json import Jsonb
from .domain import Refus, plier, empreinte, date_fr, patient_valide, creneau, montant
from .api import service_environnement
from .service import PATIENT_FIELDS,CORRESP_FIELDS,RDV_FIELDS,ACTE_FIELDS


def texte(value,field):
    if value is None:return ''
    if isinstance(value,datetime):return value.strftime('%H:%M') if 'Heure' in field else value.strftime('%d/%m/%Y')
    if isinstance(value,date):return value.strftime('%d/%m/%Y')
    if isinstance(value,time):return value.strftime('%H:%M')
    if isinstance(value,bool):return '1' if value else '0'
    return str(value).strip()


def lire(path,sheet):
    if not path.exists():return []
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    try:
        if sheet not in wb.sheetnames:return []
        rows=iter(wb[sheet].values)
        headers=[str(x or '').strip() for x in next(rows,[])]
        out=[]
        for row in rows:
            d={key:texte(value,key) for key,value in zip(headers,row) if key}
            if any(d.values()):out.append(d)
        return out
    finally:wb.close()


def construire_plan(root:Path):
    resources=[];errors=[];warnings=[];correspondants={};fingerprints={};mapping={};specialistes={}
    def ajouter_cor(d,legacy=''):
        d={**dict.fromkeys(CORRESP_FIELDS,''),**d}
        key=plier('|'.join(d.get(k,'') for k in ('Nom','Prenom','Adresse1','CP','Ville','BlocDestinataire')))
        if key in fingerprints:
            ident=fingerprints[key];old=correspondants[ident]
            old['TypesExamen']=';'.join(sorted(set(filter(None,(old['TypesExamen']+';'+d['TypesExamen']).split(';')))))
            if legacy:mapping[legacy]=ident
            return ident
        ident=d.get('ID') or 'C-'+empreinte(key)[:24]
        if ident in correspondants and correspondants[ident]!=d:
            errors.append('Identifiant correspondant en conflit : '+ident);return ident
        d['ID']=ident;d['Actif']=d['Actif'] or '1';d['CleDestination']=d['CleDestination'] or ident
        if not d['BlocDestinataire']:d['BlocDestinataire']='\n'.join([d['Nom']+' '+d['Prenom'],d['Adresse1'],d['Adresse2'],d['CP']+' '+d['Ville']]).strip()
        correspondants[ident]=d;fingerprints[key]=ident
        if legacy:mapping[legacy]=ident
        return ident
    patients=lire(root/'Base/Patients.xlsx','PATIENTS')
    for d in lire(root/'Base/Patients.xlsx','CORRESPONDANTS'):
        ajouter_cor({k:d.get(k,'') for k in CORRESP_FIELDS},d.get('ID',''))
    standalone=root/'Base/base_travail_correspondants_v1.xlsx'
    # Les lignes par examen portent les adresses et cles reellement utilisees.
    for sheet in ('Generalistes','Specialistes','Structures'):
        for old in lire(standalone,sheet):
            nom=old.get('Nom') or old.get('NomDestinataire') or old.get('NomStructure') or old.get('Structure') or ''
            if not nom:continue
            d={'Nom':nom,'Prenom':old.get('Prenom',old.get('PrenomOuInitiale','')),'Adresse1':old.get('AdresseLigne1',old.get('Adresse1','')),
               'Adresse2':old.get('AdresseLigne2',old.get('Adresse2','')),'CP':old.get('CodePostal',old.get('CP','')),'Ville':old.get('Ville',''),
               'BlocDestinataire':old.get('BlocDestinataireComplet',old.get('BlocDestinataire','')),
               'FormuleAppel':old.get('FormuleAppel',''),'FormulePolitesse':old.get('FormulePolitesse',''),
               'Tutoiement':old.get('TutoiementVouvoiement',''),'CleDestination':old.get('CleRegroupement',''),
               'TypesExamen':';'.join(filter(None,[old.get('TypeExamen',''),old.get('TypesExamensPossibles','')])),
               'Titre':old.get('Titre',''),'Tel':old.get('Telephone',''),
               'AValider':'1' if plier(old.get('AValider','')) in {'1','oui','o'} else '0',
               'Actif':'0' if plier(old.get('Actif','')) in {'0','non','n'} else '1',
               'ParDefaut':'1' if old.get('Priorite') in {'1','1.0'} else '0',
               'StructureID':old.get('ID_Structure',''),'TypeCorrespondant':sheet}
            ident=ajouter_cor(d,old.get('ID',old.get('ID_Structure','')))
            if sheet=='Specialistes':specialistes[old.get('ID','')]=ident
    # Une ligne par type d examen enrichit le meme specialiste, sans le dupliquer.
    for old in lire(standalone,'Specialistes_ParType'):
        ident=specialistes.get(old.get('ID_Specialiste',''))
        if not ident:
            errors.append('Ligne par type sans specialiste : '+old.get('ID_Ligne',''));continue
        d=correspondants[ident]
        d['TypesExamen']=';'.join(sorted(set(filter(None,(d['TypesExamen']+';'+old.get('TypeExamen','')).split(';')))))
        for key,source in [('CleDestination','CleRegroupement'),('FormulePolitesse','FormulePolitesse')]:
            if old.get(source):
                if key=='CleDestination':
                    d['ClesDestination']=';'.join(sorted(set(filter(None,(d['ClesDestination']+';'+old[source]).split(';')))))
                    if d[key]==ident:d[key]=old[source]
                elif d.get(key) and d[key]!=old[source]:warnings.append('Formule de politesse differente selon examen : '+ident)
                else:d[key]=old[source]
        if old.get('Priorite') in {'1','1.0'}:d['ParDefaut']='1'
    for d in correspondants.values():
        if d['StructureID']:d['StructureID']=mapping.get(d['StructureID'],d['StructureID'])
    for d in correspondants.values():resources.append({'genre':'CORRESPONDANTS','id':d['ID'],'data':d})
    for old in patients:
        d={k:old.get(k,'') for k in PATIENT_FIELDS}
        if d['MedTraitantID']:d['MedTraitantID']=mapping.get(d['MedTraitantID'],d['MedTraitantID'])
        if not d['ID']:errors.append('Patient sans ID.');continue
        try:patient_valide(d,date.today())
        except Refus as exc:errors.append('Patient '+d['ID']+' : '+str(exc))
        if d['MedTraitantID'] and d['MedTraitantID'] not in correspondants:errors.append('Medecin traitant absent pour '+d['ID'])
        resources.append({'genre':'PATIENTS','id':d['ID'],'data':d})
    pids={x['id'] for x in resources if x['genre']=='PATIENTS'}
    for path in sorted((root/'Base').glob('Agenda_*.xlsx')):
        year=path.stem.rsplit('_',1)[-1]
        for old in lire(path,'RDV'):
            d={k:old.get(k,'') for k in RDV_FIELDS};original=d['ID'];d['ID']='R'+year+'_'+original
            if not original:errors.append('Rendez-vous sans ID dans '+path.name)
            if d['PatientID'] not in pids:errors.append('Patient absent pour '+d['ID'])
            try:creneau(d)
            except Refus as exc:errors.append('Rendez-vous '+d['ID']+' : '+str(exc))
            if d['Statut']=='Arrive':errors.append('Rendez-vous encore arrive : terminer ou annuler avant migration ('+d['ID']+').')
            resources.append({'genre':'RDV','id':d['ID'],'data':d})
    for genre,path,sheet in [('ACTES','Config/Nomenclature.xlsx','ACTES'),('MEDICAMENTS','Config/Gras_Medicaments.xlsx','MEDICAMENTS'),('EXPRESSIONS','Config/Gras_Expressions.xlsx','EXPRESSIONS')]:
        for old in lire(root/path,sheet):
            ident=old.get('Code') or empreinte(plier(old.get('Terme','')))
            if genre=='ACTES':
                try:
                    montant(old.get('Tarif',''))
                    if old.get('CodeAssocie'):montant(old.get('TarifAssocie',''))
                except Refus as exc:errors.append('Nomenclature '+ident+' : '+str(exc))
            elif not old.get('Terme'):errors.append('Terme de dictionnaire vide : '+path)
            resources.append({'genre':genre,'id':ident,'data':dict(old,ID=ident)})
    historical=[]
    for path in sorted((root/'Actes').glob('Journal_*.xlsx')):
        grouped=defaultdict(list)
        for line in lire(path,'JOURNAL'):
            if not line.get('SeanceID'):errors.append('Ligne sans SeanceID dans '+path.name);continue
            grouped[line['SeanceID']].append(line)
        for ident,lines in grouped.items():
            new='H'+path.stem.rsplit('_',1)[-1]+'_'+ident
            ids={x.get('PatientID','') for x in lines}
            if len(ids)!=1 or not ids.issubset(pids):errors.append('Patient incoherent dans seance '+new);continue
            for x in lines:x['SeanceID']=new
            historical.append({'id':new,'patient_id':next(iter(ids)),'lignes':lines})
    for folder in ('Echange/Arrives','Echange/Arrives/EnCours','Echange/AEnvoyer'):
        if list((root/folder).glob('*.txt')):errors.append('File encore active : '+folder)
    keys=defaultdict(set)
    for d in correspondants.values():
        for key in (d['CleDestination']+';'+d['ClesDestination']).split(';'):
            if key:keys[plier(key)].add(d['ID'])
    for key,ids in keys.items():
        if len(ids)>1:warnings.append('Cle de destinataire ambigue, resolution manuelle requise : '+key)
    seen=set()
    for r in resources:
        key=(r['genre'],r['id'])
        if key in seen:errors.append('Identifiant repete : '+str(key))
        seen.add(key)
    return {'ressources':resources,'historique':historical,'erreurs':errors,'avertissements':warnings}


def appliquer(service,plan):
    if plan['erreurs']:raise Refus('Corriger les erreurs de simulation avant import.')
    digest=empreinte(plan)
    with service.connexion() as db:
        db.execute('SELECT pg_advisory_xact_lock(20260912)')
        if db.execute('SELECT 1 FROM imports WHERE empreinte=%s',(digest,)).fetchone():return 'deja importe'
        if db.execute('SELECT 1 FROM ressources LIMIT 1').fetchone():raise Refus('La base cible contient deja des donnees. Aucune fusion automatique.')
        for r in plan['ressources']:db.execute('INSERT INTO ressources(genre,id,donnees) VALUES (%s,%s,%s)',(r['genre'],r['id'],Jsonb(r['data'])))
        for h in plan['historique']:
            db.execute("INSERT INTO consultations(id,rdv_id,patient_id,etat,donnees) VALUES (%s,%s,%s,'traite',%s)",(h['id'],h['id'],h['patient_id'],Jsonb({'historique':True})))
            db.execute('INSERT INTO seances(id,patient_id,empreinte,lignes,imprimee) VALUES (%s,%s,%s,%s,%s)',(h['id'],h['patient_id'],empreinte(h['lignes']),Jsonb(h['lignes']),all(x.get('FeuilleSoinsImprimee')=='O' for x in h['lignes'])))
        db.execute('INSERT INTO imports(empreinte,rapport) VALUES (%s,%s)',(digest,Jsonb({'ressources':len(plan['ressources']),'seances':len(plan['historique'])})))
    return 'importe'


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--racine',type=Path,default=Path('/data'))
    parser.add_argument('--appliquer',action='store_true');parser.add_argument('--empreinte-validee')
    args=parser.parse_args();plan=construire_plan(args.racine)
    report={'empreinte':empreinte(plan),'ressources':len(plan['ressources']),'seances':len(plan['historique']),
            'erreurs':plan['erreurs'],'avertissements':plan['avertissements']}
    print(json.dumps(report,ensure_ascii=False,indent=2))
    if args.appliquer:
        if args.empreinte_validee!=report['empreinte']:raise SystemExit('Reprendre l empreinte de la simulation relue ; aucune modification effectuee.')
        service=service_environnement();service.initialiser();print(appliquer(service,plan))
    elif plan['erreurs']:raise SystemExit(2)

if __name__=='__main__':main()
