"""Autorite transactionnelle unique. Aucune ecriture de classeur client."""
from __future__ import annotations
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
import hashlib
import json
import re
import uuid
import psycopg
from psycopg.rows import dict_row
from psycopg.types.json import Jsonb
from .domain import Refus, patient_valide, date_fr, chevauche, creneau, montant, empreinte, comparer_clinique, plier, valider_nir
from .files import Documents

READS = {'table.read', 'record.get', 'attentes', 'reprises', 'publications', 'whoami',
         'correspondent.resolve', 'clinical.compare', 'dictionary.read', 'journal.read', 'nir.validate'}
SECRETARIAT = {'table.add', 'table.update', 'arrive', 'absent', 'cancel_arrival', 'bill', 'printed', 'ack', 'payment'}
MEDECIN = {'claim', 'release', 'draft', 'publish'}
SHARED = {'dictionary.add', 'correspondent.save'}
GENRES = {'PATIENTS', 'CORRESPONDANTS', 'RDV', 'ACTES', 'MEDICAMENTS', 'EXPRESSIONS'}
PREFIX = {'PATIENTS':'P', 'CORRESPONDANTS':'C', 'RDV':'R', 'ACTES':'A', 'MEDICAMENTS':'M', 'EXPRESSIONS':'E'}
STATUTS = {'Prevu', 'Arrive', 'Honore', 'Absent', 'Annule'}
BASE = Path(__file__).resolve().parents[2]
PATIENT_FIELDS = json.loads((BASE / 'Build/schemas.json').read_text())['PATIENTS']
CORRESP_FIELDS = json.loads((BASE / 'Build/schemas.json').read_text())['CORRESPONDANTS'] + [
    'CleDestination', 'ClesDestination', 'TypesExamen', 'ParDefaut', 'AValider', 'StructureID', 'TypeCorrespondant']
RDV_FIELDS = ['ID','PatientID','Date','Heure','DureeMin','Motif','ActePrevu','Statut','Notes','DateCreation','DateModif','Nom','Prenom','DDN','TypeActe','HeureArrivee']
ACTE_FIELDS = ['Code','Libelle','LibelleCourt','Tarif','CodeAssocie','TarifAssocie','LibelleCerfa','Depassement','Actif','Notes']

class Service:
    def __init__(self, dsn: str, documents: Documents, clock=None):
        self.dsn = dsn
        self.documents = documents
        self.clock = clock or (lambda: datetime.now(ZoneInfo('Europe/Paris')))

    def connexion(self):
        return psycopg.connect(self.dsn, row_factory=dict_row, connect_timeout=5,
                               options='-c statement_timeout=15000 -c lock_timeout=10000')

    def initialiser(self):
        with self.connexion() as db:
            db.execute((Path(__file__).with_name('schema.sql')).read_text())

    def compte(self, token: str) -> dict:
        if len(token) < 32: raise Refus('Authentification requise.', 401)
        with self.connexion() as db:
            row = db.execute('SELECT identifiant, roles FROM comptes WHERE token_sha256=%s AND actif',
                             (hashlib.sha256(token.encode()).hexdigest(),)).fetchone()
        if not row: raise Refus('Authentification refusee.', 401)
        return row

    def executer(self, acteur: dict, operation: str, params: dict, request_id: str=''):
        roles = set(acteur['roles'])
        if operation not in READS | SECRETARIAT | MEDECIN | SHARED: raise Refus('Operation inconnue.', 404)
        if operation in SECRETARIAT and 'secretariat' not in roles: raise Refus('Role secretariat requis.', 403)
        if operation in MEDECIN and 'medecin' not in roles: raise Refus('Role medecin requis.', 403)
        if not roles.intersection({'medecin','secretariat'}): raise Refus('Compte sans role cabinet.',403)
        if operation == 'journal.read' and 'secretariat' not in roles: raise Refus('Role secretariat requis.',403)
        if operation in {'attentes','reprises','clinical.compare'} and 'medecin' not in roles: raise Refus('Role medecin requis.',403)
        changing = operation not in READS
        if changing and not re.fullmatch(r'[A-Za-z0-9_-]{16,100}', request_id):
            raise Refus('Identifiant de commande obligatoire.',422)
        digest = empreinte({'operation':operation,'params':params})
        with self.connexion() as db:
            if changing:
                # Petit cabinet : serialiser les mutations evite les courses inter-tables.
                # Verrou libere automatiquement au commit ou rollback, jamais par son age.
                db.execute('SELECT pg_advisory_xact_lock(20260912)')
                old = db.execute('SELECT empreinte,resultat FROM commandes WHERE compte=%s AND id=%s',
                                 (acteur['identifiant'],request_id)).fetchone()
                if old:
                    if old['empreinte'] != digest: raise Refus('Identifiant de commande reutilise avec un contenu different.')
                    return old['resultat']
            result = self._operation(db, acteur, operation, params)
            if changing:
                db.execute('INSERT INTO commandes(compte,id,empreinte,resultat) VALUES (%s,%s,%s,%s)',
                           (acteur['identifiant'],request_id,digest,Jsonb(result)))
                # Identifiants uniquement, aucun corps de courrier ni cle dans le journal.
                object_id = str(result.get('ID',result.get('ConsultationID',''))) if isinstance(result,dict) else ''
                db.execute('INSERT INTO evenements(compte,operation,objet) VALUES (%s,%s,%s)',
                           (acteur['identifiant'],operation,object_id))
            return result

    def _record(self, db, genre: str, ident: str) -> dict:
        row = db.execute('SELECT donnees,revision FROM ressources WHERE genre=%s AND id=%s',(genre,ident)).fetchone()
        if not row: raise Refus('Enregistrement introuvable.',404)
        return dict(row['donnees'], _revision=str(row['revision']))

    def _save(self, db, genre: str, data: dict, update=False) -> dict:
        if genre not in GENRES: raise Refus('Table inconnue.',422)
        data = dict(data)
        ident = str(data.get('ID') or data.get('Code') or PREFIX[genre]+uuid.uuid4().hex)
        if not re.fullmatch(r'[\w .:-]{1,100}',ident): raise Refus('Identifiant invalide.',422)
        expected = data.pop('_revision',None)
        old = None
        if update:
            old = self._record(db,genre,ident)
            if expected is None or str(expected) != old['_revision']:
                raise Refus('Fiche modifiee sur un autre poste : rechargez avant de sauver.')
            merged = dict(old); merged.update(data); merged.pop('_revision',None); data = merged
        elif db.execute('SELECT 1 FROM ressources WHERE genre=%s AND id=%s',(genre,ident)).fetchone():
            raise Refus('Cet identifiant existe deja.')
        fields = {'PATIENTS':PATIENT_FIELDS,'CORRESPONDANTS':CORRESP_FIELDS,'RDV':RDV_FIELDS,'ACTES':ACTE_FIELDS}.get(genre)
        if fields:
            if set(data) - set(fields) - {'ID'}: raise Refus('Champ inconnu : '+', '.join(sorted(set(data)-set(fields)-{'ID'})),422)
            data = {**dict.fromkeys(fields,''),**data}
        if any(not isinstance(v,str) for v in data.values()): raise Refus('Les champs doivent etre du texte.',422)
        if any(len(v)>20000 or any(ord(c)<32 and c not in '\r\n\t' for c in v) for v in data.values()):
            raise Refus('Champ trop long ou caractere interdit.',422)
        data['ID']=ident
        if genre == 'PATIENTS':
            patient_valide(data,self.clock().date())
            if data.get('MedTraitantID'): self._record(db,'CORRESPONDANTS',data['MedTraitantID'])
        elif genre == 'CORRESPONDANTS':
            if not data['Nom'].strip(): raise Refus('Nom du correspondant requis.',422)
            data['Actif']=data['Actif'] or '1'
            for flag in ('Actif','AValider','ParDefaut'):
                if data[flag] not in {'','0','1'}:raise Refus('Indicateur invalide : '+flag,422)
            data['CleDestination']=data['CleDestination'] or ident
            if not data['BlocDestinataire'].strip():
                data['BlocDestinataire']='\n'.join([data['Nom']+' '+data['Prenom'],data['Adresse1'],data['Adresse2'],data['CP']+' '+data['Ville']]).strip()
        elif genre == 'RDV':
            creneau(data)
            self._record(db,'PATIENTS',data['PatientID'])
            if data['Statut'] not in STATUTS: raise Refus('Statut agenda inconnu.',422)
            # Seules les commandes metier modifient les statuts liees aux consultations.
            if not old and data['Statut'] not in {'Prevu','Absent','Annule'}:
                raise Refus('Un nouveau rendez-vous doit etre prevu, absent ou annule.')
            if old and data['Statut'] != old['Statut'] and data['Statut'] not in {'Prevu','Absent','Annule'}:
                raise Refus('Utilisez la commande metier de changement de statut.')
            for r in db.execute("SELECT id,donnees FROM ressources WHERE genre='RDV' AND donnees->>'Date'=%s",(data['Date'],)):
                if r['id'] != ident and chevauche(data,r['donnees']): raise Refus('Creneau deja occupe.')
            state = db.execute('SELECT etat FROM consultations WHERE rdv_id=%s',(ident,)).fetchone()
            if state and old and old['Statut']=='Annule' and data['Statut']=='Prevu':
                raise Refus('Une arrivee annulee existe deja : creez un nouveau rendez-vous.')
            if state and state['etat'] not in {'annule'} and old and any(data[k]!=old[k] for k in ('Date','Heure','PatientID','Statut')):
                raise Refus('Consultation deja commencee : modification du rendez-vous refusee.')
        elif genre == 'ACTES':
            if data['Code'] != ident: raise Refus('Le code d un acte doit etre identique a son identifiant.',422)
            data['LibelleCourt'] = data['LibelleCourt'] or data['Libelle']
            data['Libelle'] = data['Libelle'] or data['LibelleCourt']
            montant(data['Tarif'])
            if data.get('CodeAssocie'): montant(data['TarifAssocie'])
        if update:
            db.execute('UPDATE ressources SET donnees=%s,revision=revision+1 WHERE genre=%s AND id=%s',(Jsonb(data),genre,ident))
        else:
            db.execute('INSERT INTO ressources(genre,id,donnees) VALUES (%s,%s,%s)',(genre,ident,Jsonb(data)))
        return self._record(db,genre,ident)

    def _consultation(self, db, ident, actor=None):
        row = db.execute('SELECT * FROM consultations WHERE id=%s',(ident,)).fetchone()
        if not row: raise Refus('Consultation introuvable.',404)
        if actor and row['proprietaire'] != actor['identifiant']: raise Refus('Consultation reservee par un autre compte.',403)
        return row

    def _set_rdv(self, db, ident, status):
        db.execute("UPDATE ressources SET donnees=jsonb_set(donnees,'{Statut}',%s),revision=revision+1 WHERE genre='RDV' AND id=%s",(Jsonb(status),ident))

    def _operation(self, db, actor, op, p):
        if op == 'whoami': return {'ID':actor['identifiant'],'roles':actor['roles'],'version':'2026.09.12','protocole':2}
        if op == 'record.get': return self._record(db,p['genre'],p['id'])
        if op in {'table.add','table.update','correspondent.save'}:
            genre='CORRESPONDANTS' if op=='correspondent.save' else p['genre']
            if genre=='CORRESPONDANTS' and op != 'correspondent.save' and 'secretariat' not in actor['roles']: raise Refus('Role requis.',403)
            return self._save(db,genre,p['data'],op=='table.update' or (op=='correspondent.save' and bool(p['data'].get('_revision'))))
        if op in {'table.read','dictionary.read'}:
            genre=p['genre']
            if genre not in GENRES: raise Refus('Table inconnue.',422)
            limit=min(200,max(1,int(p.get('limit',200)))); offset=max(0,int(p.get('offset',0)))
            args=[genre];where=['genre=%s']
            if p.get('q'):
                # Le filtre reste cote NAS ; aucune copie complete de Patients.xlsx.
                where.append("lower(concat_ws(' ',donnees->>'Nom',donnees->>'Prenom',donnees->>'DDN',donnees->>'ID')) LIKE %s")
                args.append('%'+str(p['q']).lower()+'%')
            if p.get('year') and genre=='RDV': where.append("right(donnees->>'Date',4)=%s");args.append(str(p['year']))
            if p.get('date') and genre=='RDV': where.append("donnees->>'Date'=%s");args.append(p['date'])
            rows=db.execute('SELECT donnees,revision FROM ressources WHERE '+' AND '.join(where)+' ORDER BY id LIMIT %s OFFSET %s',args+[limit+1,offset]).fetchall()
            return {'items':[dict(r['donnees'],_revision=str(r['revision'])) for r in rows[:limit]],
                    'next':offset+limit if len(rows)>limit else None}
        if op=='arrive':
            rdv=self._record(db,'RDV',p['id']);pat=self._record(db,'PATIENTS',rdv['PatientID'])
            if date_fr(rdv['Date']) != self.clock().date(): raise Refus('Seul un rendez-vous du jour peut etre marque arrive.')
            if rdv['Statut'] in {'Annule','Absent','Honore'}: raise Refus('Rendez-vous non disponible pour une arrivee.')
            patient_valide(pat,self.clock().date())
            ident='consult-'+rdv['ID']
            old=db.execute('SELECT donnees,etat FROM consultations WHERE rdv_id=%s',(rdv['ID'],)).fetchone()
            if old:
                if old['etat']=='annule': raise Refus('Arrivee annulee : creez un nouveau rendez-vous.')
                return old['donnees']
            data={k:pat.get(k,'') for k in ['Nom','Prenom','DDN','Sexe','MedTraitantID']}
            data.update(ID=ident,ConsultationID=ident,PatientID=pat['ID'],RdvID=rdv['ID'],DateRdv=rdv['Date'],
                        HeureRdv=rdv['Heure'],DateArrivee=self.clock().strftime('%d/%m/%Y'),HeureArrivee=self.clock().strftime('%H:%M'),Statut='Arrive',SourceNas=ident)
            db.execute("INSERT INTO consultations(id,rdv_id,patient_id,etat,donnees) VALUES (%s,%s,%s,'arrive',%s)",(ident,rdv['ID'],pat['ID'],Jsonb(data)))
            self._set_rdv(db,rdv['ID'],'Arrive');return data
        if op=='attentes':
            return {'items':[r['donnees'] for r in db.execute("SELECT donnees FROM consultations WHERE etat='arrive' AND donnees->>'DateArrivee'=%s ORDER BY modifie_le",(self.clock().strftime('%d/%m/%Y'),))]}
        if op=='absent':
            rdv=self._record(db,'RDV',p['id'])
            row=db.execute('SELECT * FROM consultations WHERE rdv_id=%s',(p['id'],)).fetchone()
            if row and row['etat'] not in {'arrive','annule'}: raise Refus('Consultation deja reservee ou publiee.')
            if row: db.execute("UPDATE consultations SET etat='annule',proprietaire=NULL,modifie_le=now() WHERE id=%s",(row['id'],))
            self._set_rdv(db,p['id'],'Absent');return {'ID':p['id']}
        if op=='cancel_arrival':
            self._record(db,'RDV',p['id'])
            row=db.execute('SELECT * FROM consultations WHERE rdv_id=%s',(p['id'],)).fetchone()
            if row and row['etat'] not in {'arrive','annule'}: raise Refus('Consultation deja reservee ou publiee.')
            if row: db.execute("UPDATE consultations SET etat='annule',modifie_le=now() WHERE id=%s",(row['id'],))
            self._set_rdv(db,p['id'],'Annule');return {'ID':p['id']}
        if op=='claim':
            row=self._consultation(db,p['id'])
            if row['etat']!='arrive': raise Refus('Cette arrivee a deja ete reservee. Utilisez la reprise.')
            data=dict(row['donnees'],ReservationNas=row['id'])
            db.execute("UPDATE consultations SET etat='encours',proprietaire=%s,donnees=%s,modifie_le=now() WHERE id=%s",(actor['identifiant'],Jsonb(data),row['id']))
            return data
        if op=='release':
            row=self._consultation(db,p['id'],actor)
            if row['etat']!='encours' or row['donnees'].get('CheminBrouillon'): raise Refus('Brouillon present : reprendre la consultation, pas la liberer.')
            db.execute("UPDATE consultations SET etat='arrive',proprietaire=NULL,modifie_le=now() WHERE id=%s",(row['id'],));return {'ID':row['id']}
        if op=='draft':
            row=self._consultation(db,p['id'],actor)
            if row['etat'] not in {'encours','publie'}: raise Refus('Consultation non modifiable.')
            if not self.documents.resoudre(p['path']).is_file(): raise Refus('Brouillon non disponible sur le NAS.')
            data=dict(row['donnees'],CheminBrouillon=p['path'])
            db.execute('UPDATE consultations SET donnees=%s,modifie_le=now() WHERE id=%s',(Jsonb(data),row['id']));return {'ID':row['id']}
        if op=='reprises':
            return {'items':[r['donnees'] for r in db.execute("SELECT donnees FROM consultations WHERE proprietaire=%s AND etat IN ('encours','publie') ORDER BY modifie_le",(actor['identifiant'],))]}
        if op=='publish':
            row=self._consultation(db,p['ConsultationID'],actor)
            if row['etat'] not in {'encours','publie','traite'}: raise Refus('Consultation non publiable.')
            pat=self._record(db,'PATIENTS',row['patient_id'])
            if p.get('PatientID')!=pat['ID']: raise Refus('Patient du courrier different de la consultation.')
            if any(p.get('Patient_'+k)!=pat[k] for k in ('Nom','Prenom','DDN','Sexe')): raise Refus('Identite modifiee depuis la dictee.')
            if p.get('Relu') is not True: raise Refus('Relecture medicale requise avant publication.',422)
            cor=self._record(db,'CORRESPONDANTS',p['DestinataireID'])
            if cor.get('Actif')=='0' or cor.get('AValider')=='1': raise Refus('Destinataire non valide.')
            pub=str(p['PublicationID'])
            existing=db.execute('SELECT donnees FROM publications WHERE id=%s',(pub,)).fetchone()
            if existing:
                if existing['donnees'].get('empreinte_commande') != empreinte(p):raise Refus('Publication deja utilisee avec un contenu different.')
                return existing['donnees']
            docx,sha_docx=self.documents.conserver(p['CheminDocx'],'.docx')
            pdf,sha_pdf=self.documents.conserver(p['CheminPdf'],'.pdf')
            data={k:pat.get(k,'') for k in PATIENT_FIELDS};data.pop('ID',None)
            data.update({k:v for k,v in p.items() if not k.startswith('Patient_')})
            data.update({k:pat.get(k,'') for k in PATIENT_FIELDS if k != 'ID'})
            data['empreinte_commande']=empreinte(p)
            data.update(PatientID=pat['ID'],RdvID=row['rdv_id'],CheminDocx=docx,CheminPdf=pdf,sha_docx=sha_docx,sha_pdf=sha_pdf,
                        CheminDrapeau=pub,SeanceID=row['id'],AnneeAgenda=str(date_fr(row['donnees']['DateRdv']).year))
            if date_fr(data['DateActe']) != date_fr(row['donnees']['DateRdv']):raise Refus('Date de l acte differente du rendez-vous.')
            db.execute("UPDATE publications SET etat='remplacee' WHERE consultation_id=%s AND etat='a_traiter'",(row['id'],))
            db.execute("INSERT INTO publications(id,consultation_id,etat,donnees) VALUES (%s,%s,'a_traiter',%s)",(pub,row['id'],Jsonb(data)))
            db.execute("UPDATE consultations SET etat='publie',modifie_le=now() WHERE id=%s",(row['id'],))
            return data
        if op=='publications':
            return {'items':[r['donnees'] for r in db.execute("SELECT donnees FROM publications WHERE etat='a_traiter' ORDER BY cree_le")]}
        if op=='bill': return self._facturer(db,p)
        if op=='journal.read':
            limit=min(200,max(1,int(p.get('limit',200))));offset=max(0,int(p.get('offset',0)))
            where=[];args=[]
            if p.get('id'):where.append('s.id=%s');args.append(p['id'])
            if p.get('year'):where.append("right(l.value->>'Date',4)=%s");args.append(str(p['year']))
            query="SELECT l.value AS ligne,s.imprimee FROM seances s CROSS JOIN LATERAL jsonb_array_elements(s.lignes) WITH ORDINALITY l(value,n)"
            if where:query+=' WHERE '+' AND '.join(where)
            rows=db.execute(query+' ORDER BY s.id,l.n LIMIT %s OFFSET %s',args+[limit+1,offset]).fetchall()
            return {'items':[dict(r['ligne'],FeuilleSoinsImprimee='O' if r['imprimee'] else 'N') for r in rows[:limit]],
                    'next':offset+limit if len(rows)>limit else None}
        if op=='printed':
            if not db.execute('UPDATE seances SET imprimee=true WHERE id=%s RETURNING id',(p['id'],)).fetchone():raise Refus('Seance introuvable.',404)
            return {'ID':p['id']}
        if op=='payment':
            row=db.execute('SELECT lignes FROM seances WHERE id=%s',(p['id'],)).fetchone()
            if not row: raise Refus('Seance introuvable.',404)
            date_fr(p['date']);lines=row['lignes']
            if not p.get('mode'): raise Refus('Mode de paiement requis.',422)
            for line in lines:line.update(Paye='O',DateEncaissement=p['date'],ModePaiement=p['mode'])
            db.execute('UPDATE seances SET lignes=%s WHERE id=%s',(Jsonb(lines),p['id']));return {'ID':p['id']}
        if op=='ack':
            pub=db.execute('SELECT consultation_id,etat FROM publications WHERE id=%s',(p['id'],)).fetchone()
            if not pub: raise Refus('Publication introuvable.',404)
            if pub['etat']=='remplacee': raise Refus('Une nouvelle version est disponible : rechargez la file.')
            if not db.execute('SELECT 1 FROM seances WHERE id=%s',(pub['consultation_id'],)).fetchone(): raise Refus('Enregistrez les actes avant de terminer.')
            row=self._consultation(db,pub['consultation_id'])
            db.execute("UPDATE publications SET etat='traite' WHERE id=%s",(p['id'],))
            db.execute("UPDATE consultations SET etat='traite',modifie_le=now() WHERE id=%s",(row['id'],))
            self._set_rdv(db,row['rdv_id'],'Honore');return {'ID':p['id']}
        if op=='correspondent.resolve':
            candidates=[]
            for r in db.execute("SELECT donnees,revision FROM ressources WHERE genre='CORRESPONDANTS'"):
                d=r['donnees']
                if d.get('Actif')=='0' or d.get('AValider')=='1':continue
                if p.get('id') and d['ID']==p['id']:candidates.append(dict(d,_revision=str(r['revision'])))
                elif p.get('cle') and plier(p['cle']) in [plier(x) for x in (d.get('CleDestination','')+';'+d.get('ClesDestination','')+';'+d['ID']).split(';')]:candidates.append(dict(d,_revision=str(r['revision'])))
                elif p.get('examen') and d.get('ParDefaut')=='1' and plier(p['examen']) in [plier(x) for x in d.get('TypesExamen','').split(';')]:candidates.append(dict(d,_revision=str(r['revision'])))
            if len(candidates)!=1:raise Refus('Destinataire absent ou ambigu : selectionnez un correspondant identifie.')
            return candidates[0]
        if op=='dictionary.add':
            if p['genre'] not in {'MEDICAMENTS','EXPRESSIONS'}:raise Refus('Dictionnaire inconnu.',422)
            text=str(p['texte']).strip()
            if not text or len(text)>200:raise Refus('Expression invalide.',422)
            ident=empreinte(plier(text))
            old=db.execute('SELECT donnees FROM ressources WHERE genre=%s AND id=%s',(p['genre'],ident)).fetchone()
            if old:return old['donnees']
            return self._save(db,p['genre'],{'ID':ident,'Terme':text,'Variantes':'','Actif':'1'})
        if op=='nir.validate':return {'nir':valider_nir(p['nir'])}
        if op=='clinical.compare':return comparer_clinique(p['source'],p['resultat'])
        raise Refus('Operation non implementee.',404)

    def _facturer(self,db,p):
        row=self._consultation(db,p['id'])
        if row['etat'] not in {'publie','traite'}:raise Refus('Courrier non publie : facturation refusee.')
        lines=[dict(x) for x in p['lignes']] if isinstance(p.get('lignes'),list) else None
        if not isinstance(lines,list) or not lines:raise Refus('Aucun acte.',422)
        seen=set();day=None
        for line in lines:
            if line['PatientID']!=row['patient_id'] or line['SeanceID']!=row['id']:raise Refus('Actes rattaches a une autre consultation.')
            current=date_fr(line['Date'])
            if day and current!=day:raise Refus('Dates de seance incoherentes.',422)
            day=current
            if current != date_fr(row['donnees']['DateRdv']):raise Refus('Date comptable differente de la consultation.')
            if line['CodeActe'] in seen:raise Refus('Acte en double.',422)
            seen.add(line['CodeActe']);line['Montant']=str(montant(line['Montant']))
            if line.get('TiersPayant')=='O' and line.get('Paye')=='O':raise Refus('Tiers payant non encaisse : utiliser la commande de reglement.',422)
        digest=empreinte(lines)
        old=db.execute('SELECT patient_id,empreinte FROM seances WHERE id=%s',(row['id'],)).fetchone()
        if old:
            if old['patient_id']!=row['patient_id']:raise Refus('Collision de patient.')
            return {'ID':row['id'],'ajoute':False,'selection_differente':old['empreinte']!=digest}
        # Verifier la nomenclature serveur avant la premiere comptabilisation.
        tarifs={}
        for item in db.execute("SELECT donnees FROM ressources WHERE genre='ACTES'"):
            a=item['donnees']
            if a.get('Actif')=='0':continue
            for code,tarif in ((a.get('Code'),a.get('Tarif')),(a.get('CodeAssocie'),a.get('TarifAssocie'))):
                if code:tarifs.setdefault(code,set()).add(str(montant(tarif)))
        for line in lines:
            if line['Montant'] not in tarifs.get(line['CodeActe'],set()):raise Refus('Acte ou tarif different de la nomenclature NAS : rechargez la selection.')
        # Figement de l identite et de l assure pour toutes les reimpressions.
        pat=self._record(db,'PATIENTS',row['patient_id'])
        for line in lines:
            for key in ('Nom','Prenom','DDN','NIR','AssureNom','AssurePrenom','AssureDDN','AssureNIR'):line[key]=pat.get(key,'')
        db.execute('INSERT INTO seances(id,patient_id,empreinte,lignes) VALUES (%s,%s,%s,%s)',(row['id'],row['patient_id'],digest,Jsonb(lines)))
        return {'ID':row['id'],'ajoute':True,'selection_differente':False}
