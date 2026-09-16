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
from . import APPLICATION_VERSION, PROTOCOL_VERSION, SERVICE_REVISION
from .domain import Refus, patient_valide, date_fr, chevauche, creneau, montant, empreinte, comparer_clinique, plier, valider_nir
from .files import Documents
from .contract import validate
from .maintenance import resultat_rejouable
from .actes import FIELDS as ACTE_FIELDS, normaliser as normaliser_acte
from .correspondants import normaliser_indicateurs

READS = {'table.read', 'record.get', 'attentes', 'reprises', 'publications', 'whoami',
         'correspondent.resolve', 'clinical.compare', 'dictionary.read', 'journal.read', 'nir.validate', 'command.result', 'billing.get', 'stale_arrivals'}
SECRETARIAT = {'table.add', 'table.update', 'arrive', 'cancel_arrival', 'bill', 'printed', 'ack', 'payment', 'agenda.status', 'print.request'}
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
PAYMENT_MODES = {'CB', 'Cheque', 'Especes', 'Virement', 'Impaye'}

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
        if operation in {'attentes','reprises','clinical.compare','stale_arrivals'} and 'medecin' not in roles: raise Refus('Role medecin requis.',403)
        validate(operation, params)
        changing = operation not in READS
        if changing and not re.fullmatch(r'[A-Za-z0-9_-]{16,100}', request_id):
            raise Refus('Identifiant de commande obligatoire.',422)
        digest = empreinte({'operation':operation,'params':params})
        with self.connexion() as db:
            if changing:
                # Petit cabinet : serialiser les mutations evite les courses inter-tables.
                # Verrou libere automatiquement au commit ou rollback, jamais par son age.
                db.execute('SELECT pg_advisory_xact_lock(20260912)')
                compte = db.execute('SELECT actif,roles FROM comptes WHERE identifiant=%s FOR SHARE', (acteur['identifiant'],)).fetchone()
                if not compte or not compte['actif'] or set(compte['roles']) != roles:
                    raise Refus('Compte revoque ou modifie ; reconnectez le poste.', 401)
                old = db.execute('SELECT empreinte,resultat FROM commandes WHERE compte=%s AND id=%s',
                                 (acteur['identifiant'],request_id)).fetchone()
                if old:
                    if old['empreinte'] != digest: raise Refus('Identifiant de commande reutilise avec un contenu different.')
                    return resultat_rejouable(old['resultat'])
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
        data = normaliser_acte(row['donnees'], ident) if genre == 'ACTES' else row['donnees']
        if genre == 'CORRESPONDANTS': data = normaliser_indicateurs(data)
        return dict(data, _revision=str(row['revision']))

    def _save(self, db, genre: str, data: dict, update=False) -> dict:
        if genre not in GENRES: raise Refus('Table inconnue.',422)
        data = dict(data)
        if genre == 'ACTES' and not update and not data.get('Code', '').strip():
            raise Refus('Code ACTES explicite obligatoire.', 422)
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
        if genre == 'ACTES': data = normaliser_acte(data, ident, old)
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
            data = normaliser_indicateurs(data)
            data['CleDestination']=data['CleDestination'] or ident
            if not data['BlocDestinataire'].strip():
                data['BlocDestinataire']='\n'.join(x.strip() for x in [data['Nom']+' '+data['Prenom'],data['Adresse1'],data['Adresse2'],data['CP']+' '+data['Ville']] if x.strip())
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
            if state and state['etat'] == 'annule' and old and data['Statut'] == 'Prevu' and old['Statut'] != 'Prevu':
                raise Refus('Arrivee annulee : creez un nouveau rendez-vous.')
            if old and data['Statut'] != old['Statut']:
                raise Refus('Utilisez la commande atomique de statut agenda.')
            if state and state['etat'] not in {'annule'} and old and any(data[k]!=old[k] for k in ('Date','Heure','PatientID','Statut')):
                raise Refus('Consultation deja commencee : modification du rendez-vous refusee.')
        elif genre == 'ACTES':
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
        if op == 'whoami':
            schema = db.execute('SELECT max(version) AS version FROM schema_version').fetchone()['version']
            return {'ID':actor['identifiant'],'roles':actor['roles'],'version':APPLICATION_VERSION,
                    'protocole':PROTOCOL_VERSION,'schema':schema,'revision':SERVICE_REVISION}
        if op == 'command.result':
            row = db.execute('SELECT resultat FROM commandes WHERE compte=%s AND id=%s', (actor['identifiant'], p['id'])).fetchone()
            return {'trouve': row is not None, 'resultat': resultat_rejouable(row['resultat']) if row else None}
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
            if p.get('year') and genre=='RDV': where.append("right(donnees->>'Date',4)=%s");args.append(str(p['year']))
            if p.get('date') and genre=='RDV': where.append("donnees->>'Date'=%s");args.append(p['date'])
            query='SELECT donnees,revision FROM ressources WHERE '+' AND '.join(where)+' ORDER BY id'
            search=plier(p.get('q',''))
            if search:
                # Meme normalisation Unicode des deux cotes ; %, _ et \\ sont litteraux.
                # Curseur serveur : memoire bornee, pagination apres filtrage.
                rows=[]; matched=0
                with db.cursor(name='recherche_'+uuid.uuid4().hex) as cursor:
                    cursor.itersize=200
                    cursor.execute(query,args)
                    for row in cursor:
                        haystack=plier(' '.join(row['donnees'].get(k,'') for k in ('Nom','Prenom','DDN','ID')))
                        if search not in haystack: continue
                        if matched >= offset: rows.append(row)
                        matched+=1
                        if len(rows)>limit: break
            else:
                rows=db.execute(query+' LIMIT %s OFFSET %s',args+[limit+1,offset]).fetchall()
            return {'items':[dict(normaliser_acte(r['donnees'], r['donnees'].get('ID')) if genre == 'ACTES' else normaliser_indicateurs(r['donnees']) if genre == 'CORRESPONDANTS' else r['donnees'],_revision=str(r['revision'])) for r in rows[:limit]],
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
        if op=='stale_arrivals':
            return {'items':[dict(r['donnees'], Etat=r['etat']) for r in db.execute("SELECT donnees,etat FROM consultations WHERE etat='arrive' AND donnees->>'DateArrivee'<>%s ORDER BY modifie_le", (self.clock().strftime('%d/%m/%Y'),))]}
        if op in {'cancel_arrival', 'agenda.status'}:
            rdv = self._record(db, 'RDV', p['id'])
            target = p.get('statut', 'Annule')
            if op == 'agenda.status' and p['revision'] != rdv['_revision']:
                raise Refus('Rendez-vous modifie : rechargez avant de changer son statut.')
            if target not in {'Prevu', 'Absent', 'Annule'}: raise Refus('Transition agenda invalide.', 422)
            row = db.execute('SELECT * FROM consultations WHERE rdv_id=%s', (p['id'],)).fetchone()
            if rdv['Statut'] == 'Honore' or (row and row['etat'] not in {'arrive', 'annule'}):
                raise Refus('Consultation deja reservee, publiee ou honoree.')
            if target == 'Prevu' and row:
                raise Refus('Arrivee annulee ou ancienne : creez un nouveau rendez-vous.')
            if target == 'Prevu':
                candidate = dict(rdv, Statut=target)
                for other in db.execute("SELECT id,donnees FROM ressources WHERE genre='RDV' AND donnees->>'Date'=%s", (rdv['Date'],)):
                    if other['id'] != p['id'] and chevauche(candidate, other['donnees']): raise Refus('Creneau deja occupe.')
            if row: db.execute("UPDATE consultations SET etat='annule',modifie_le=now() WHERE id=%s", (row['id'],))
            self._set_rdv(db, p['id'], target)
            return self._record(db, 'RDV', p['id'])
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
            if cor['Actif']!='1' or cor['AValider']!='0': raise Refus('Destinataire non valide.')
            pub=str(p['PublicationID'])
            existing=db.execute('SELECT donnees FROM publications WHERE id=%s',(pub,)).fetchone()
            if existing:
                if existing['donnees'].get('empreinte_commande') != empreinte(p):raise Refus('Publication deja utilisee avec un contenu different.')
                return existing['donnees']
            if date_fr(p['DateActe']) != date_fr(row['donnees']['DateRdv']): raise Refus('Date de l acte differente du rendez-vous.')
            for field in ('CheminDocx', 'CheminPdf'):
                if not self.documents.resoudre(p[field]).is_file(): raise Refus('Fichier de publication indisponible.')
            (docx,sha_docx),(pdf,sha_pdf)=self.documents.conserver_publication(p['CheminDocx'],p['CheminPdf'])
            # L'identite historique de facturation reste figee dans seances.
            # La file des courriers ne duplique ni adresse, ni telephone, ni NIR.
            data={k:v for k,v in p.items() if not k.startswith('Patient_') and k not in {'Nom','Prenom','DDN','NIR'}}
            data.update({k:pat.get(k,'') for k in ('Nom','Prenom','DDN','Sexe')})
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
            if p.get('date'):
                date_fr(p['date'])
                where.append("l.value->>'Date'=%s");args.append(p['date'])
            query="SELECT l.value AS ligne,s.imprimee FROM seances s CROSS JOIN LATERAL jsonb_array_elements(s.lignes) WITH ORDINALITY l(value,n)"
            if where:query+=' WHERE '+' AND '.join(where)
            rows=db.execute(query+' ORDER BY s.id,l.n LIMIT %s OFFSET %s',args+[limit+1,offset]).fetchall()
            return {'items':[dict(r['ligne'],FeuilleSoinsImprimee='O' if r['imprimee'] else 'N') for r in rows[:limit]],
                    'next':offset+limit if len(rows)>limit else None}
        if op == 'billing.get': return self._seance(db, p['id'])
        if op == 'print.request':
            saved = self._seance(db, p['id'])
            if saved['impression_etat'] != 'actes_enregistres' and not p['reimpression_confirmee']:
                raise Refus('Resultat papier incertain ou impression deja confirmee : confirmez une reimpression explicite.')
            attempt = uuid.uuid4().hex
            # Avant PrintOut : un crash laisse un resultat inconnu, jamais une
            # autorisation de reimpression automatique.
            db.execute("UPDATE seances SET impression_etat='inconnue',impression_tentative=%s WHERE id=%s", (attempt, p['id']))
            return {'ID': p['id'], 'tentative': attempt, 'impression_etat': 'inconnue'}
        if op == 'printed':
            saved = self._seance(db, p['id'])
            if not p['confirmee'] or p['tentative'] != saved['tentative'] or not p['tentative']:
                raise Refus('Confirmation papier et tentative concordante requises.')
            db.execute("UPDATE seances SET imprimee=true,impression_etat='confirmee' WHERE id=%s", (p['id'],))
            return self._seance(db, p['id'])
        if op == 'payment':
            saved = self._seance(db, p['id']); lines = saved['lignes']
            if p['empreinte'] != saved['empreinte']: raise Refus('Reglement modifie : rechargez la seance.')
            date_fr(p['date'])
            if p['mode'] not in PAYMENT_MODES - {'Impaye'}: raise Refus('Mode de paiement invalide.', 422)
            explicit='payeur' in p or 'montant' in p
            if explicit and not {'payeur','montant'} <= p.keys():
                raise Refus('Indiquez le payeur et le montant reellement recu.',422)
            if any(line.get('TiersPayant') == 'O' for line in lines) and not explicit:
                raise Refus('Tiers payant : choisissez le payeur et confirmez le montant recu.',422)
            payeur=p.get('payeur','Patient')
            if payeur not in {'Patient','Organisme'}: raise Refus('Payeur invalide.',422)
            selected=[line for line in lines if line.get('Paye') != 'O' and (line.get('TiersPayant') == 'O') == (payeur == 'Organisme')]
            if not selected: raise Refus('Aucune ligne impayee pour ce payeur : rechargez la seance.')
            total=sum((montant(line['Montant']) for line in selected),montant('0'))
            if total <= 0 or (explicit and montant(p['montant']) != total):
                raise Refus('Montant recu different du solde de ce payeur. Aucun reglement enregistre.',422)
            audit_fields=('CodeActe','Montant','TiersPayant','Paye','DateEncaissement','ModePaiement','PayeurReglement')
            before = [{k: line.get(k, '') for k in audit_fields} for line in lines]
            for line in selected: line.update(Paye='O', DateEncaissement=p['date'], ModePaiement=p['mode'],PayeurReglement=payeur)
            after = [{k: line.get(k, '') for k in audit_fields} for line in lines]
            db.execute('INSERT INTO reglements_audit(seance_id,compte,avant,apres) VALUES (%s,%s,%s,%s)', (p['id'], actor['identifiant'], Jsonb(before), Jsonb(after)))
            db.execute('UPDATE seances SET lignes=%s WHERE id=%s', (Jsonb(lines), p['id']))
            return self._seance(db, p['id'])
        if op=='ack':
            pub=db.execute('SELECT consultation_id,etat FROM publications WHERE id=%s',(p['id'],)).fetchone()
            if not pub: raise Refus('Publication introuvable.',404)
            if pub['etat']=='remplacee': raise Refus('Une nouvelle version est disponible : rechargez la file.')
            if not db.execute('SELECT 1 FROM seances WHERE id=%s',(pub['consultation_id'],)).fetchone(): raise Refus('Enregistrez les actes avant de terminer.')
            if self._seance(db, pub['consultation_id'])['impression_etat'] == 'inconnue':
                raise Refus('Resultat papier inconnu : verifiez l impression avant de terminer.')
            row=self._consultation(db,pub['consultation_id'])
            db.execute("UPDATE publications SET etat='traite' WHERE id=%s",(p['id'],))
            db.execute("UPDATE consultations SET etat='traite',modifie_le=now() WHERE id=%s",(row['id'],))
            self._set_rdv(db,row['rdv_id'],'Honore');return {'ID':p['id']}
        if op=='correspondent.resolve':
            candidates=[]
            for r in db.execute("SELECT donnees,revision FROM ressources WHERE genre='CORRESPONDANTS'"):
                try: d=normaliser_indicateurs(r['donnees'])
                except Refus: continue  # Un statut inconnu ne peut jamais devenir eligible.
                if d['Actif']!='1' or d['AValider']!='0':continue
                if p.get('id'):
                    if d['ID']==p['id']:candidates.append(dict(d,_revision=str(r['revision'])))
                    continue  # Un ID explicite est autoritaire, meme absent/inactif.
                if p.get('cle') and plier(p['cle']) in [plier(x) for x in (d.get('CleDestination','')+';'+d.get('ClesDestination','')+';'+d['ID']).split(';')]:candidates.append(dict(d,_revision=str(r['revision'])))
                elif p.get('examen') and d.get('ParDefaut')=='1' and plier(p['examen']) in [plier(x) for x in d.get('TypesExamen','').split(';')]:candidates.append(dict(d,_revision=str(r['revision'])))
            if len(candidates)!=1:raise Refus('Destinataire absent ou ambigu : selectionnez un correspondant identifie.', 409, 'destination_ambigue' if candidates else 'destination_absente')
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
        pub = db.execute('SELECT consultation_id,etat,donnees FROM publications WHERE id=%s', (p['publication_id'],)).fetchone()
        if not pub or pub['consultation_id'] != p['id'] or pub['etat'] == 'remplacee':
            raise Refus('Une nouvelle version est disponible : rechargez la file avant facturation.')
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
            if line.get('ModePaiement', 'Impaye') not in PAYMENT_MODES: raise Refus('Mode de paiement invalide.', 422)
            if line['Paye'] not in {'O','N'} or line['TiersPayant'] not in {'O','N'}: raise Refus('Indicateur comptable invalide.', 422)
            if line['Paye'] == 'O':
                date_fr(line.get('DateEncaissement', ''))
                if line.get('ModePaiement', 'Impaye') == 'Impaye': raise Refus('Mode de reglement requis.', 422)
        digest=self._empreinte_selection(lines)
        old=db.execute('SELECT patient_id,empreinte FROM seances WHERE id=%s',(row['id'],)).fetchone()
        if old:
            if old['patient_id']!=row['patient_id']:raise Refus('Collision de patient.')
            saved = self._seance(db, row['id'])
            return dict(saved, ajoute=False, selection_differente=self._empreinte_selection(saved['lignes']) != digest)
        # Verifier la nomenclature serveur avant la premiere comptabilisation.
        autorisations={}
        for item in db.execute("SELECT donnees FROM ressources WHERE genre='ACTES'"):
            a=item['donnees']
            if a.get('Actif')=='0':continue
            if a.get('Depassement') and montant(a['Depassement']) != 0:
                if a.get('Code') in seen: raise Refus('Depassement historique non qualifie : verifier la nomenclature avant facturation.')
            couples=((a.get('Code'),a.get('Tarif'),a.get('LibelleCerfa') or a.get('Code')),
                     (a.get('CodeAssocie'),a.get('TarifAssocie'),a.get('CodeAssocie')))
            for code,tarif,cerfa in couples:
                if code:autorisations.setdefault((code,str(montant(tarif))),set()).add(cerfa)
        for line in lines:
            cerfas=autorisations.get((line['CodeActe'],line['Montant']),set())
            if len(cerfas)!=1:raise Refus('Acte, tarif ou libelle CERFA ambigu dans la nomenclature NAS : rechargez la selection.')
            attendu=next(iter(cerfas))
            if line.get('CodeCerfa','') not in {'',attendu}:raise Refus('Libelle CERFA different de la nomenclature NAS : rechargez la selection.')
            line['CodeCerfa']=attendu
        # Figement de l identite et de l assure pour toutes les reimpressions.
        pat=self._record(db,'PATIENTS',row['patient_id'])
        if any(pub['donnees'].get(k,'') != pat.get(k,'') for k in ('Nom','Prenom','DDN')):
            raise Refus('Identite modifiee depuis le courrier : verifier et republier avant facturation.')
        for line in lines:
            for key in ('Nom','Prenom','DDN','NIR','AssureNom','AssurePrenom','AssureDDN','AssureNIR'):line[key]=pat.get(key,'')
        db.execute('INSERT INTO seances(id,patient_id,empreinte,lignes) VALUES (%s,%s,%s,%s)',(row['id'],row['patient_id'],digest,Jsonb(lines)))
        return dict(self._seance(db, row['id']), ajoute=True, selection_differente=False)

    @staticmethod
    def _empreinte_selection(lines):
        keys = ('Date', 'SeanceID', 'PatientID', 'CodeActe', 'Montant', 'TiersPayant')
        return empreinte(sorted(({k: line.get(k, '') for k in keys} for line in lines), key=lambda x: x['CodeActe']))

    def _seance(self, db, ident):
        row = db.execute('SELECT * FROM seances WHERE id=%s', (ident,)).fetchone()
        if not row: raise Refus('Seance introuvable.', 404)
        return {'ID': ident, 'lignes': row['lignes'], 'empreinte': empreinte(row['lignes']),
                'impression_etat': row['impression_etat'], 'tentative': row['impression_tentative']}
