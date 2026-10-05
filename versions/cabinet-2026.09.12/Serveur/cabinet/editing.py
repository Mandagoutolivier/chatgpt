"""Revisions secretaire : nouvelle publication, jamais remplacement des originaux.

Pas de verrou long pendant la saisie : controle optimiste de la publication source
sous le verrou transactionnel du service. La facturation garde la meme consultation.
"""
from pathlib import PureWindowsPath
from xml.etree import ElementTree as ET
import hashlib
import io
import re
import zipfile
from .domain import Refus, empreinte

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
IDENTITE=('PatientID','ConsultationID','Patient_Nom','Patient_Prenom','Patient_DDN','Patient_Sexe')


def _xml(archive,name):
    if name not in archive.namelist():raise Refus('Reperes du document absents : edition refusee.',422)
    raw=archive.read(name)
    if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():raise Refus('XML de document interdit.',422)
    try:return ET.fromstring(raw)
    except ET.ParseError as exc:raise Refus('XML de document invalide.',422) from exc


def reperes(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        settings=_xml(z,'word/settings.xml')
        variables={}
        for node in settings.iter(W+'docVar'):
            name=node.get(W+'name','')
            if name in variables:raise Refus('Variable de document dupliquee.',422)
            variables[name]=node.get(W+'val','')
        doc=_xml(z,'word/document.xml')
        active={};textes={};noms=set()
        for event,node in ET.iterparse(io.BytesIO(z.read('word/document.xml')),events=('start','end')):
            if event=='start' and node.tag==W+'bookmarkStart':
                name=node.get(W+'name','');ident=node.get(W+'id','')
                if name=='DESTINATAIRE' or re.fullmatch(r'U2ANN_DEST_[0-9]{3}',name):
                    if name in noms or ident in active:raise Refus('Signet destinataire ambigu.',422)
                    noms.add(name);textes[name]=[];active[ident]=name
            elif event=='start' and node.tag==W+'bookmarkEnd':
                active.pop(node.get(W+'id',''),None)
            elif event=='end':
                fragment=''
                if node.tag==W+'t':fragment=node.text or ''
                elif node.tag in {W+'br',W+'cr',W+'p'}:fragment='\n'
                elif node.tag==W+'tab':fragment='\t'
                if fragment:
                    for name in active.values():textes[name].append(fragment)
        if active:raise Refus('Signet destinataire non ferme.',422)
        return variables,{k:''.join(v).strip() for k,v in textes.items()}


def verifier_revision(original,revision,source,service,db):
    before,adresses_before=reperes(original)
    after,adresses_after=reperes(revision)
    for name in IDENTITE:
        if not before.get(name) or before[name]!=after.get(name):
            raise Refus('Identite technique du courrier modifiee ou absente.',422)
    if after['PatientID']!=source['PatientID'] or after['ConsultationID']!=source['ConsultationID']:
        raise Refus('Document rattache a une autre consultation.',422)
    # Le destinataire principal ne fait pas partie de cette commande d edition des annexes.
    if not adresses_before.get('DESTINATAIRE') or adresses_before['DESTINATAIRE']!=adresses_after.get('DESTINATAIRE'):
        raise Refus('Destinataire principal modifie : edition limitee aux destinataires des annexes.',422)
    annexes_before={k for k in adresses_before if k.startswith('U2ANN_DEST_')}
    if not annexes_before:raise Refus('Ce courrier ne contient pas de reperes d annexes editables.',422)
    if annexes_before!={k for k in adresses_after if k.startswith('U2ANN_DEST_')}:
        raise Refus('Annexe ajoutee ou repere supprime : conserver toutes les annexes.',422)
    annexes=[]
    for name in sorted(annexes_before):
        if not adresses_after[name]:raise Refus('Destinataire annexe vide.',422)
        number=name[-3:]
        dest=after.get('AnnexeDestinataire_'+number,'').strip()
        if not dest:raise Refus('Choisissez le destinataire annexe avec le bouton dedie.',422)
        cor=service._record(db,'CORRESPONDANTS',dest)
        if cor['Actif']!='1' or cor['AValider']!='0':raise Refus('Destinataire annexe non valide.')
        normaliser=lambda s:' '.join(s.split())
        if normaliser(adresses_after[name])!=normaliser(cor['BlocDestinataire']):
            raise Refus('Adresse annexe modifiee : selectionnez son correspondant avant enregistrement.')
        annexes.append({'Numero':number,'DestinataireID':dest})
    return annexes


def reviser_publication(service,db,actor,p):
    from psycopg.types.json import Jsonb
    if not re.fullmatch(r'[0-9a-f]{32}',p['revision_id']):raise Refus('Identifiant de revision invalide.',422)
    for k in ('source_sha','sha_docx','sha_pdf'):
        if not re.fullmatch(r'[0-9a-f]{64}',p[k]):raise Refus('Empreinte de revision invalide.',422)
    existing=db.execute('SELECT donnees FROM publications WHERE id=%s',(p['revision_id'],)).fetchone()
    if existing:
        if existing['donnees'].get('empreinte_commande')!=empreinte(p):raise Refus('Revision deja utilisee avec un autre contenu.')
        return existing['donnees']
    row=db.execute('SELECT * FROM publications WHERE id=%s',(p['source_id'],)).fetchone()
    if not row:raise Refus('Publication source introuvable.',404)
    latest=db.execute('SELECT id FROM publications WHERE consultation_id=%s ORDER BY cree_le DESC,id DESC LIMIT 1',(row['consultation_id'],)).fetchone()
    if row['etat']=='remplacee' or latest['id']!=p['source_id']:
        raise Refus('Une nouvelle version existe : conservez votre copie et rechargez la publication.')
    source=row['donnees']
    if p['patient_id']!=source['PatientID'] or p['source_sha']!=source['sha_docx']:
        raise Refus('Patient ou version source differente.',422)
    current=service._record(db,'PATIENTS',source['PatientID'])
    if any(current[k]!=source[k] for k in ('Nom','Prenom','DDN','Sexe')):
        raise Refus('Identite patient modifiee : faire verifier le courrier par le medecin.')
    # Un dossier unique de travail ; ni autre dossier patient ni archive medicale modifiable.
    folder=service.documents.unc / 'Patients' / '_EditionsSecretariat' / p['revision_id']
    if PureWindowsPath(p['docx'])!=folder/'revision.docx' or PureWindowsPath(p['pdf'])!=folder/'revision.pdf':
        raise Refus('Fichiers de revision hors du dossier reserve.',422)
    original=service.documents._lire_valide(source['CheminDocx'],'.docx')
    if hashlib.sha256(original).hexdigest()!=source['sha_docx']:raise Refus('Archive medicale alteree.',503)
    docx=service.documents._lire_valide(p['docx'],'.docx')
    pdf=service.documents._lire_valide(p['pdf'],'.pdf')
    if hashlib.sha256(docx).hexdigest()!=p['sha_docx'] or hashlib.sha256(pdf).hexdigest()!=p['sha_pdf']:
        raise Refus('Copie modifiee pendant l enregistrement. Reexportez la revision.')
    annexes=verifier_revision(original,docx,source,service,db)
    (word_path,word_sha)=service.documents._conserver_octets(docx,'.docx')
    (pdf_path,pdf_sha)=service.documents._conserver_octets(pdf,'.pdf')
    result=dict(source,PublicationID=p['revision_id'],CheminDrapeau=p['revision_id'],
                CheminDocx=word_path,CheminPdf=pdf_path,sha_docx=word_sha,sha_pdf=pdf_sha,
                PublicationSourceID=p['source_id'],PublicationMedicaleID=source.get('PublicationMedicaleID',p['source_id']),
                AuteurRevision=actor['identifiant'],NatureRevision='correction_secretariat',Relu=False,
                DateRevision=service.clock().isoformat(),Annexes=annexes,empreinte_commande=empreinte(p))
    db.execute("UPDATE publications SET etat='remplacee' WHERE id=%s",(p['source_id'],))
    db.execute("INSERT INTO publications(id,consultation_id,etat,donnees) VALUES (%s,%s,'a_traiter',%s)",
               (p['revision_id'],row['consultation_id'],Jsonb(result)))
    db.execute("UPDATE consultations SET etat='publie',modifie_le=now() WHERE id=%s",(row['consultation_id'],))
    # Aucune ecriture dans seances ni changement de reglement.
    return result
