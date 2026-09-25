from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import hashlib
import io
import json
import os
import uuid
import zipfile
from xml.etree import ElementTree as ET
import pytest
import psycopg
from psycopg import sql
from psycopg.types.json import Jsonb
from fastapi.testclient import TestClient
from cabinet.service import Service
from cabinet.files import Documents
from cabinet.domain import Refus
from cabinet.api import create_app

@pytest.fixture
def service(tmp_path):
    dsn=os.getenv('CABINET_TEST_DATABASE_URL')
    if not dsn:pytest.skip('CABINET_TEST_DATABASE_URL absent : PostgreSQL requis.')
    schema='cabinet_test_'+uuid.uuid4().hex
    with psycopg.connect(dsn) as db:db.execute(sql.SQL('CREATE SCHEMA {}').format(sql.Identifier(schema)))
    class TestService(Service):
        def connexion(self):
            db=super().connexion()
            db.execute(sql.SQL('SET search_path TO {}').format(sql.Identifier(schema)))
            return db
    service=TestService(dsn,Documents(tmp_path,r'\\NAS-FICTIF\Cabinet'),lambda:datetime(2026,9,12,10,30,tzinfo=ZoneInfo('Europe/Paris')))
    service.initialiser()
    with service.connexion() as db:
        for ident,roles in [('medecin',['medecin']),('secretariat',['secretariat']),('autre',['medecin'])]:
            token=('test-'+ident+'-')*6
            db.execute('INSERT INTO comptes(identifiant,token_sha256,roles) VALUES (%s,%s,%s)',(ident,hashlib.sha256(token.encode()).hexdigest(),Jsonb(roles)))
    yield service
    with psycopg.connect(dsn) as db:db.execute(sql.SQL('DROP SCHEMA {} CASCADE').format(sql.Identifier(schema)))


def rpc(s,op,p=None,role='secretariat',rid=None):
    roles=['medecin'] if role in {'medecin','autre'} else ['secretariat']
    return s.executer({'identifiant':role,'roles':roles},op,p or {},rid or uuid.uuid4().hex)


def parcours(s):
    rpc(s,'table.add',{'genre':'ACTES','data':{'Code':'TEST','Tarif':'12.30','Actif':'1'}})
    cor=rpc(s,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'TEST DESTINATAIRE','Prenom':'Fictif','Adresse1':'ADRESSE FICTIVE','Ville':'TEST','Actif':'1'}})
    pat=rpc(s,'table.add',{'genre':'PATIENTS','data':{'Nom':'PATIENT FICTIF','Prenom':'Essai','Sexe':'F','DDN':'01/01/1980','MedTraitantID':cor['ID']}})
    rdv=rpc(s,'table.add',{'genre':'RDV','data':{'PatientID':pat['ID'],'Date':'12/09/2026','Heure':'10:00','DureeMin':'15','TypeActe':'TEST','Statut':'Prevu'}})
    arr=rpc(s,'arrive',{'id':rdv['ID']})
    return cor,pat,rdv,arr


def publication(s):
    cor,pat,rdv,arr=parcours(s)
    rpc(s,'claim',{'id':arr['ID']},'medecin')
    root=s.documents.root
    with zipfile.ZipFile(root/'essai.docx','w') as z:z.writestr('word/document.xml','<document>FICTIF</document>')
    (root/'essai.pdf').write_bytes(b'%PDF-1.4\n% document fictif pour test de conservation')
    data={'ConsultationID':arr['ID'],'PublicationID':uuid.uuid4().hex,'PatientID':pat['ID'],'DestinataireID':cor['ID'],'DateActe':'12/09/2026',
          'CheminDocx':r'\\NAS-FICTIF\Cabinet\essai.docx','CheminPdf':r'\\NAS-FICTIF\Cabinet\essai.pdf','Relu':True}
    data.update({'Patient_'+k:pat[k] for k in ('Nom','Prenom','DDN','Sexe')})
    return arr,pat,data


def ligne(arr,pat):return {'SeanceID':arr['ID'],'PatientID':pat['ID'],'Date':'12/09/2026','CodeActe':'TEST','Montant':'12.30','TiersPayant':'N','Paye':'N'}


def test_roles_et_authentification(service):
    with TestClient(create_app(service)) as c:
        assert c.post('/v1/rpc',json={'operation':'whoami'}).status_code==401
        token=('test-medecin-')*6
        assert c.post('/v1/rpc',headers={'Authorization':'Bearer '+token},json={'operation':'whoami'}).json()['result']['roles']==['medecin']
    with pytest.raises(Refus,match='secretariat'):rpc(service,'table.add',{'genre':'PATIENTS','data':{}},'medecin')
    with pytest.raises(Refus,match='medecin'):rpc(service,'claim',{'id':'x'})


def test_u0_explicit_correspondent_cannot_fall_back(service):
    first=rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'FICTIF A','CleDestination':'cle-a'}})
    second=rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'FICTIF B','CleDestination':'cle-b'}})
    with pytest.raises(Refus):
        rpc(service,'correspondent.resolve',{'id':'ID_ABSENT','cle':'cle-b'},'medecin')
    assert rpc(service,'correspondent.resolve',{'id':first['ID'],'cle':'cle-b'},'medecin')['ID']==first['ID']
    rpc(service,'table.update',{'genre':'CORRESPONDANTS','data':dict(first,Actif='0')})
    with pytest.raises(Refus):
        rpc(service,'correspondent.resolve',{'id':first['ID'],'cle':'cle-b'},'medecin')


def test_u0_service_release_contract(service):
    result=rpc(service,'whoami',role='medecin')
    assert result['protocole']==2 and result['schema']==1 and result['revision']=='2026.09.21-u2c'


def test_u0_rebase_sql_preserves_identity_and_cached_results(service):
    from cabinet.recovery import rebase_database, regular_tree, check_references
    arr, pat, data=publication(service)
    result=rpc(service,'publish',data,'medecin')
    target=r'\\NAS-RECETTE\AutrePartage'
    with service.connexion() as db:
        rebase_database(db,str(service.documents.unc),target)
        references=check_references(db,regular_tree(service.documents.root),target,service.documents.root)
        assert references['archives']==2 and references['publications']==1
        saved=db.execute('SELECT donnees FROM publications').fetchone()['donnees']
        assert saved['PatientID']==pat['ID'] and saved['sha_docx']==result['sha_docx']
        assert saved['CheminDocx'].startswith(target)
        cached=db.execute("SELECT resultat FROM commandes WHERE resultat ? 'sha_docx'").fetchone()['resultat']
        assert cached['CheminPdf'].startswith(target)


def test_crud_revision_et_pas_de_perte_mise_a_jour(service):
    cor,pat,_,_=parcours(service)
    a=dict(pat,Tel='0100000000');r=rpc(service,'table.update',{'genre':'PATIENTS','data':a})
    assert r['_revision']=='2'
    with pytest.raises(Refus,match='autre poste'):rpc(service,'table.update',{'genre':'PATIENTS','data':dict(pat,Ville='Autre')})
    assert rpc(service,'record.get',{'genre':'PATIENTS','id':pat['ID']})['Tel']=='0100000000'


def test_idempotence_contenu_different_et_audit(service):
    rid=uuid.uuid4().hex;p={'genre':'CORRESPONDANTS','data':{'Nom':'UNIQUE'}}
    a=rpc(service,'table.add',p,rid=rid);b=rpc(service,'table.add',p,rid=rid)
    assert a==b
    with pytest.raises(Refus,match='contenu different'):rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'AUTRE'}},rid=rid)
    with service.connexion() as db:
        assert db.execute('SELECT count(*) AS n FROM evenements').fetchone()['n']==1
        assert db.execute('SELECT objet FROM evenements').fetchone()['objet']==a['ID']


def test_arrivee_unique_et_patient_deux_rdv(service):
    _,pat,rdv,arr=parcours(service)
    assert rpc(service,'arrive',{'id':rdv['ID']})==arr
    r2=rpc(service,'table.add',{'genre':'RDV','data':{'PatientID':pat['ID'],'Date':'12/09/2026','Heure':'11:00','DureeMin':'15','Statut':'Prevu'}})
    rpc(service,'arrive',{'id':r2['ID']})
    assert len(rpc(service,'attentes',role='medecin')['items'])==2
    rpc(service,'claim',{'id':arr['ID']},'medecin')
    with pytest.raises(Refus,match='reservee'):rpc(service,'claim',{'id':arr['ID']},'autre')
    assert len(rpc(service,'attentes',role='medecin')['items'])==1


def test_creneau_overlap_meme_patient(service):
    _,pat,_,_=parcours(service)
    with pytest.raises(Refus,match='occupe'):rpc(service,'table.add',{'genre':'RDV','data':{'PatientID':pat['ID'],'Date':'12/09/2026','Heure':'10:10','DureeMin':'15','Statut':'Prevu'}})


def test_reservation_et_reprise_brouillon(service):
    _,_,_,arr=parcours(service);rpc(service,'claim',{'id':arr['ID']},'medecin')
    path=service.documents.root/'brouillon.docx';path.write_bytes(b'test')
    rpc(service,'draft',{'id':arr['ID'],'path':r'\\NAS-FICTIF\Cabinet\brouillon.docx'},'medecin')
    assert rpc(service,'reprises',role='medecin')['items'][0]['CheminBrouillon'].endswith('brouillon.docx')
    with pytest.raises(Refus,match='Brouillon'):rpc(service,'release',{'id':arr['ID']},'medecin')
    with pytest.raises(Refus):rpc(service,'draft',{'id':arr['ID'],'path':str(path)},'autre')


def test_annulation_reservation_bloquee(service):
    _,_,rdv,arr=parcours(service);rpc(service,'claim',{'id':arr['ID']},'medecin')
    with pytest.raises(Refus):rpc(service,'cancel_arrival',{'id':rdv['ID']})
    rpc(service,'release',{'id':arr['ID']},'medecin')
    rpc(service,'cancel_arrival',{'id':rdv['ID']})
    assert rpc(service,'attentes',role='medecin')['items']==[]


def test_publication_echec_pdf_pas_de_file(service):
    arr,pat,p=publication(service);(service.documents.root/'essai.pdf').unlink()
    with pytest.raises(Refus):rpc(service,'publish',p,'medecin')
    assert rpc(service,'publications')['items']==[]
    with service.connexion() as db:assert db.execute('SELECT etat FROM consultations WHERE id=%s',(arr['ID'],)).fetchone()['etat']=='encours'


def test_publication_identite_relecture_destinataire(service):
    _,_,p=publication(service)
    with pytest.raises(Refus,match='Relecture'):rpc(service,'publish',dict(p,Relu=False),'medecin')
    with pytest.raises(Refus,match='Identite'):rpc(service,'publish',dict(p,Patient_DDN='02/01/1980'),'medecin')
    out=rpc(service,'publish',p,'medecin')
    assert out['CheminDocx'].startswith(r'\\NAS-FICTIF\Cabinet\Documents')
    assert len(out['sha_docx'])==64


def test_facturation_idempotence_revision_et_traitement_atomique(service):
    arr,pat,p=publication(service);pub=rpc(service,'publish',p,'medecin')
    args={'id':arr['ID'],'publication_id':p['PublicationID'],'lignes':[ligne(arr,pat)]}
    with pytest.raises(Refus,match='actes'):rpc(service,'ack',{'id':pub['PublicationID']})
    assert rpc(service,'bill',args)['ajoute'] is True
    assert rpc(service,'bill',args)['ajoute'] is False
    changed={'id':arr['ID'],'publication_id':p['PublicationID'],'lignes':[dict(ligne(arr,pat),Montant='13.00')]}
    assert rpc(service,'bill',changed)['selection_differente'] is True
    assert rpc(service,'journal.read')['items'][0]['Montant']=='12.30'
    attempt=rpc(service,'print.request',{'id':arr['ID'],'reimpression_confirmee':False})
    rpc(service,'printed',{'id':arr['ID'],'tentative':attempt['tentative'],'confirmee':True})
    rpc(service,'ack',{'id':pub['PublicationID']})
    assert rpc(service,'publications')['items']==[]
    assert rpc(service,'record.get',{'genre':'RDV','id':arr['RdvID']})['Statut']=='Honore'
    assert rpc(service,'journal.read')['items'][0]['FeuilleSoinsImprimee']=='O'


def test_correspondants_ambigus(service):
    for prenom in ['A','B']:rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'HOMONYME','Prenom':prenom,'CleDestination':'MEME'}})
    with pytest.raises(Refus,match='ambigu'):rpc(service,'correspondent.resolve',{'cle':'MEME'},'medecin')


def test_revision_publication_remplace_file(service):
    arr,_,p=publication(service);a=rpc(service,'publish',p,'medecin')
    b=rpc(service,'publish',dict(p,PublicationID=uuid.uuid4().hex),'medecin')
    items=rpc(service,'publications')['items'];assert len(items)==1 and items[0]['PublicationID']==b['PublicationID']
    with pytest.raises(Refus,match='nouvelle version'):rpc(service,'ack',{'id':a['PublicationID']})


def test_dictionnaire_unicode(service):
    p={'genre':'MEDICAMENTS','texte':'Médicament fictif'}
    a=rpc(service,'dictionary.add',p,'medecin');b=rpc(service,'dictionary.add',p)
    assert a['ID']==b['ID']
    assert rpc(service,'dictionary.read',{'genre':'MEDICAMENTS'})['items'][0]['Terme']=='Médicament fictif'


def test_publication_id_reutilise_different(service):
    _,_,p=publication(service)
    first=rpc(service,'publish',p,'medecin')
    assert rpc(service,'publish',p,'medecin')==first
    with pytest.raises(Refus,match='contenu different'):rpc(service,'publish',dict(p,DateActe='11/09/2026'),'medecin')


def test_tarif_altere_refuse_avant_facturation(service):
    arr,pat,p=publication(service);rpc(service,'publish',p,'medecin')
    with pytest.raises(Refus,match='nomenclature'):rpc(service,'bill',{'id':arr['ID'],'publication_id':p['PublicationID'],'lignes':[dict(ligne(arr,pat),Montant='1.00')]})
    assert rpc(service,'journal.read')['items']==[]


def test_identite_archivee_et_pagination_journal(service):
    arr,pat,p=publication(service);pub=rpc(service,'publish',dict(p,Nom='NE DOIT PAS REMPLACER IDENTITE'),'medecin')
    assert pub['Nom']==pat['Nom']
    rpc(service,'bill',{'id':arr['ID'],'publication_id':p['PublicationID'],'lignes':[ligne(arr,pat)]})
    out=rpc(service,'journal.read',{'limit':1,'id':arr['ID']})
    assert out['items'][0]['Nom']==pat['Nom'] and out['next'] is None
    rpc(service,'payment',{'id':arr['ID'],'date':'12/09/2026','mode':'CB','empreinte':rpc(service,'billing.get',{'id':arr['ID']})['empreinte']})
    assert rpc(service,'journal.read')['items'][0]['Paye']=='O'


def test_un_seul_medecin_reserve_en_concurrence(service):
    if os.getenv('CABINET_TEST_PGLITE')=='1':pytest.skip('Concurrence exige PostgreSQL natif ; PGlite multiplexe une connexion.')
    from concurrent.futures import ThreadPoolExecutor
    from threading import Barrier
    _,_,_,arr=parcours(service);barrier=Barrier(2)
    def reserve(role):
        barrier.wait(timeout=5)
        try:rpc(service,'claim',{'id':arr['ID']},role);return 'acquis'
        except Refus:return 'refuse'
    with ThreadPoolExecutor(max_workers=2) as pool:
        assert sorted(pool.map(reserve,['medecin','autre']))==['acquis','refuse']


def test_migration_import_atomique_et_reexecution(service,tmp_path):
    from cabinet.migration import appliquer
    plan={'ressources':[{'genre':'CORRESPONDANTS','id':'C-TEST','data':{'ID':'C-TEST','Nom':'FICTIF'}}], 'historique':[], 'erreurs':[], 'avertissements':[]}
    assert appliquer(service,plan)=='importe'
    assert appliquer(service,plan)=='deja importe'
    modified={**plan,'avertissements':['different']}
    with pytest.raises(Refus,match='contient deja'):appliquer(service,modified)


def test_refus_statut_arrive_sans_transition(service):
    _,pat,_,_=parcours(service)
    with pytest.raises(Refus,match='nouveau rendez-vous'):rpc(service,'table.add',{'genre':'RDV','data':{'PatientID':pat['ID'],'Date':'12/09/2026','Heure':'13:00','DureeMin':'15','Statut':'Arrive'}})


def _u2c_docx(patient, consultation_id, principal, annexe, annexe_id):
    ns='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    settings=ET.Element(ns+'settings'); variables=ET.SubElement(settings,ns+'docVars')
    values={'PatientID':patient['ID'],'ConsultationID':consultation_id,
            'Patient_Nom':patient['Nom'],'Patient_Prenom':patient['Prenom'],
            'Patient_DDN':patient['DDN'],'Patient_Sexe':patient['Sexe'],
            'AnnexeDestinataire_001':annexe_id}
    for key,value in values.items():
        ET.SubElement(variables,ns+'docVar',{ns+'name':key,ns+'val':str(value)})
    document=ET.Element(ns+'document');body=ET.SubElement(document,ns+'body')
    for ident,name,text in [('1','DESTINATAIRE',principal),('2','U2ANN_DEST_001',annexe)]:
        p=ET.SubElement(body,ns+'p')
        ET.SubElement(p,ns+'bookmarkStart',{ns+'id':ident,ns+'name':name})
        ET.SubElement(ET.SubElement(p,ns+'r'),ns+'t').text=text
        ET.SubElement(p,ns+'bookmarkEnd',{ns+'id':ident})
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/settings.xml',ET.tostring(settings))
        z.writestr('word/document.xml',ET.tostring(document))
    return out.getvalue()


def test_u2c_selection_arrivee_refuse_identite_perimee(service):
    _,pat,_,arr=parcours(service)
    before=rpc(service,'attentes',role='medecin')['items']
    item=next(x for x in before if x['ConsultationID']==arr['ID'])
    old=item['RevisionSelection']
    current=rpc(service,'record.get',{'genre':'PATIENTS','id':pat['ID']})
    updated=rpc(service,'table.update',{'genre':'PATIENTS','data':dict(current,Prenom='MODIFIE')})
    with pytest.raises(Refus,match='Identite modifiee'):
        rpc(service,'claim',{'id':arr['ID'],'selection':old},'medecin')
    after=rpc(service,'attentes',role='medecin')['items']
    fresh=next(x for x in after if x['ConsultationID']==arr['ID'])
    assert fresh['RevisionSelection']!=old
    claimed=rpc(service,'claim',{'id':arr['ID'],'selection':fresh['RevisionSelection']},'medecin')
    assert claimed['Prenom']==updated['Prenom']=='MODIFIE'


def test_u2c_revision_secretariat_preserve_original_et_facturation_unique(service,monkeypatch):
    if os.name == 'nt':
        def conserver(data,extension):
            sha=hashlib.sha256(data).hexdigest();service.documents.objects.mkdir(parents=True,exist_ok=True)
            target=service.documents.objects/(sha+extension);target.write_bytes(data)
            return str(service.documents.unc/'Documents'/target.name),sha
        monkeypatch.setattr(service.documents,'_conserver_octets',conserver)
    arr,pat,p=publication(service)
    main=rpc(service,'record.get',{'genre':'CORRESPONDANTS','id':p['DestinataireID']})
    ann_a=rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{
        'ID':'C-ANN-A','Nom':'ANNEXE A','Actif':'1','AValider':'0',
        'BlocDestinataire':'Docteur ANNEXE A\n1 RUE A'}})
    ann_b=rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{
        'ID':'C-ANN-B','Nom':'ANNEXE B','Actif':'1','AValider':'0',
        'BlocDestinataire':'Docteur ANNEXE B\n2 RUE B'}})
    source=service.documents.root/'essai.docx'
    source.write_bytes(_u2c_docx(pat,arr['ID'],main['BlocDestinataire'],
                                  ann_a['BlocDestinataire'],ann_a['ID']))
    pub=rpc(service,'publish',p,'medecin')
    revid=uuid.uuid4().hex
    folder=service.documents.root/'Patients'/'_EditionsSecretariat'/revid
    folder.mkdir(parents=True)
    rdoc=folder/'revision.docx';rpdf=folder/'revision.pdf'
    rdoc.write_bytes(_u2c_docx(pat,arr['ID'],main['BlocDestinataire'],
                               ann_b['BlocDestinataire'],ann_b['ID']))
    rpdf.write_bytes(b'%PDF-1.4\n% REVISION FICTIVE U2C')

    args={'source_id':p['PublicationID'],'source_sha':pub['sha_docx'],
          'revision_id':revid,'patient_id':pat['ID'],
          'docx':str(service.documents.unc/'Patients'/'_EditionsSecretariat'/revid/'revision.docx'),
          'pdf':str(service.documents.unc/'Patients'/'_EditionsSecretariat'/revid/'revision.pdf'),
          'sha_docx':hashlib.sha256(rdoc.read_bytes()).hexdigest(),
          'sha_pdf':hashlib.sha256(rpdf.read_bytes()).hexdigest()}
    revision=rpc(service,'publication.revise',args)
    assert revision['PublicationMedicaleID']==p['PublicationID']
    assert revision['Annexes']==[{'Numero':'001','DestinataireID':ann_b['ID']}]
    with pytest.raises(Refus):
        rpc(service,'publication.revise',dict(args,revision_id=uuid.uuid4().hex))
    selection={'id':arr['ID'],'publication_id':revid,'lignes':[ligne(arr,pat)]}
    assert rpc(service,'bill',selection)['ajoute'] is True
    assert rpc(service,'bill',selection)['ajoute'] is False
    with pytest.raises(Refus):
        rpc(service,'bill',dict(selection,publication_id=p['PublicationID']))
    with service.connexion() as db:
        rows=db.execute('SELECT id,etat FROM publications WHERE consultation_id=%s ORDER BY cree_le,id',
                        (arr['ID'],)).fetchall()
        count=db.execute('SELECT count(*) AS n FROM seances WHERE id=%s',(arr['ID'],)).fetchone()['n']
    assert [(x['id'],x['etat']) for x in rows]==[(p['PublicationID'],'remplacee'),(revid,'a_traiter')]
    assert count==1
