import copy
import logging
import uuid
import pytest
import psycopg
from psycopg.types.json import Jsonb
from fastapi.testclient import TestClient
from cabinet.api import create_app
from cabinet.domain import Refus, comparer_clinique
from cabinet.contract import validate
from cabinet.migration_u1 import migrer
from test_service import service, rpc, parcours, publication, ligne


@pytest.mark.parametrize('op,p', [
    ('bill', {'id':'x','publication_id':'y','lignes':['invalide']}),
    ('bill', {'id':'x','publication_id':'y','lignes':[None]}),
    ('table.add', {'genre':'PATIENTS','data':{'Nom': {'secret':'ne pas copier'}}}),
    ('table.read', {'genre':'PATIENTS','offset':True}),
    ('table.read', {'genre':'PATIENTS','limit':0}),
    ('clinical.compare', {'source':None,'resultat':'texte'}),
    ('dictionary.add', {'genre':'MEDICAMENTS','texte':['test']}),
    ('draft', {'id':'x','path':3}),
])
def test_contract_bad_types_before_transaction(op,p):
    with pytest.raises(Refus) as error: validate(op,p)
    assert error.value.status == 422


@pytest.mark.parametrize('left,right,expected', [
    ('TEST 5µg', 'TEST 5', 'nombres_et_unites'),
    ('absence de douleur, dyspnee presente', 'douleur presente, absence de dyspnee', 'contexte_negations'),
    ('[[PATIENT]] TEST', 'TEST', 'identifiants_masques'),
])
def test_clinical_regressions(left,right,expected):
    result = comparer_clinique(left,right)
    assert expected in {d['controle'] for d in result['differences']}
    assert result['relecture_obligatoire'] is True


def test_dose_spacing_is_not_a_clinical_change():
    assert comparer_clinique('TEST 5 mg', 'TEST 5mg')['differences'] == []


def test_command_recovery_is_scoped_to_account(service):
    rid=uuid.uuid4().hex
    saved=rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'FICTIF'}},rid=rid)
    assert rpc(service,'command.result',{'id':rid}) == {'trouve':True,'resultat':saved}
    assert rpc(service,'command.result',{'id':rid},'medecin') == {'trouve':False,'resultat':None}


def test_existing_imported_act_editable_and_finances_preserved(service):
    old={'ID':'ACTE_FICTIF','Code':'ACTE_FICTIF','LibelleCourt':'Libelle historique','Tarif':'12.30',
         'Depassement':'4.50','Actif':'1'}
    with service.connexion() as db:
        db.execute("INSERT INTO ressources(genre,id,donnees) VALUES ('ACTES',%s,%s)", ('ACTE_FICTIF',Jsonb(old)))
    simulation=migrer(service)
    assert simulation['modifications']==1 and not simulation['applique']
    with service.connexion() as db:
        assert db.execute("SELECT donnees FROM ressources").fetchone()['donnees']==old
    migrer(service,simulation['empreinte'])
    repeat=migrer(service); assert repeat['modifications']==0
    migrer(service,repeat['empreinte'])
    a=rpc(service,'record.get',{'genre':'ACTES','id':'ACTE_FICTIF'})
    assert a['Libelle']==a['LibelleCourt']=='Libelle historique' and a['Depassement']=='4.50'
    updated=rpc(service,'table.update',{'genre':'ACTES','data':dict(a,Libelle='Nouveau libelle')})
    assert updated['LibelleCourt']=='Nouveau libelle' and updated['Depassement']=='4.50'
    with pytest.raises(Refus):rpc(service,'table.update',{'genre':'ACTES','data':dict(updated,Code='AUTRE')})
    with service.connexion() as db:
        assert db.execute('SELECT avant FROM migrations_ressources').fetchone()['avant']==old
        assert db.execute('SELECT count(*) n FROM migrations_ressources').fetchone()['n']==1


def test_migration_conflict_is_all_or_nothing(service):
    with service.connexion() as db:
        for code, a, b in [('A','identique','identique'), ('B','premier','second')]:
            db.execute("INSERT INTO ressources(genre,id,donnees) VALUES ('ACTES',%s,%s)",
                       (code,Jsonb({'ID':code,'Code':code,'Libelle':a,'LibelleCourt':b,'Tarif':'1'})))
    plan=migrer(service); assert len(plan['conflits'])==1
    with pytest.raises(Refus):migrer(service,plan['empreinte'])
    with service.connexion() as db:
        assert db.execute('SELECT sum(revision) n FROM ressources').fetchone()['n']==2
        assert db.execute('SELECT count(*) n FROM migrations_ressources').fetchone()['n']==0


def test_publish_validation_precedes_archive_copy(service,monkeypatch):
    _,_,p=publication(service)
    def forbidden(*a):pytest.fail('Ne doit pas commencer les copies')
    monkeypatch.setattr(service.documents,'conserver_publication',forbidden)
    for values in (dict(p,DateActe='11/09/2026'),dict(p,DestinataireID='A_COMPLETER'),dict(p,CheminPdf=42)):
        with pytest.raises(Refus):rpc(service,'publish',values,'medecin')


def test_stale_publication_cannot_create_a_bill(service):
    arr,pat,p=publication(service);rpc(service,'publish',p,'medecin')
    rpc(service,'publish',dict(p,PublicationID=uuid.uuid4().hex),'medecin')
    with pytest.raises(Refus,match='nouvelle version'):
        rpc(service,'bill',{'id':arr['ID'],'publication_id':p['PublicationID'],'lignes':[ligne(arr,pat)]})
    assert rpc(service,'journal.read')['items']==[]


def test_business_comparison_and_authoritative_lines(service):
    arr,pat,p=publication(service);rpc(service,'publish',p,'medecin')
    rpc(service,'table.add',{'genre':'ACTES','data':{'Code':'SECOND','Tarif':'2'}})
    lines=[ligne(arr,pat),dict(ligne(arr,pat),CodeActe='SECOND',Montant='2')]
    args={'id':arr['ID'],'publication_id':p['PublicationID'],'lignes':lines}
    saved=rpc(service,'bill',args)
    reordered=[dict(x,Notes='annotation sans effet sur acte') for x in reversed(lines)]
    again=rpc(service,'bill',dict(args,lignes=reordered))
    assert not again['selection_differente'] and again['lignes']==saved['lignes']
    changed=rpc(service,'bill',dict(args,lignes=[dict(lines[0],Montant='1')]))
    assert changed['selection_differente'] and changed['lignes']==saved['lignes']


def test_print_uncertainty_blocks_ack_and_automatic_reprint(service):
    arr,pat,p=publication(service);rpc(service,'publish',p,'medecin')
    rpc(service,'bill',{'id':arr['ID'],'publication_id':p['PublicationID'],'lignes':[ligne(arr,pat)]})
    attempt=rpc(service,'print.request',{'id':arr['ID'],'reimpression_confirmee':False})
    with pytest.raises(Refus,match='inconnu'):rpc(service,'ack',{'id':p['PublicationID']})
    with pytest.raises(Refus):rpc(service,'print.request',{'id':arr['ID'],'reimpression_confirmee':False})
    with pytest.raises(Refus):rpc(service,'printed',{'id':arr['ID'],'tentative':'autre','confirmee':True})
    rpc(service,'printed',{'id':arr['ID'],'tentative':attempt['tentative'],'confirmee':True})
    rpc(service,'ack',{'id':p['PublicationID']})


def test_payment_is_audited_and_cannot_overwrite(service):
    arr,pat,p=publication(service);rpc(service,'publish',p,'medecin')
    saved=rpc(service,'bill',{'id':arr['ID'],'publication_id':p['PublicationID'],'lignes':[ligne(arr,pat)]})
    args={'id':arr['ID'],'mode':'CB','date':'12/09/2026','empreinte':saved['empreinte']}
    for bad in (dict(args,mode='INCONNU'),dict(args,empreinte='obsolete')):
        with pytest.raises(Refus):rpc(service,'payment',bad)
    rpc(service,'payment',args)
    with pytest.raises(Refus):rpc(service,'payment',args)
    with service.connexion() as db:
        audit=db.execute('SELECT avant,apres FROM reglements_audit').fetchall()
        assert len(audit)==1 and audit[0]['avant'][0]['Paye']=='N' and audit[0]['apres'][0]['Paye']=='O'


def test_absence_atomic_and_reactivation_explicit(service):
    _,_,rdv,arr=parcours(service)
    current=rpc(service,'record.get',{'genre':'RDV','id':rdv['ID']})
    out=rpc(service,'agenda.status',{'id':rdv['ID'],'revision':current['_revision'],'statut':'Absent'})
    assert out['Statut']=='Absent' and rpc(service,'attentes',role='medecin')['items']==[]
    with pytest.raises(Refus,match='nouveau rendez-vous'):
        rpc(service,'agenda.status',{'id':rdv['ID'],'revision':out['_revision'],'statut':'Prevu'})
    with service.connexion() as db:
        assert db.execute('SELECT etat FROM consultations WHERE id=%s',(arr['ID'],)).fetchone()['etat']=='annule'


@pytest.mark.parametrize('initial', ['Absent','Annule'])
def test_unstarted_rdv_can_return_to_planned(service,initial):
    _,pat,_,_=parcours(service)
    rdv=rpc(service,'table.add',{'genre':'RDV','data':{'PatientID':pat['ID'],'Date':'12/09/2026','Heure':'13:00','DureeMin':'15','Statut':initial}})
    out=rpc(service,'agenda.status',{'id':rdv['ID'],'revision':rdv['_revision'],'statut':'Prevu'})
    assert out['Statut']=='Prevu'


@pytest.mark.parametrize('error,status,code', [
    (psycopg.OperationalError('SECRET_PHI'),503,'indisponible'),
    (psycopg.errors.UniqueViolation('SECRET_PHI'),409,'contrainte'),
    (psycopg.errors.UndefinedTable('SECRET_PHI'),500,'base_interne'),
])
def test_error_classification_and_private_diagnostics(service,monkeypatch,caplog,error,status,code):
    def fail(*a):raise error
    monkeypatch.setattr(service,'compte',fail)
    with TestClient(create_app(service)) as client, caplog.at_level(logging.WARNING):
        result=client.post('/v1/rpc',headers={'Authorization':'Bearer '+'SECRET_TOKEN'*4},json={'operation':'whoami'})
    assert result.status_code==status and result.json()['code']==code
    assert 'SECRET_' not in result.text and 'SECRET_' not in caplog.text
    assert result.json()['correlation'] in caplog.text


@pytest.mark.parametrize('data', [ {'Libelle': 'FICTIF', 'Tarif': '1'}, {'ID': 'X', 'Code': ' ', 'Tarif': '1'} ])
def test_new_act_requires_explicit_business_code(service, data):
    with pytest.raises(Refus, match='Code ACTES explicite'):
        rpc(service, 'table.add', {'genre': 'ACTES', 'data': data})
    assert rpc(service, 'table.read', {'genre': 'ACTES'})['items'] == []
