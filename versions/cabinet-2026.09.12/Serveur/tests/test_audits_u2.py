"""Regressions des audits U2 : donnees strictement fictives."""
import hashlib
import uuid
import zipfile
import pytest
from cabinet.domain import Refus
from cabinet.files import Documents
from cabinet.migration import construire_plan, texte, appliquer
from test_service import service, rpc, publication, ligne
from test_migration import workbook


def mixed_bill(s):
    arr, pat, data = publication(s)
    rpc(s, 'publish', data, 'medecin')
    acte=next(x for x in rpc(s,'table.read',{'genre':'ACTES'})['items'] if x['Code']=='TEST')
    rpc(s,'table.update',{'genre':'ACTES','data':dict(acte,CodeAssocie='TESTA',TarifAssocie='2.70')})
    rpc(s, 'table.add', {'genre':'ACTES','data':{'Code':'TIERS','Tarif':'7.20','CodeAssocie':'TIERSA','TarifAssocie':'0.80'}})
    lines = [ligne(arr, pat),
             dict(ligne(arr, pat), CodeActe='TESTA', Montant='2.70'),
             dict(ligne(arr, pat), CodeActe='TIERS', Montant='7.20', TiersPayant='O'),
             dict(ligne(arr, pat), CodeActe='TIERSA', Montant='0.80', TiersPayant='O')]
    rpc(s, 'bill', {'id':arr['ID'], 'publication_id':data['PublicationID'], 'lignes':lines})
    return rpc(s, 'billing.get', {'id':arr['ID']})


def payment(saved, **overrides):
    return dict(id=saved['ID'], date='12/09/2026', mode='CB', empreinte=saved['empreinte'], **overrides)


def test_mixed_payment_preserves_other_payer_and_retry(service):
    saved = mixed_bill(service)
    args = payment(saved, payeur='Patient', montant='15.00')
    request_id = uuid.uuid4().hex
    first = rpc(service, 'payment', args, rid=request_id)
    assert [x['Paye'] for x in first['lignes']] == ['O','O','N','N']
    assert first['lignes'][2:] == saved['lignes'][2:]
    assert rpc(service, 'payment', args, rid=request_id) == first
    with pytest.raises(Refus, match='rechargez'):
        rpc(service, 'payment', args)
    with pytest.raises(Refus, match='Aucune ligne'):
        rpc(service, 'payment', payment(first, payeur='Patient', montant='15.00'))
    last = rpc(service, 'payment', payment(first, payeur='Organisme', montant='8,00'))
    assert last['lignes'][:2] == first['lignes'][:2]
    assert [x['PayeurReglement'] for x in last['lignes']] == ['Patient','Patient','Organisme','Organisme']
    assert [x['Paye'] for x in last['lignes']] == ['O','O','O','O']
    journal=rpc(service,'journal.read',{'id':saved['ID']})['items']
    assert [(x['CodeActe'],x['Montant'],x['TiersPayant'],x['Paye'],x.get('PayeurReglement','')) for x in journal] == [
        ('TEST','12.30','N','O','Patient'),('TESTA','2.70','N','O','Patient'),
        ('TIERS','7.20','O','O','Organisme'),('TIERSA','0.80','O','O','Organisme')]
    with service.connexion() as db:
        assert db.execute('SELECT count(*) n FROM reglements_audit').fetchone()['n'] == 2


@pytest.mark.parametrize('fields', [{}, {'payeur':'Patient'}, {'montant':'12.30'},
    {'payeur':'Patient','montant':'23.00'}, {'payeur':'Organisme','montant':'1.00'},
    {'payeur':'Patient','montant':'12.300'}, {'payeur':'Inconnu','montant':'12.30'}])
def test_mixed_payment_refuses_ambiguous_or_partial_without_changes(service, fields):
    saved = mixed_bill(service)
    with pytest.raises(Refus): rpc(service, 'payment', payment(saved, **fields))
    assert rpc(service, 'billing.get', {'id':saved['ID']}) == saved
    with service.connexion() as db:
        assert db.execute('SELECT count(*) n FROM reglements_audit').fetchone()['n'] == 0


def test_date_filter_for_session_selection(service):
    saved = mixed_bill(service)
    assert {x['SeanceID'] for x in rpc(service,'journal.read',{'date':'12/09/2026'})['items']} == {saved['ID']}
    assert rpc(service,'journal.read',{'date':'11/09/2026'})['items'] == []
    with pytest.raises(Refus):rpc(service,'journal.read',{'date':'31/02/2026'})


def test_accent_search_literal_wildcards_and_pagination(service):
    names=['Éloïse','ELOISE','E\u0301loi\u0308se','EL_100%','AUTRE']
    rows=[rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':n}}) for n in names]
    args={'genre':'CORRESPONDANTS','q':'éLOÏSE','limit':1}
    found=[]; offset=0
    while offset is not None:
        page=rpc(service,'table.read',dict(args,offset=offset))
        found += [x['ID'] for x in page['items']]
        offset=page['next']
    assert sorted(found) == sorted(x['ID'] for x in rows[:3])
    for q in ('%', '_', 'el_100%'):
        assert [x['ID'] for x in rpc(service,'table.read',{'genre':'CORRESPONDANTS','q':q})['items']] == [rows[3]['ID']]
    assert rpc(service,'table.read',{'genre':'CORRESPONDANTS','q':'\\'})['items'] == []


def test_address_has_no_generated_blank_line(service):
    cor=rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'FICTIF','Adresse1':'ADRESSE TEST','CP':'00000','Ville':'TEST'}})
    assert cor['BlocDestinataire'].splitlines() == ['FICTIF','ADRESSE TEST','00000 TEST']


def test_publication_minimised_but_billing_identity_preserved(service):
    arr, pat, data=publication(service)
    base='2800100000001'; nir=base+str(97-int(base)%97).zfill(2)
    pat=rpc(service,'table.update',{'genre':'PATIENTS','data':dict(pat,NIR=nir,Adresse1='ADRESSE FICTIVE',Tel='0000000000')})
    pub=rpc(service,'publish',data,'medecin')
    assert pub['Nom']==pat['Nom'] and pub['Sexe']=='F'
    assert not {'NIR','AssureNIR','Adresse1','Tel','Patient_Nom'} & pub.keys()
    acte=next(x for x in rpc(service,'table.read',{'genre':'ACTES'})['items'] if x['Code']=='TEST')
    rpc(service,'table.update',{'genre':'ACTES','data':dict(acte,LibelleCerfa='TEST CERFA')})
    with pytest.raises(Refus,match='Libelle CERFA'):
        rpc(service,'bill',{'id':arr['ID'],'publication_id':data['PublicationID'],
                            'lignes':[dict(ligne(arr,pat),CodeCerfa='AUTRE')]})
    saved=rpc(service,'bill',{'id':arr['ID'],'publication_id':data['PublicationID'],'lignes':[ligne(arr,pat)]})
    assert saved['lignes'][0]['NIR']==nir
    assert saved['lignes'][0]['CodeCerfa']=='TEST CERFA'
    first_print=rpc(service,'print.request',{'id':arr['ID'],'reimpression_confirmee':False})
    rpc(service,'printed',{'id':arr['ID'],'tentative':first_print['tentative'],'confirmee':True})
    rpc(service,'table.update',{'genre':'PATIENTS','data':dict(pat,Nom='NOUVEAU NOM',Prenom='Nouveau',DDN='02/02/1982',NIR='')})
    acte=rpc(service,'record.get',{'genre':'ACTES','id':acte['ID']})
    rpc(service,'table.update',{'genre':'ACTES','data':dict(acte,LibelleCerfa='AUTRE LIBELLE')})
    frozen=rpc(service,'billing.get',{'id':arr['ID']})['lignes'][0]
    assert (frozen['Nom'],frozen['Prenom'],frozen['DDN'],frozen['NIR']) == (pat['Nom'],pat['Prenom'],pat['DDN'],nir)
    assert frozen['CodeCerfa']=='TEST CERFA'
    retry=rpc(service,'print.request',{'id':arr['ID'],'reimpression_confirmee':True})
    assert retry['tentative'] != first_print['tentative']


def files(root):
    with zipfile.ZipFile(root/'Lettre.docx','w') as z:z.writestr('word/document.xml','<document>FICTIF</document>')
    (root/'Lettre.pdf').write_bytes(b'%PDF-1.4\nFICTIF')
    return Documents(root,r'\\NAS\Cabinet')


def test_invalid_second_file_does_not_leave_first_archive(tmp_path):
    docs=files(tmp_path)
    (tmp_path/'Lettre.pdf').write_bytes(b'pas un PDF')
    with pytest.raises(Refus,match='PDF invalide'):
        docs.conserver_publication(r'\\NAS\Cabinet\Lettre.docx',r'\\NAS\Cabinet\Lettre.pdf')
    assert not docs.objects.exists()


def test_archive_uses_validated_bytes_even_if_source_changes(tmp_path,monkeypatch):
    docs=files(tmp_path); expected=(tmp_path/'Lettre.pdf').read_bytes()
    original=docs._conserver_octets
    def change(data,ext):
        (tmp_path/'Lettre.pdf').write_bytes(b'corrompu apres validation')
        return original(data,ext)
    monkeypatch.setattr(docs,'_conserver_octets',change)
    _, (path,digest)=docs.conserver_publication(r'\\NAS\Cabinet\LETTRE.DOCX',r'\\NAS\Cabinet\lettre.PDF')
    assert docs.resoudre(path).read_bytes()==expected
    assert digest==hashlib.sha256(expected).hexdigest()


def test_unc_case_resolution_and_ambiguity(tmp_path):
    folder=tmp_path/'Courriers';folder.mkdir();(folder/'Lettre.pdf').write_bytes(b'%PDF-1.4')
    docs=Documents(tmp_path,r'\\NAS\Cabinet')
    assert docs.resoudre(r'\\nas\CABINET\COURRIERS\lettre.PDF')==folder/'Lettre.pdf'
    (folder/'lettre.pdf').write_bytes(b'autre')
    with pytest.raises(Refus,match='ambigu'):docs.resoudre(r'\\NAS\Cabinet\Courriers\Lettre.pdf')


def test_unc_cannot_escape_through_symlink(tmp_path):
    root=tmp_path/'racine';root.mkdir();(root/'Lien').symlink_to(tmp_path,target_is_directory=True)
    docs=Documents(root,r'\\NAS\Cabinet')
    with pytest.raises(Refus,match='hors'):docs.resoudre(r'\\NAS\Cabinet\lien\secret.pdf')


@pytest.mark.parametrize('value,expected',[(1.0,'1'),(1,'1'),('001','001'),('1.0','1.0'),(True,'1'),(1.25,'1.25')])
def test_excel_identifiers_keep_strings_and_normalise_numeric(value,expected):
    assert texte(value,'ID')==expected
    assert texte(value,'PatientID')==expected


def legacy(root,sex='F',status='Prevu'):
    workbook(root/'Base/Patients.xlsx',{'PATIENTS':[['ID','Nom','Prenom','DDN','Sexe'],[1.0,'FICTIF','Essai','01/01/1980',sex]]})
    workbook(root/'Base/Agenda_2026.xlsx',{'RDV':[['ID','PatientID','Date','Heure','DureeMin','Statut'],[2.0,1.0,'12/09/2026','10:00',15.0,status]]})
    workbook(root/'Actes/Journal_2026.xlsx',{'JOURNAL':[['SeanceID','PatientID','Date','Montant'],[3.0,1.0,'12/09/2026','1.00']]})


def test_numeric_migration_links_and_status_validation(tmp_path):
    legacy(tmp_path)
    plan=construire_plan(tmp_path)
    assert not plan['erreurs']
    rdv=next(x for x in plan['ressources'] if x['genre']=='RDV')
    assert rdv['id']=='R2026_2' and rdv['data']['PatientID']=='1' and rdv['data']['DureeMin']=='15'
    assert plan['historique'][0]['patient_id']=='1'
    legacy(tmp_path,status='INCONNU')
    assert any('Statut agenda inconnu' in e for e in construire_plan(tmp_path)['erreurs'])


def test_invalid_patient_dependencies_block_import_without_discarding(service,tmp_path):
    legacy(tmp_path,sex='')
    plan=construire_plan(tmp_path)
    assert len(plan['ressources'])==2 and len(plan['historique'])==1
    assert any('Sexe' in e for e in plan['erreurs'])
    assert any('Rendez-vous R2026_2 lie' in e for e in plan['erreurs'])
    assert any('Seance historique H2026_3 liee' in e for e in plan['erreurs'])
    with pytest.raises(Refus,match='simulation'):appliquer(service,plan)
    assert rpc(service,'table.read',{'genre':'PATIENTS'})['items']==[]


def test_numeric_identifier_collision_is_reported(tmp_path):
    workbook(tmp_path/'Base/Patients.xlsx',{'PATIENTS':[['ID','Nom','Prenom','DDN','Sexe'],[1,'FICTIF A','Essai','01/01/1980','F'],['1','FICTIF B','Essai','01/01/1980','F']]})
    assert any('Identifiant repete' in e for e in construire_plan(tmp_path)['erreurs'])


def test_identity_changed_since_letter_blocks_first_bill(service):
    arr, pat, data=publication(service)
    rpc(service,'publish',data,'medecin')
    rpc(service,'table.update',{'genre':'PATIENTS','data':dict(pat,Nom='AUTRE NOM FICTIF')})
    with pytest.raises(Refus,match='Identite modifiee'):
        rpc(service,'bill',{'id':arr['ID'],'publication_id':data['PublicationID'],'lignes':[ligne(arr,pat)]})
    assert rpc(service,'journal.read')['items']==[]


def test_simulation_report_private_and_not_overwritten(tmp_path,monkeypatch,capsys):
    import json
    from cabinet.migration import main
    root=tmp_path/'source';root.mkdir()
    report=tmp_path/'rapport.json'
    monkeypatch.setattr('sys.argv',['migration','--racine',str(root),'--rapport',str(report)])
    main()
    assert report.stat().st_mode & 0o777 == 0o600
    result=json.loads(report.read_text())
    assert result['erreurs']==[] and 'ressources' in result
    before=report.read_bytes()
    with pytest.raises(FileExistsError):main()
    assert report.read_bytes()==before
    capsys.readouterr()
