from pathlib import Path
from collections import Counter
import openpyxl
import pytest
from cabinet.domain import Refus
from cabinet.migration import appliquer, construire_plan
from test_service import service, rpc


def workbook(path,sheets):
    path.parent.mkdir(parents=True,exist_ok=True)
    wb=openpyxl.Workbook();wb.remove(wb.active)
    for name,rows in sheets.items():
        ws=wb.create_sheet(name)
        for row in rows:ws.append(row)
    wb.save(path);wb.close()


def test_specialiste_deux_types_un_identifiant(tmp_path):
    workbook(tmp_path/'Base/base_travail_correspondants_v1.xlsx',{
        'Specialistes':[['ID','Nom','PrenomOuInitiale','TypeExamen'],['S1','FICTIF','Essai','TEST1']],
        'Specialistes_ParType':[['ID_Specialiste','TypeExamen','CleRegroupement'],['S1','TEST1','K1'],['S1','TEST2','K2']]})
    plan=construire_plan(tmp_path)
    assert not plan['erreurs']
    cors=[x['data'] for x in plan['ressources'] if x['genre']=='CORRESPONDANTS']
    assert len(cors)==1 and cors[0]['Prenom']=='Essai'
    assert set(cors[0]['TypesExamen'].split(';'))=={'TEST1','TEST2'}
    assert set(cors[0]['ClesDestination'].split(';'))=={'K1','K2'}


def test_identite_incomplete_et_file_active(tmp_path):
    workbook(tmp_path/'Base/Patients.xlsx',{'PATIENTS':[['ID','Nom','Prenom','DDN','Sexe'],['P1','FICTIF','Essai','31/02/2000','']]})
    folder=tmp_path/'Echange/AEnvoyer';folder.mkdir(parents=True);(folder/'test.txt').write_text('test')
    assert len(construire_plan(tmp_path)['erreurs'])>=2


def test_ressources_initiales_livrables():
    root=Path(__file__).resolve().parents[2]/'DonneesInitiales'
    plan=construire_plan(root)
    assert not plan['erreurs']
    counts=Counter(x['genre'] for x in plan['ressources'])
    assert counts['ACTES']==6 and counts['CORRESPONDANTS']>200 and counts['PATIENTS']==0


def ligne_historique(**changes):
    # Identite fictive figee volontairement differente de la fiche courante.
    return dict({'SeanceID':'S1','PatientID':'P1','Date':'12/09/2020',
                 'Nom':'FICTIF ANCIEN','Prenom':'Historique','DDN':'01/01/1980',
                 'NIR':'180010100000192','CodeActe':'ANCIEN','Montant':'12,30',
                 'Paye':'N','TiersPayant':'N','FeuilleSoinsImprimee':'N'},**changes)


def historique(root,lines):
    workbook(root/'Base/Patients.xlsx',{'PATIENTS':[
        ['ID','Nom','Prenom','DDN','Sexe','NIR'],
        ['P1','FICTIF ACTUEL','Courant','02/01/1980','M','180010100000192']]})
    headers=list(dict.fromkeys(key for line in lines for key in line))
    workbook(root/'Actes/Journal_2020.xlsx',{'JOURNAL':[
        headers,*[[line.get(key,'') for key in headers] for line in lines]]})
    return construire_plan(root)


def test_historique_valide_sans_champs_nouveaux_conserve_son_identite(tmp_path):
    original=ligne_historique()
    plan=historique(tmp_path,[original])
    assert not plan['erreurs']
    saved=plan['historique'][0]['lignes'][0]
    assert saved==dict(original,SeanceID='H2020_S1')
    assert 'CodeCerfa' not in saved and 'PayeurReglement' not in saved
    assert 'AssureNom' not in saved


@pytest.mark.parametrize('champ',[
    'Nom','Prenom','DDN','CodeActe','Date','Montant','TiersPayant','Paye','FeuilleSoinsImprimee'])
@pytest.mark.parametrize('absent',[False,True])
def test_historique_incomplet_bloque_sans_deduire_depuis_patient(tmp_path,champ,absent):
    original=ligne_historique()
    if absent:original.pop(champ)
    else:original[champ]=''
    plan=historique(tmp_path,[original])
    assert any(champ in error for error in plan['erreurs'])
    assert len(plan['historique'])==1  # anomalie visible, jamais ignoree
    assert plan['historique'][0]['lignes'][0]==dict(original,SeanceID='H2020_S1')
    # Le refus doit preceder toute connexion, donc toute ecriture en base.
    with pytest.raises(Refus,match='simulation'):appliquer(None,plan)


@pytest.mark.parametrize('changes,champ',[
    ({'Date':'31/02/2020'},'Date'),({'DDN':'31/02/1980'},'DDN'),
    ({'NIR':'180010100000100'},'NIR'),({'Montant':'invalide'},'Montant'),
    ({'TiersPayant':'oui'},'TiersPayant'),({'Paye':'0'},'Paye'),
    ({'FeuilleSoinsImprimee':'inconnue'},'FeuilleSoinsImprimee'),
    ({'FeuilleSoinsImprimee':'0'},'FeuilleSoinsImprimee'),
    ({'ModePaiement':'inconnu'},'ModePaiement'),
    ({'Paye':'O','DateEncaissement':'12/09/2020'},'ModePaiement'),
    ({'Paye':'O','ModePaiement':'CB'},'DateEncaissement'),
    ({'Paye':'O','ModePaiement':'CB','DateEncaissement':'31/02/2020'},'DateEncaissement'),
    ({'DateEncaissement':'12/09/2020'},'DateEncaissement'),
    ({'Paye':'O','ModePaiement':'CB','DateEncaissement':'12/09/2020','PayeurReglement':'Organisme'},'PayeurReglement'),
])
def test_historique_donnees_comptables_invalides_signalees(tmp_path,changes,champ):
    plan=historique(tmp_path,[ligne_historique(**changes)])
    assert any(champ in error for error in plan['erreurs'])


def test_historique_assure_distinct_accepte_sans_nir_patient(tmp_path):
    original=ligne_historique(NIR='',AssureNom='FICTIF ASSURE',AssurePrenom='Essai',
                             AssureDDN='01/01/1980',AssureNIR='180010100000192')
    plan=historique(tmp_path,[original])
    assert not plan['erreurs']
    assert plan['historique'][0]['lignes'][0]==dict(original,SeanceID='H2020_S1')


@pytest.mark.parametrize('champ',['AssureNom','AssurePrenom','AssureDDN'])
def test_historique_assure_distinct_incomplet_refuse(tmp_path,champ):
    original=ligne_historique(AssureNom='FICTIF ASSURE',AssurePrenom='Essai',
                             AssureDDN='01/01/1980',AssureNIR='180010100000192')
    original.pop(champ)
    plan=historique(tmp_path,[original])
    assert any(champ in error for error in plan['erreurs'])


@pytest.mark.parametrize('naissance',['01/01/1980','31/02/1980'])
def test_historique_assure_ddn_seule_ne_peut_pas_etre_ignoree(tmp_path,naissance):
    original=ligne_historique(AssureDDN=naissance)
    plan=historique(tmp_path,[original])
    for champ in ('AssureNom','AssurePrenom'):
        assert any(champ in error for error in plan['erreurs'])
    if naissance=='31/02/1980':
        assert any('AssureDDN' in error for error in plan['erreurs'])
    assert plan['historique'][0]['lignes'][0]==dict(original,SeanceID='H2020_S1')
    with pytest.raises(Refus,match='simulation'):appliquer(None,plan)


@pytest.mark.parametrize('changes,attendu',[
    ({'Nom':'AUTRE FICTIF'},'Identites figees incoherentes'),
    ({'Date':'13/09/2020'},'Dates figees incoherentes'),
    ({'FeuilleSoinsImprimee':'O'},'Etats FeuilleSoinsImprimee incoherents'),
])
def test_historique_snapshot_incoherent_entre_lignes_refuse(tmp_path,changes,attendu):
    plan=historique(tmp_path,[ligne_historique(),ligne_historique(CodeActe='SECOND',**changes)])
    assert any(attendu in error for error in plan['erreurs'])
    assert len(plan['historique'][0]['lignes'])==2


@pytest.mark.parametrize('tiers',['O','N'])
def test_historique_deja_encaisse_accepte_meme_en_tiers_payant(tmp_path,tiers):
    plan=historique(tmp_path,[ligne_historique(
        TiersPayant=tiers,Paye='O',ModePaiement='Virement',DateEncaissement='14/09/2020')])
    assert not plan['erreurs']


def test_historique_plus_de_quatre_lignes_conserve_integralement_avec_alerte(tmp_path):
    original=[ligne_historique(CodeActe='ANCIEN'+str(index)) for index in range(5)]
    plan=historique(tmp_path,original)
    assert not plan['erreurs']
    assert len(plan['historique'][0]['lignes'])==5
    assert any('reimpression CERFA non prise en charge' in warning for warning in plan['avertissements'])


@pytest.mark.parametrize('printed,etat',[('O','confirmee'),('N','actes_enregistres')])
def test_import_historique_aligne_impression_et_encaissement(service,tmp_path,printed,etat):
    plan=historique(tmp_path,[ligne_historique(FeuilleSoinsImprimee=printed)])
    assert appliquer(service,plan)=='importe'
    assert appliquer(service,plan)=='deja importe'
    saved=rpc(service,'billing.get',{'id':'H2020_S1'})
    assert saved['impression_etat']==etat
    assert saved['lignes'][0]['Nom']=='FICTIF ANCIEN'
    assert rpc(service,'journal.read',{'id':'H2020_S1'})['items'][0]['FeuilleSoinsImprimee']==printed
    if printed=='O':
        with pytest.raises(Refus,match='reimpression explicite'):
            rpc(service,'print.request',{'id':'H2020_S1','reimpression_confirmee':False})
    result=rpc(service,'payment',{'id':'H2020_S1','empreinte':saved['empreinte'],
                                 'date':'16/09/2020','mode':'CB'})
    assert result['lignes'][0]['Paye']=='O'
    assert result['lignes'][0]['Nom']=='FICTIF ANCIEN'


@pytest.mark.parametrize('etats',[[None],['inconnue'],['O','N']])
def test_import_etat_papier_absent_inconnu_ou_mixte_ne_cree_aucune_seance(service,tmp_path,etats):
    lines=[]
    for index,etat in enumerate(etats):
        line=ligne_historique(CodeActe='ANCIEN'+str(index))
        if etat is None:line.pop('FeuilleSoinsImprimee')
        else:line['FeuilleSoinsImprimee']=etat
        lines.append(line)
    plan=historique(tmp_path,lines)
    assert any('FeuilleSoinsImprimee' in error for error in plan['erreurs'])
    with pytest.raises(Refus,match='simulation'):appliquer(service,plan)
    with service.connexion() as db:
        assert db.execute('SELECT count(*) AS n FROM seances').fetchone()['n']==0
        assert db.execute('SELECT count(*) AS n FROM consultations').fetchone()['n']==0
    with pytest.raises(Refus,match='introuvable'):
        rpc(service,'print.request',{'id':'H2020_S1','reimpression_confirmee':False})


@pytest.mark.parametrize('absent',[False,True])
@pytest.mark.parametrize('assure_distinct',[False,True])
@pytest.mark.parametrize('imprime',['O','N'])
def test_historique_sans_nir_conserve_ses_valeurs_sans_substitution(tmp_path,absent,assure_distinct,imprime):
    original=ligne_historique(NIR='',FeuilleSoinsImprimee=imprime)
    if absent:original.pop('NIR')
    if assure_distinct:
        original.update(AssureNom='ASSURE FICTIF',AssurePrenom='Essai',AssureDDN='02/02/1970')
        if not absent:original['AssureNIR']=''
    plan=historique(tmp_path,[original])
    assert not plan['erreurs']
    saved=plan['historique'][0]['lignes'][0]
    assert saved==dict(original,SeanceID='H2020_S1')
    assert not saved.get('NIR') and not saved.get('AssureNIR')
    # La fiche courante contient pourtant un NIR : aucun ajout a l'historique.
    if absent:assert 'NIR' not in saved and 'AssureNIR' not in saved


@pytest.mark.parametrize('champ',['NIR','AssureNIR'])
def test_historique_nir_present_invalide_reste_refuse(tmp_path,champ):
    original=ligne_historique(NIR='',AssureNom='ASSURE FICTIF',AssurePrenom='Essai',
                             AssureDDN='02/02/1970',AssureNIR='')
    original[champ]='180010100000100'
    plan=historique(tmp_path,[original])
    assert any(champ in error for error in plan['erreurs'])
    with pytest.raises(Refus,match='simulation'):appliquer(None,plan)
