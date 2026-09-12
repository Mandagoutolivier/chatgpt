from pathlib import Path
from collections import Counter
import openpyxl
from cabinet.migration import construire_plan


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
