from datetime import date
from pathlib import Path
import pytest
from cabinet.domain import date_fr, montant, patient_valide, Refus, chevauche, comparer_clinique, valider_nir
from cabinet.files import Documents

@pytest.mark.parametrize('value',['31/02/2026','29/02/2025','1/1/2020','2020-01-01'])
def test_dates_invalides(value):
    with pytest.raises(Refus):date_fr(value)

@pytest.mark.parametrize('value',['-1','12.345','1e3','NaN','12 euros'])
def test_montants_invalides(value):
    with pytest.raises(Refus):montant(value)


def test_montants_decimal():assert str(montant('12,30'))=='12.30'


def test_nir_controle():
    base='1800101001001';valid=base+f'{97-int(base)%97:02}'
    assert valider_nir(valid)==valid
    with pytest.raises(Refus):valider_nir(valid[:-2]+'00')


def test_modification_dose_negation():
    r=comparer_clinique('[[PATIENT]] sans douleur. TEST 5 mg.', '[[PATIENT]] douleur. TEST 10 mg.')
    assert {d['controle'] for d in r['differences']}=={'nombres_et_unites','negations','contexte_negations'}
    assert r['relecture_obligatoire'] is True


def test_document_hors_nas(tmp_path):
    d=Documents(tmp_path,r'\\DS224\CabinetCardio')
    for path in [r'C:\patient.docx',r'\\AUTRE\Partage\x.docx',r'\\DS224\CabinetCardio\..\secret.docx']:
        with pytest.raises(Refus):d.resoudre(path)


def test_sexe_et_date_future():
    for data in [{'Nom':'FICTIF','Prenom':'Test','DDN':'01/01/2030','Sexe':'F'}, {'Nom':'FICTIF','Prenom':'Test','DDN':'01/01/2000','Sexe':''}]:
        with pytest.raises(Refus):patient_valide(data,date(2026,9,12))


def test_archive_immuable_et_zip_expansion(tmp_path):
    import zipfile
    import hashlib
    d=Documents(tmp_path,r'\\DS224\CabinetCardio')
    p=tmp_path/'test.docx'
    with zipfile.ZipFile(p,'w') as z:z.writestr('word/document.xml','<document>TEST</document>')
    target,sha=d.conserver(r'\\DS224\CabinetCardio\test.docx','.docx')
    p.write_bytes(b'change')
    assert hashlib.sha256(d.resoudre(target).read_bytes()).hexdigest()==sha
    with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml','<document/>')
        for i in range(4097):z.writestr(str(i),'x')
    with pytest.raises(Refus,match='DOCX invalide'):d.conserver(r'\\DS224\CabinetCardio\test.docx','.docx')
