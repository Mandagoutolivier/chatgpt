"""Le service livre des statuts explicites, le client IA reste strict."""
from copy import deepcopy
from pathlib import Path

import pytest
from psycopg.types.json import Jsonb

from cabinet.correspondants import normaliser_indicateurs
from cabinet.domain import Refus
from cabinet.files import Documents
from cabinet.migration import appliquer, construire_plan
from cabinet.service import Service
from test_migration import workbook
from test_service import rpc, service


class Resultat(list):
    def fetchone(self): return self[0] if self else None
    def fetchall(self): return list(self)


class RessourcesMemoire:
    """Seulement le stockage CRUD ; la logique testee reste celle de Service."""
    def __init__(self, donnees=None):
        self.rows = {}
        if donnees:
            self.rows[('CORRESPONDANTS', donnees['ID'])] = {
                'donnees': deepcopy(donnees), 'revision': 1}

    def execute(self, query, args=()):
        if query.startswith('SELECT 1 FROM ressources WHERE'):
            return Resultat([{'present': 1}] if tuple(args) in self.rows else [])
        if query.startswith('SELECT donnees,revision FROM ressources WHERE genre=%s AND id=%s'):
            row = self.rows.get(tuple(args))
            return Resultat([deepcopy(row)] if row else [])
        if query.startswith('INSERT INTO ressources('):
            genre, ident, valeur = args
            self.rows[genre, ident] = {'donnees': deepcopy(valeur.obj), 'revision': 1}
            return Resultat()
        if query.startswith('UPDATE ressources SET donnees='):
            valeur, genre, ident = args
            old = self.rows[genre, ident]
            self.rows[genre, ident] = {'donnees': deepcopy(valeur.obj), 'revision': old['revision'] + 1}
            return Resultat()
        if query.startswith('SELECT donnees,revision FROM ressources WHERE'):
            genre = args[0] if args else 'CORRESPONDANTS'
            rows = [deepcopy(row) for (g, _), row in sorted(self.rows.items()) if g == genre]
            if 'LIMIT %s OFFSET %s' in query: rows = rows[args[2]:args[2] + args[1]]
            return Resultat(rows)
        raise AssertionError(query)


@pytest.fixture
def logique(tmp_path):
    return Service('inutilise', Documents(tmp_path, r'\\NAS-FICTIF\Cabinet'))


def resoudre(logique, db, ident):
    return logique._operation(db, {}, 'correspondent.resolve', {'id': ident})


@pytest.mark.parametrize('data', [{}, {'Actif': '', 'AValider': '', 'ParDefaut': ''}])
def test_normalisation_anciens_indicateurs_ne_modifie_pas_la_source(data):
    original = deepcopy(data)
    assert normaliser_indicateurs(data) == {'Actif': '1', 'AValider': '0', 'ParDefaut': '0'}
    assert data == original


@pytest.mark.parametrize('champ', ['Actif', 'AValider', 'ParDefaut'])
@pytest.mark.parametrize('value', ['inconnu', '2', None, True])
def test_indicateur_inconnu_jamais_interprete_comme_valide(champ, value):
    with pytest.raises(Refus, match=champ):
        normaliser_indicateurs({champ: value})


@pytest.mark.parametrize('operation', ['table.add', 'correspondent.save'])
@pytest.mark.parametrize('flags', [{}, {'Actif': '', 'AValider': '', 'ParDefaut': ''}])
def test_creation_normale_resolue_avec_statuts_canoniques(logique, operation, flags):
    db = RessourcesMemoire()
    data = {'ID': 'C1', 'Nom': 'FICTIF', **flags}
    created = logique._operation(db, {'roles': ['secretariat']}, operation,
                                {'genre': 'CORRESPONDANTS', 'data': data})
    assert (created['Actif'], created['AValider'], created['ParDefaut']) == ('1', '0', '0')
    assert resoudre(logique, db, created['ID']) == created
    assert db.rows['CORRESPONDANTS', 'C1']['donnees']['AValider'] == '0'


@pytest.mark.parametrize('flags', [{}, {'Actif': '', 'AValider': '', 'ParDefaut': ''}])
def test_lecture_et_resolution_ancien_stockage_sans_reecriture(logique, flags):
    old = {'ID': 'C1', 'Nom': 'FICTIF', **flags}
    db = RessourcesMemoire(old)
    record = logique._record(db, 'CORRESPONDANTS', 'C1')
    listed = logique._operation(db, {}, 'table.read', {'genre': 'CORRESPONDANTS'})['items'][0]
    assert resoudre(logique, db, 'C1') == record == listed
    assert (record['Actif'], record['AValider'], record['ParDefaut']) == ('1', '0', '0')
    assert db.rows['CORRESPONDANTS', 'C1']['donnees'] == old


@pytest.mark.parametrize('flags', [
    {'Actif': '0'}, {'AValider': '1'}, {'AValider': 'inconnu'}, {'Actif': 'inconnu'},
    {'ParDefaut': 'inconnu'}, {'AValider': None}])
def test_resolution_refuse_inactif_non_valide_ou_corrompu(logique, flags):
    db = RessourcesMemoire({'ID': 'C1', 'Nom': 'FICTIF', **flags})
    with pytest.raises(Refus, match='Destinataire absent'):
        resoudre(logique, db, 'C1')


def test_edition_partielle_ne_valide_pas_un_correspondant(logique):
    db = RessourcesMemoire()
    old = logique._save(db, 'CORRESPONDANTS', {'ID': 'C1', 'Nom': 'FICTIF', 'AValider': '1', 'ParDefaut': '1'})
    saved = logique._save(db, 'CORRESPONDANTS', {'ID': 'C1', '_revision': old['_revision'], 'Ville': 'TEST'}, True)
    assert saved['AValider'] == saved['ParDefaut'] == '1'
    with pytest.raises(Refus): resoudre(logique, db, 'C1')
    validated = logique._save(db, 'CORRESPONDANTS', dict(saved, AValider='0'), True)
    assert resoudre(logique, db, 'C1') == validated


@pytest.mark.parametrize('standalone', [False, True])
@pytest.mark.parametrize('flags', [{}, {'Actif': '', 'AValider': ''}, {'Actif': 'Oui', 'AValider': 'Non'}])
def test_import_historique_statuts_canoniques_resolubles(tmp_path, logique, standalone, flags):
    fields = ['ID', 'Nom', *flags]
    row = ['C1', 'FICTIF', *flags.values()]
    name = 'base_travail_correspondants_v1.xlsx' if standalone else 'Patients.xlsx'
    sheet = 'Generalistes' if standalone else 'CORRESPONDANTS'
    workbook(tmp_path / 'Base' / name, {sheet: [fields, row]})
    plan = construire_plan(tmp_path)
    assert not plan['erreurs']
    data = plan['ressources'][0]['data']
    assert (data['Actif'], data['AValider'], data['ParDefaut']) == ('1', '0', '0')
    assert resoudre(logique, RessourcesMemoire(data), data['ID'])['AValider'] == '0'


@pytest.mark.parametrize('standalone', [False, True])
@pytest.mark.parametrize('champ', ['Actif', 'AValider'])
def test_import_statut_inconnu_bloque_avant_ecriture(tmp_path, standalone, champ):
    name = 'base_travail_correspondants_v1.xlsx' if standalone else 'Patients.xlsx'
    sheet = 'Generalistes' if standalone else 'CORRESPONDANTS'
    workbook(tmp_path / 'Base' / name, {sheet: [['ID', 'Nom', champ], ['C1', 'FICTIF', 'inconnu']]})
    plan = construire_plan(tmp_path)
    assert any(champ in error for error in plan['erreurs'])
    with pytest.raises(Refus, match='simulation'): appliquer(None, plan)


@pytest.mark.parametrize('inverse', [False, True])
def test_fusion_historique_ne_reactive_ni_ne_valide(tmp_path, logique, inverse):
    rows = [['C1', 'FICTIF', '1', '0'], ['C2', 'FICTIF', '0', '1']]
    if inverse: rows.reverse()
    workbook(tmp_path / 'Base/Patients.xlsx', {'CORRESPONDANTS': [
        ['ID', 'Nom', 'Actif', 'AValider'], *rows]})
    plan = construire_plan(tmp_path)
    assert not plan['erreurs']
    assert len(plan['ressources']) == 1
    data = plan['ressources'][0]['data']
    assert (data['Actif'], data['AValider']) == ('0', '1')
    with pytest.raises(Refus): resoudre(logique, RessourcesMemoire(data), data['ID'])


def test_creation_et_edition_excel_ne_reinitialise_pas_un_statut_existant():
    root = Path(__file__).resolve().parents[2]
    form = (root / 'Src/Excel/ufCorrespEdit.vba').read_text(encoding='utf-8-sig')
    creation = form.split('If Len(mID) = 0 Then', 1)[1].split('Else', 1)
    assert 'd("AValider") = "0": d("ParDefaut") = "0"' in creation[0]
    assert 'd("AValider")' not in creation[1] and 'd("ParDefaut")' not in creation[1]
    destinations = (root / 'Src/Prod6/modDestinations.bas').read_text(encoding='utf-8-sig')
    assert 'If CStr(cor("Actif")) = "1" And CStr(cor("AValider")) = "0" Then' in destinations


@pytest.mark.parametrize('operation', ['table.add', 'correspondent.save'])
def test_postgresql_creation_edition_et_resolution(service, operation):
    params = {'data': {'Nom': 'FICTIF', 'Actif': '1'}}
    if operation == 'table.add': params['genre'] = 'CORRESPONDANTS'
    saved = rpc(service, operation, params)
    assert saved['AValider'] == saved['ParDefaut'] == '0'
    assert rpc(service, 'correspondent.resolve', {'id': saved['ID']}, 'medecin') == saved
    pending = rpc(service, 'correspondent.save', {'data': dict(saved, AValider='1')})
    edited = rpc(service, 'table.update', {'genre': 'CORRESPONDANTS', 'data': {
        'ID': pending['ID'], '_revision': pending['_revision'], 'Ville': 'TEST'}})
    assert edited['AValider'] == '1'
    with pytest.raises(Refus): rpc(service, 'correspondent.resolve', {'id': saved['ID']}, 'medecin')


def test_postgresql_import_et_resolution(service, tmp_path):
    workbook(tmp_path / 'Base/Patients.xlsx', {'CORRESPONDANTS': [['ID', 'Nom'], ['C1', 'FICTIF']]})
    assert appliquer(service, construire_plan(tmp_path)) == 'importe'
    resolved = rpc(service, 'correspondent.resolve', {'id': 'C1'}, 'medecin')
    assert (resolved['Actif'], resolved['AValider'], resolved['ParDefaut']) == ('1', '0', '0')
    # Une base deja peuplee par U0 beneficie aussi du contrat sans migration SQL.
    legacy = {'ID': 'C2', 'Nom': 'FICTIF ANCIEN', 'AValider': ''}
    with service.connexion() as db:
        db.execute('INSERT INTO ressources(genre,id,donnees) VALUES (%s,%s,%s)',
                   ('CORRESPONDANTS', 'C2', Jsonb(legacy)))
    assert rpc(service, 'correspondent.resolve', {'id': 'C2'}, 'medecin')['AValider'] == '0'
    assert rpc(service, 'record.get', {'genre': 'CORRESPONDANTS', 'id': 'C2'})['Actif'] == '1'
