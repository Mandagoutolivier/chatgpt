"""Edition de la fiche existante du patient de la consultation du medecin."""
import uuid
import pytest
from cabinet.domain import Refus
from test_service import service, rpc, parcours, publication


def fiche_reservee(service):
    _, patient, _, arrivee = parcours(service)
    rpc(service, 'claim', {'id': arrivee['ID']}, 'medecin')
    return patient, arrivee


def corriger(service, patient, arrivee, **champs):
    data = dict(ID=patient['ID'], _revision=patient['_revision'], **champs)
    return rpc(service, 'patient.update', {'consultation_id': arrivee['ID'], 'data': data}, 'medecin')


def test_correction_medecin_conserve_identifiant_et_champs_non_edites(service):
    patient, arrivee = fiche_reservee(service)
    saved = corriger(service, patient, arrivee, Tel='0100000000', Prenom='Correction fictive')
    assert saved['ID'] == patient['ID']
    assert saved['Tel'] == '0100000000' and saved['Prenom'] == 'Correction fictive'
    assert saved['DDN'] == patient['DDN'] and saved['MedTraitantID'] == patient['MedTraitantID']
    assert int(saved['_revision']) == int(patient['_revision']) + 1
    assert rpc(service, 'record.get', {'genre': 'PATIENTS', 'id': patient['ID']}) == saved


def test_revision_obsolete_refusee_sans_perdre_correction_secretariat(service):
    patient, arrivee = fiche_reservee(service)
    saved = rpc(service, 'table.update', {'genre': 'PATIENTS', 'data': dict(patient, Tel='0100000001')})
    with pytest.raises(Refus, match='autre poste'):
        corriger(service, patient, arrivee, Tel='0100000002')
    assert rpc(service, 'record.get', {'genre': 'PATIENTS', 'id': patient['ID']}) == saved


@pytest.mark.parametrize('role', ['autre', 'secretariat'])
def test_consultation_autre_medecin_et_role_secretariat_refuses(service, role):
    patient, arrivee = fiche_reservee(service)
    with pytest.raises(Refus) as error:
        rpc(service, 'patient.update', {'consultation_id': arrivee['ID'], 'data': patient}, role)
    assert error.value.status == 403


def test_patients_et_tables_hors_perimetre_refuses(service):
    patient, arrivee = fiche_reservee(service)
    params = {'consultation_id': arrivee['ID'], 'data': dict(patient, ID='AUTRE-FICTIF')}
    with pytest.raises(Refus, match='Patient different'):
        rpc(service, 'patient.update', params, 'medecin')
    for op in ('table.add', 'table.update'):
        with pytest.raises(Refus) as error:
            rpc(service, op, {'genre': 'PATIENTS', 'data': patient}, 'medecin')
        assert error.value.status == 403
    assert rpc(service, 'record.get', {'genre': 'PATIENTS', 'id': patient['ID']}) == patient


@pytest.mark.parametrize('modification', [dict(DDN='31/02/1980'), dict(Sexe='inconnu'), dict(Nom='')])
def test_fiche_invalide_refusee_atomiquement(service, modification):
    patient, arrivee = fiche_reservee(service)
    with pytest.raises(Refus):
        corriger(service, patient, arrivee, **modification)
    assert rpc(service, 'record.get', {'genre': 'PATIENTS', 'id': patient['ID']}) == patient


def test_revision_absente_et_consultation_non_reservee_refusees(service):
    _, patient, _, arrivee = parcours(service)
    params = {'consultation_id': arrivee['ID'], 'data': patient}
    with pytest.raises(Refus):
        rpc(service, 'patient.update', params, 'medecin')
    rpc(service, 'claim', {'id': arrivee['ID']}, 'medecin')
    params['data'] = {'ID': patient['ID'], 'Tel': '0100000000'}
    with pytest.raises(Refus, match='autre poste'):
        rpc(service, 'patient.update', params, 'medecin')


def test_reponse_perdue_rejouee_sans_double_revision(service):
    patient, arrivee = fiche_reservee(service)
    params = {'consultation_id': arrivee['ID'], 'data': dict(patient, Tel='0100000000')}
    rid = uuid.uuid4().hex
    first = rpc(service, 'patient.update', params, 'medecin', rid)
    assert rpc(service, 'patient.update', params, 'medecin', rid) == first
    assert int(first['_revision']) == int(patient['_revision']) + 1


def test_identite_corrigee_bloque_publication_du_courrier_perime(service):
    arrivee, patient, payload = publication(service)
    corriger(service, patient, arrivee, Prenom='Prenom corrige fictif')
    with pytest.raises(Refus, match='Identite modifiee'):
        rpc(service, 'publish', payload, 'medecin')
    assert not rpc(service, 'publications')['items']


def test_correction_fiche_ne_modifie_pas_les_archives_publiees(service):
    arrivee, patient, payload = publication(service)
    archived = rpc(service, 'publish', payload, 'medecin')
    corriger(service, patient, arrivee, Prenom='Prenom corrige fictif')
    assert rpc(service, 'publications')['items'] == [archived]
