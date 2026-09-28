"""Qualification additionnelle sur PostgreSQL isole ; aucune donnee clinique."""
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor
from threading import Barrier
import uuid
import pytest
from test_service import service, rpc, publication, parcours
from cabinet.domain import Refus


@pytest.mark.parametrize('phase', ['apres_docx', 'avant_commit'])
def test_interruption_publication_et_rejeu(service, monkeypatch, phase):
    arr, _, params = publication(service)
    rid = uuid.uuid4().hex
    if phase == 'apres_docx':
        original = service.documents._conserver_octets
        def fail_pdf(data, extension):
            if extension == '.pdf':
                raise OSError('INTERRUPTION FICTIVE APRES DOCX')
            return original(data, extension)
        monkeypatch.setattr(service.documents, '_conserver_octets', fail_pdf)
    else:
        original_connexion = service.connexion
        @contextmanager
        def interrompue():
            with original_connexion() as db:
                yield db
                raise OSError('INTERRUPTION FICTIVE AVANT COMMIT')
        monkeypatch.setattr(service, 'connexion', interrompue)
    with pytest.raises(OSError, match='INTERRUPTION FICTIVE'):
        rpc(service, 'publish', params, 'medecin', rid)
    monkeypatch.undo()
    assert rpc(service, 'publications')['items'] == []
    with service.connexion() as db:
        assert db.execute('SELECT etat FROM consultations WHERE id=%s', (arr['ID'],)).fetchone()['etat'] == 'encours'
        assert db.execute('SELECT count(*) AS n FROM commandes WHERE id=%s', (rid,)).fetchone()['n'] == 0
    result = rpc(service, 'publish', params, 'medecin', rid)
    assert rpc(service, 'publish', params, 'medecin', rid) == result
    assert len(rpc(service, 'publications')['items']) == 1
    assert len(list(service.documents.objects.iterdir())) == 2


def test_reponse_perdue_apres_commit(service):
    _, _, params = publication(service)
    rid = uuid.uuid4().hex
    committed = rpc(service, 'publish', params, 'medecin', rid)
    # La reponse est volontairement ignoree par le client ; nouvel objet serveur.
    restarted = type(service)(service.dsn, service.documents, service.clock)
    assert rpc(restarted, 'publish', params, 'medecin', rid) == committed
    assert len(rpc(restarted, 'publications')['items']) == 1
    with restarted.connexion() as db:
        assert db.execute('SELECT count(*) AS n FROM commandes WHERE id=%s', (rid,)).fetchone()['n'] == 1


def test_reprise_brouillon_apres_recreation_service(service):
    _, _, _, arr = parcours(service)
    rpc(service, 'claim', {'id': arr['ID']}, 'medecin')
    path = service.documents.root / 'reprise.docx'
    path.write_bytes(b'BROUILLON FICTIF DE QUALIFICATION')
    unc = str(service.documents.unc / path.name)
    rpc(service, 'draft', {'id': arr['ID'], 'path': unc}, 'medecin')
    restarted = type(service)(service.dsn, service.documents, service.clock)
    items = rpc(restarted, 'reprises', role='medecin')['items']
    assert len(items) == 1 and items[0]['CheminBrouillon'] == unc
    assert rpc(restarted, 'reprises', role='autre')['items'] == []
    assert path.read_bytes() == b'BROUILLON FICTIF DE QUALIFICATION'


def test_deux_prises_simultanees_meme_creneau(service):
    _, patient, _, _ = parcours(service)
    barrier = Barrier(2)
    def reserve(_):
        barrier.wait(timeout=5)
        try:
            rpc(service, 'table.add', {'genre': 'RDV', 'data': {
                'PatientID': patient['ID'], 'Date': '12/09/2026', 'Heure': '14:00',
                'DureeMin': '15', 'Statut': 'Prevu'}})
            return 'acquis'
        except Refus:
            return 'refuse'
    with ThreadPoolExecutor(max_workers=2) as pool:
        assert sorted(pool.map(reserve, [1, 2])) == ['acquis', 'refuse']
