"""Restaurer dans une cible imbriquee sans accepter un ancien chemin residuel."""
import pytest
from cabinet.recovery import rebase_json, reject_old_paths, rebase_database, regular_tree, check_references
from test_service import service, publication, rpc


def test_nested_destination_is_not_mistaken_for_source():
    old = r'\\SOURCE\Cabinet'
    new = old + r'\Patients\Restauration'
    value = {'CheminDocx': old + r'\Documents\a.docx',
             'items': [{'CheminBrouillon': old + r'\Patients\b.docx'}]}
    updated = rebase_json(value, old, new)
    reject_old_paths(updated, old, new)
    assert updated['CheminDocx'] == new + r'\Documents\a.docx'
    for unexpected in (old + r'\ancien.docx', new + r'Bis\a.docx', new + r'\..\ancien.docx'):
        with pytest.raises(ValueError, match='non prevu'):
            reject_old_paths({'inconnu': unexpected}, old, new)


def test_nested_sql_restore_preserves_identity_and_cached_results(service):
    _, patient, data = publication(service)
    published = rpc(service, 'publish', data, 'medecin')
    old = str(service.documents.unc)
    target = str(service.documents.unc / 'Patients' / 'Restauration')
    with service.connexion() as db:
        rebase_database(db, old, target)
        counts = check_references(db, regular_tree(service.documents.root), target, service.documents.root)
        assert counts == {'publications': 1, 'archives': 2, 'brouillons': 0}
        saved = db.execute('SELECT donnees FROM publications').fetchone()['donnees']
        assert saved['PatientID'] == patient['ID'] and saved['sha_docx'] == published['sha_docx']
        assert saved['CheminDocx'].startswith(target + '\\')
        cached = db.execute("SELECT resultat FROM commandes WHERE resultat ? 'sha_docx'").fetchone()['resultat']
        assert cached['CheminPdf'].startswith(target + '\\')
