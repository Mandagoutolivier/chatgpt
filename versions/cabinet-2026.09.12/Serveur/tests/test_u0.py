import importlib.util
import io
import os
from pathlib import Path
import subprocess
import tarfile
import zipfile
import pytest
from cabinet.files import Documents
from cabinet.domain import Refus
from cabinet.recovery import regular_tree, pack, verify_archive, rebase_json, reject_old_paths


def test_full_file_archive_roundtrip(tmp_path):
    source = tmp_path/'source'; source.mkdir()
    (source/'Patients').mkdir()
    (source/'Patients'/'brouillon.docx').write_bytes(b'BROUILLON FICTIF')
    (source/'Config').mkdir()
    (source/'Config'/'config.ini').write_text('[FICTIF]\nOption=oui\n')
    os.chmod(source/'Config'/'config.ini', 0o600)
    inventory = regular_tree(source)
    archive = tmp_path/'backup.tar.gz'
    pack(source, archive, inventory)
    target = tmp_path/'target'; target.mkdir()
    verify_archive(archive, inventory, target)
    assert regular_tree(target) == inventory


def test_altered_file_and_symlink_refused(tmp_path):
    source = tmp_path/'source'; source.mkdir()
    file = source/'lettre'; file.write_text('FICTIF')
    inventory = regular_tree(source)
    file.write_text('AUTRES')
    archive = tmp_path/'archive.tar.gz'; pack(source, archive, inventory)
    with pytest.raises(ValueError, match='altere'): verify_archive(archive, inventory)
    (source/'lien').symlink_to(file)
    with pytest.raises(ValueError, match='Lien'): regular_tree(source)


@pytest.mark.parametrize('name,kind', [('../escape', tarfile.REGTYPE), ('/escape', tarfile.REGTYPE),
                                     ('lettre', tarfile.SYMTYPE), ('lettre', tarfile.LNKTYPE)])
def test_unsafe_tar_entries_refused(tmp_path, name, kind):
    path = tmp_path/'unsafe.tar.gz'
    with tarfile.open(path, 'w:gz') as archive:
        item = tarfile.TarInfo(name); item.type = kind
        archive.addfile(item, io.BytesIO(b'') if kind == tarfile.REGTYPE else None)
    inventory = {'lettre': {'type': 'file', 'size': 0, 'sha256': 'unused'}}
    with pytest.raises(ValueError): verify_archive(path, inventory)


def test_rebase_only_known_fields_and_nested_command_result():
    old = r'\\SOURCE\Cabinet'; new = r'\\RECETTE\AutrePartage'
    value = {'PatientID': 'P0123456789ABC', 'CheminDocx': old+r'\Documents\abc.docx',
             'items': [{'CheminBrouillon': old+r'\Patients\b.docx'}], 'sha_docx': 'abc'}
    updated = rebase_json(value, old, new)
    assert updated['PatientID'] == value['PatientID']
    assert updated['sha_docx'] == 'abc'
    assert updated['items'][0]['CheminBrouillon'] == new+r'\Patients\b.docx'
    assert value['CheminDocx'].startswith(old)
    reject_old_paths(updated, old)
    with pytest.raises(ValueError, match='non prevu'):
        reject_old_paths({'inconnu': old+r'\ancien.docx'}, old)
    with pytest.raises(ValueError): rebase_json({'CheminPdf': r'\\AUTRE\Partage\f.pdf'}, old, new)


def test_archives_private_or_dedicated_group_never_world_readable(tmp_path):
    (tmp_path/'f.pdf').write_bytes(b'%PDF-1.4\nFICTIF')
    docs = Documents(tmp_path, r'\\NAS\Cabinet', os.getgid())
    unc, sha = docs.conserver(r'\\NAS\Cabinet\f.pdf', '.pdf')
    assert (docs.objects/(sha+'.pdf')).stat().st_mode & 0o777 == 0o640
    assert docs.objects.stat().st_mode & 0o777 == 0o750
    os.chmod(docs.objects, 0o770)
    with pytest.raises(Refus, match='modifiable'): docs.conserver(r'\\NAS\Cabinet\f.pdf', '.pdf')


def test_launcher_uses_git_object_despite_dirty_worktree(tmp_path, monkeypatch):
    repo = Path(__file__).resolve().parents[4]
    generator = repo/'Installateur/generer_lanceur.py'
    spec = importlib.util.spec_from_file_location('u0_launcher', generator)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
    version = tmp_path/'versions/cabinet-2026.09.12'
    (version/'Build').mkdir(parents=True); (tmp_path/'Installateur').mkdir()
    (version/'Build/outils_telechargement.ps1').write_text('# OUTIL ORIGINAL')
    (tmp_path/'Installateur/lanceur.template.ps1').write_text('@@COMMIT@@\n@@HASHES@@')
    subprocess.run(['git', 'init', '-q', str(tmp_path)], check=True)
    def git(*args):
        return subprocess.check_output(['git', '-C', str(tmp_path), *args]).decode().strip()
    git('add', '.')
    git('-c', 'user.name=Test', '-c', 'user.email=test@example.invalid', 'commit', '-qm', 'fixture')
    commit = git('rev-parse', 'HEAD')
    monkeypatch.setattr(module, 'ROOT', tmp_path); monkeypatch.setattr(module, 'VERSION', version)
    first = module.payload(commit)
    (version/'Build/outils_telechargement.ps1').write_text('# MODIFICATION LOCALE')
    (version/'intrus.txt').write_text('NE PAS INCLURE')
    assert module.payload(commit) == first
    assert 'MODIFICATION LOCALE' not in first and 'intrus.txt' not in first


def test_verification_shell_cleanup_even_on_failure(tmp_path):
    import json
    import shutil
    scripts = Path(__file__).resolve().parents[1]
    for name in ('verifier_sauvegarde.sh', 'compose.verification.yaml'):
        shutil.copy(scripts/name, tmp_path/name)
    fake = tmp_path/'docker'
    fake.write_text('#!/usr/bin/env python3\nimport json,os,sys\n'
                    'with open(os.environ["U0_LOG"],"a") as f: f.write(json.dumps(sys.argv[1:])+"\\n")\n'
                    'sys.exit(7 if "run" in sys.argv else 0)\n')
    fake.chmod(0o755)
    log = tmp_path/'calls'
    env = dict(os.environ, PATH=str(tmp_path)+os.pathsep+os.environ['PATH'], U0_LOG=str(log))
    result = subprocess.run(['sh', str(tmp_path/'verifier_sauvegarde.sh'), 'cabinet-FICTIF'], env=env)
    assert result.returncode == 7
    calls = [json.loads(line) for line in log.read_text().splitlines()]
    assert 'down' in calls[-1] and '--volumes' in calls[-1]
    projects = [c[c.index('-p')+1] for c in calls if '-p' in c]
    assert len(set(projects)) == 1 and projects[0].startswith('cabinet-verif-')
    assert not any('stop' in c for c in calls)
