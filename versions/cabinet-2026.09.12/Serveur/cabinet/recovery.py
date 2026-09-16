"""Sauvegarde coherente et restauration complete dans une cible vide.

L'API doit etre arretee et les ecritures SMB suspendues par l'operateur.
Les empreintes avant/apres detectent une modification observee ; elles ne
remplacent pas la suspension des ecritures SMB. Aucun secret n'est affiche.
"""
from __future__ import annotations
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path, PurePosixPath, PureWindowsPath
import re
import subprocess
import tarfile
import uuid
import zipfile
import psycopg
from psycopg import sql
from psycopg.rows import dict_row
from psycopg.types.json import Jsonb
from . import SERVICE_REVISION

FORMAT = 1
SUPPORTED_SCHEMAS = {1, 2}
FILES = ('base.dump', 'fichiers.tar.gz', 'configuration.tar.gz', 'manifest.json')
PATH_KEYS = {'CheminDocx', 'CheminPdf', 'CheminBrouillon'}
JSON_COLUMNS = (('ressources', 'donnees', ('genre', 'id')),
                ('consultations', 'donnees', ('id',)),
                ('publications', 'donnees', ('id',)),
                ('commandes', 'resultat', ('compte', 'id')))


def digest(path):
    with Path(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def require(condition, message):
    if not condition:
        raise ValueError(message)


def regular_tree(root):
    """Inventaire strict : aucun lien, fichier special ou nom ambigu."""
    result = {}
    for path in sorted(root.rglob('*')):
        relative = path.relative_to(root).as_posix()
        require(not path.is_symlink() and (path.is_dir() or path.is_file()),
                'Lien ou fichier special interdit dans la sauvegarde.')
        require('\\' not in relative and '\n' not in relative and '\r' not in relative,
                'Nom de fichier ambigu interdit.')
        stat = path.stat()
        row = {'type': 'dir' if path.is_dir() else 'file', 'mode': stat.st_mode & 0o777,
               'uid': stat.st_uid, 'gid': stat.st_gid}
        if path.is_file():
            row.update(size=stat.st_size, sha256=digest(path))
        result[relative] = row
    return result


def pack(root, target, inventory):
    with tarfile.open(target, 'w:gz', dereference=False) as archive:
        for name in inventory:
            archive.add(root / name, arcname=name, recursive=False)


def safe_members(archive, inventory):
    seen = set()
    for item in archive:
        name = item.name
        path = PurePosixPath(name)
        require(name not in seen and name in inventory and not path.is_absolute()
                and '..' not in path.parts and '\\' not in name,
                'Archive contenant un chemin inattendu ou duplique.')
        seen.add(name)
        spec = inventory[name]
        require((item.isdir() and spec['type'] == 'dir') or
                (item.isfile() and spec['type'] == 'file' and item.size == spec['size']),
                'Archive contenant un lien ou un type/taille inattendu.')
        yield item, spec
    require(seen == set(inventory), 'Archive incomplete.')


def verify_archive(path, inventory, destination=None):
    """Verifier les octets, extraire sans tar.extract ni liens, restaurer les droits."""
    with tarfile.open(path, 'r:gz') as archive:
        directories = []
        for item, spec in safe_members(archive, inventory):
            target = destination / item.name if destination else None
            if item.isdir():
                if target:
                    target.mkdir(parents=True, exist_ok=True, mode=0o700)
                    directories.append((target, spec))
                continue
            hasher = hashlib.sha256()
            if target:
                target.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
            output = target.open('xb') if target else None
            try:
                with archive.extractfile(item) as source:
                    for block in iter(lambda: source.read(1024 * 1024), b''):
                        hasher.update(block)
                        if output:
                            output.write(block)
                require(hasher.hexdigest() == spec['sha256'], 'Contenu archive altere.')
                if output:
                    output.flush()
                    os.fsync(output.fileno())
            finally:
                if output:
                    output.close()
            if target:
                restore_permissions(target, spec)
        for target, spec in reversed(directories):
            restore_permissions(target, spec)


def restore_permissions(path, spec):
    if os.geteuid() == 0:
        os.chown(path, spec['uid'], spec['gid'])
    os.chmod(path, spec['mode'])


def relative_unc(value, unc):
    require(isinstance(value, str) and value.startswith('\\\\'), 'Chemin UNC absent.')
    base = PureWindowsPath(unc)
    require(base.drive.startswith('\\\\') and '..' not in base.parts, 'Racine UNC invalide.')
    try:
        relative = PureWindowsPath(value).relative_to(base)
    except ValueError:
        raise ValueError('Document reference hors du partage attendu.') from None
    require(relative.parts and '..' not in relative.parts and
            not any(':' in part for part in relative.parts), 'Chemin document interdit.')
    return PurePosixPath(*relative.parts).as_posix()


def check_references(db, inventory, unc, root=None):
    """Toutes publications et tous brouillons references, pas un echantillon."""
    counts = {'publications': 0, 'archives': 0, 'brouillons': 0}
    for row in db.execute('SELECT donnees FROM publications'):
        counts['publications'] += 1
        for key, hash_key in (('CheminDocx', 'sha_docx'), ('CheminPdf', 'sha_pdf')):
            name = relative_unc(row['donnees'][key], unc)
            spec = inventory.get(name, {})
            require(spec.get('type') == 'file' and spec.get('sha256') == row['donnees'].get(hash_key),
                    'Archive referencee absente ou empreinte differente.')
            counts['archives'] += 1
            if root:
                path = root / name
                require(path.is_file() and digest(path) == spec['sha256'], 'Document restaure altere.')
                if key == 'CheminDocx':
                    with zipfile.ZipFile(path) as document:
                        require('word/document.xml' in document.namelist() and document.testzip() is None,
                                'DOCX restaure invalide.')
                else:
                    with path.open('rb') as document:
                        require(document.read(5) == b'%PDF-', 'PDF restaure invalide.')
    for row in db.execute('SELECT donnees FROM consultations'):
        value = row['donnees'].get('CheminBrouillon')
        if value:
            name = relative_unc(value, unc)
            require(inventory.get(name, {}).get('type') == 'file', 'Brouillon reference absent.')
            counts['brouillons'] += 1
    return counts


def rebase_json(value, old_unc, new_unc, key=''):
    if isinstance(value, dict):
        return {k: rebase_json(v, old_unc, new_unc, k) for k, v in value.items()}
    if isinstance(value, list):
        return [rebase_json(v, old_unc, new_unc, key) for v in value]
    if key in PATH_KEYS and value:
        relative = relative_unc(value, old_unc)
        return str(PureWindowsPath(new_unc).joinpath(*PurePosixPath(relative).parts))
    return value


def reject_old_paths(value, old_unc):
    if isinstance(value, dict):
        for child in value.values():
            reject_old_paths(child, old_unc)
    elif isinstance(value, list):
        for child in value:
            reject_old_paths(child, old_unc)
    elif isinstance(value, str):
        require(not value.lower().startswith(old_unc.rstrip('\\').lower()+'\\'),
                'Ancien chemin restant dans un champ non prevu : intervention explicite requise.')


def rebase_database(db, old_unc, new_unc):
    if PureWindowsPath(old_unc) == PureWindowsPath(new_unc):
        return
    for table, column, keys in JSON_COLUMNS:
        columns = sql.SQL(',').join(sql.Identifier(x) for x in (*keys, column))
        rows = db.execute(sql.SQL('SELECT {} FROM {}').format(columns, sql.Identifier(table))).fetchall()
        for row in rows:
            updated = rebase_json(row[column], old_unc, new_unc)
            reject_old_paths(updated, old_unc)
            if updated != row[column]:
                condition = sql.SQL(' AND ').join(sql.SQL('{}=%s').format(sql.Identifier(k)) for k in keys)
                db.execute(sql.SQL('UPDATE {} SET {}=%s WHERE {}').format(
                    sql.Identifier(table), sql.Identifier(column), condition),
                    [Jsonb(updated), *(row[k] for k in keys)])


def run(command):
    # La sortie PostgreSQL peut contenir des valeurs : ne jamais la recopier dans un log.
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    require(result.returncode == 0, Path(command[0]).name+' a echoue ; cible incomplete conservee, aucune bascule.')


def connection():
    return psycopg.connect('', row_factory=dict_row, autocommit=True, connect_timeout=10)


def backup(root, backups, config, unc):
    require(not backups.resolve().is_relative_to(root.resolve()), 'Sauvegardes dans les donnees interdites.')
    require(root.is_dir() and config.is_dir(), 'Volume de donnees ou configuration absent.')
    stamp = 'cabinet-'+datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')+'-'+uuid.uuid4().hex[:12]
    target = backups / stamp
    target.mkdir(mode=0o700)
    with connection() as db:
        # Meme verrou que les mutations applicatives. Maintenu pendant toute la capture.
        db.execute('SELECT pg_advisory_lock(20260912)')
        inventory = regular_tree(root)
        configuration = regular_tree(config)
        schema = db.execute('SELECT max(version) AS v FROM schema_version').fetchone()['v']
        require(type(schema) is int and schema in SUPPORTED_SCHEMAS, 'Schema de base non pris en charge.')
        references = check_references(db, inventory, unc)
        run(['pg_dump', '--format=custom', '--no-owner', '--file='+str(target/'base.dump')])
        pack(root, target/'fichiers.tar.gz', inventory)
        pack(config, target/'configuration.tar.gz', configuration)
        verify_archive(target/'fichiers.tar.gz', inventory)
        verify_archive(target/'configuration.tar.gz', configuration)
        require(regular_tree(root) == inventory and regular_tree(config) == configuration,
                'Ecritures detectees pendant la sauvegarde : aucun marqueur TERMINE.')
        manifest = {'format': FORMAT, 'source_unc': unc, 'schema': schema,
                    'revision': SERVICE_REVISION, 'fichiers': inventory,
                    'configuration': configuration, 'references': references,
                    'acl_synology': 'A sauvegarder et restaurer via DSM ; modes POSIX conserves.',
                    'ecritures_smb': 'Suspension declaree par operateur, controle de stabilite effectue.'}
        (target/'manifest.json').write_text(json.dumps(manifest, ensure_ascii=True, indent=2))
        sums = {name: digest(target/name) for name in FILES}
        (target/'SHA256SUMS.json').write_text(json.dumps(sums, sort_keys=True))
        run(['pg_restore', '--list', str(target/'base.dump')])
        # Une coupure avant ce marqueur laisse un jeu explicitement incomplet.
        for path in target.iterdir():
            os.chmod(path, 0o600)
            with path.open('rb') as stream:
                os.fsync(stream.fileno())
        with (target/'TERMINE').open('x') as stream:
            stream.write(digest(target/'SHA256SUMS.json')+'\n')
            stream.flush()
            os.fsync(stream.fileno())
        fd = os.open(target, os.O_RDONLY | os.O_DIRECTORY)
        try:
            os.fsync(fd)
        finally:
            os.close(fd)
    return {'sauvegarde': stamp, **references, 'terminee': True}


def inspect_bundle(bundle):
    require(bundle.is_dir() and not bundle.is_symlink(), 'Jeu de sauvegarde absent.')
    require((bundle/'TERMINE').is_file(), 'Sauvegarde incomplete : TERMINE absent.')
    sums_path = bundle/'SHA256SUMS.json'
    require((bundle/'TERMINE').read_text().strip() == digest(sums_path), 'Manifeste altere.')
    sums = json.loads(sums_path.read_text())
    require(set(sums) == set(FILES), 'Liste de composants de sauvegarde inattendue.')
    for name, expected in sums.items():
        require(not (bundle/name).is_symlink() and digest(bundle/name) == expected, 'Composant de sauvegarde altere.')
    manifest = json.loads((bundle/'manifest.json').read_text())
    require(manifest['format'] == FORMAT and type(manifest['schema']) is int
            and manifest['schema'] in SUPPORTED_SCHEMAS, 'Format de sauvegarde incompatible.')
    verify_archive(bundle/'fichiers.tar.gz', manifest['fichiers'])
    verify_archive(bundle/'configuration.tar.gz', manifest['configuration'])
    run(['pg_restore', '--list', str(bundle/'base.dump')])
    return manifest


def restore(bundle, root, config_target, unc, role):
    # Aucun stop d'API, aucune suppression ni creation de base. La cible doit deja etre vide.
    require(root.is_dir() and not root.is_symlink() and not any(root.iterdir()), 'Volume cible non vide.')
    require(config_target.is_dir() and not config_target.is_symlink() and not any(config_target.iterdir()),
            'Dossier de configuration recuperee non vide.')
    require(not bundle.resolve().is_relative_to(root.resolve()), 'Sauvegarde dans la cible interdite.')
    relative_unc(unc.rstrip('\\')+'\\controle', unc)
    manifest = inspect_bundle(bundle)
    with connection() as db:
        count = db.execute("SELECT count(*) AS n FROM pg_class c JOIN pg_namespace n ON n.oid=c.relnamespace "
                           "WHERE n.nspname NOT IN ('pg_catalog','information_schema') "
                           "AND n.nspname NOT LIKE 'pg_toast%' AND c.relkind IN ('r','p','v','m','S','f')").fetchone()['n']
        require(count == 0, 'Base cible non vide : restauration refusee avant toute ecriture.')
    command = ['pg_restore', '--exit-on-error', '--single-transaction', '--no-owner', '--no-acl',
               '--dbname='+os.environ.get('PGDATABASE', 'cabinet')]
    if role:
        require(bool(re.fullmatch(r'[a-z_][a-z0-9_]{0,62}', role)), 'Role cible invalide.')
        command.append('--role='+role)
    run(command+[str(bundle/'base.dump')])
    with connection() as db:
        restored_schema = db.execute('SELECT max(version) AS v FROM schema_version').fetchone()['v']
        require(restored_schema == manifest['schema'], 'Schema restaure different du manifeste.')
    os.chmod(config_target, 0o700)
    verify_archive(bundle/'fichiers.tar.gz', manifest['fichiers'], root)
    # La configuration d'origine est recuperee a part, jamais appliquee au nouveau projet.
    verify_archive(bundle/'configuration.tar.gz', manifest['configuration'], config_target)
    with connection() as db, db.transaction():
        rebase_database(db, manifest['source_unc'], unc)
        counts = check_references(db, manifest['fichiers'], unc, root)
        require(counts == manifest['references'], 'Comptes de references divergents apres restauration.')
    return {'restauration_technique': 'reussie', **counts, 'chemins_adaptes': unc != manifest['source_unc'],
            'recette_office_smb_requise': True, 'acl_synology_restaures': False}


def main():
    os.umask(0o077)
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['backup', 'restore', 'inspect'])
    parser.add_argument('--bundle', type=Path)
    parser.add_argument('--root', type=Path, default=Path('/data'))
    parser.add_argument('--backups', type=Path, default=Path('/backups'))
    parser.add_argument('--config', type=Path, default=Path('/configuration'))
    parser.add_argument('--unc', default=os.getenv('CABINET_UNC', ''))
    parser.add_argument('--role', default='cabinet')
    parser.add_argument('--ecritures-smb-suspendues', action='store_true')
    args = parser.parse_args()
    password_file = os.getenv('PGPASSWORD_FILE')
    if password_file:
        os.environ['PGPASSWORD'] = Path(password_file).read_text().strip()
    try:
        if args.action == 'backup':
            require(args.ecritures_smb_suspendues, 'Suspendre les ecritures SMB et fournir --ecritures-smb-suspendues.')
            result = backup(args.root, args.backups, args.config, args.unc)
        else:
            require(args.bundle is not None, '--bundle obligatoire.')
            if args.action == 'inspect':
                result = {'integrite': 'verifiee', **inspect_bundle(args.bundle)['references']}
            else:
                result = restore(args.bundle, args.root, args.config, args.unc, args.role)
        print(json.dumps(result, ensure_ascii=True, sort_keys=True))
    except Exception as exc:
        # Les exceptions PostgreSQL peuvent contenir des donnees : message generique.
        message = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
        parser.exit(1, 'Controle interrompu : '+message+'\n')


if __name__ == '__main__':
    main()
