"""Mise a jour ponctuelle de la seule API U2, source qualifiee 06506de.

Lancer sur DS224 avec sudo python3 et --appliquer. Sans cet argument,
effectue les controles seulement. Aucun jeton applicatif n'est lu.
"""
from pathlib import Path
import argparse
import hashlib
import json
import os
import re
import shutil
import socket
import subprocess
import time
import urllib.request
import zipfile

COMMIT = '06506de063c76cc6811da46b5c503e5b549fc571'
ZIP_SHA = '2a3c05875a0cbc6dee71e68019fed2561366e0defe4d41b9f8356834aff17707'
PROJECT = 'cabinetcardio-test-u2'
API = PROJECT + '-api-1'
DB = PROJECT + '-db-1'
BASE = Path('/volume1/docker') / PROJECT
SHARE = Path('/volume1/CabinetCardioTestU2/Patients/_Qualification20260929')
PACKAGE = SHARE / 'CorrectifsFichePatient-06506de'
CHANGED = ('contract.py', 'service.py', 'recovery.py')
OLD = {
    'contract.py': '6b07e2277cd85819ea0a0ee6e943cc97cc5cb9191acfc99c73c218917a9151d5',
    'service.py': '38ac38c5e13f0c1e876caaf12596758155c9dd8b3e5de120f8fffc32977ce726',
}
RECOVERY_7E3 = 'd9b65b5bd4015cf6ce4b06b3968260249598d4a9fd695291d9959f4d81c185c6'


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def normalized(data):
    return sha(data.decode('utf-8-sig').replace('\r\n', '\n').encode())


def run(args, **kw):
    result = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, **kw)
    # Ne pas recopier une erreur Docker pouvant contenir la configuration.
    require(result.returncode == 0, 'Commande echouee : ' + ' '.join(args[:3]))
    return result.stdout


def inspect(name):
    return json.loads(run(['docker', 'inspect', name]))[0]


def inside(path, parent):
    return str(path.resolve()).startswith(str(parent.resolve()) + '/')


def health():
    with urllib.request.urlopen('http://127.0.0.1:8766/health', timeout=3) as response:
        result = json.load(response)
    require(result.get('status') == 'ok' and result.get('protocole') == 2,
            'Sante du service U2 invalide')
    return result


def wait_health():
    for _ in range(30):
        try:
            return health()
        except Exception:
            time.sleep(1)
    raise RuntimeError('Service U2 indisponible apres redemarrage')


def runtime_hashes(command, paths):
    code = ('from pathlib import Path; import hashlib,json; '
            'paths=' + repr(paths) + '; '
            'print(json.dumps({p:hashlib.sha256(Path(p).read_bytes()'
            '.decode("utf-8-sig").replace("\\r\\n","\\n").encode()).hexdigest() '
            'for p in paths}))')
    return json.loads(run(command + ['python', '-c', code]))


def validate_api(api):
    labels = api['Config']['Labels']
    require(labels.get('com.docker.compose.project') == PROJECT and
            labels.get('com.docker.compose.service') == 'api', 'Projet API incorrect')
    require(api['State']['Running'], 'API arretee')
    require(api['HostConfig']['PortBindings'] == {
        '8765/tcp': [{'HostIp': '127.0.0.1', 'HostPort': '8766'}]}, 'Port API incorrect')
    data = [m for m in api['Mounts'] if m['Destination'] == '/data']
    require(len(data) == 1 and data[0]['Source'] == '/volume1/CabinetCardioTestU2',
            'Donnees API hors recette')
    require(Path(data[0]['Source']).resolve() == Path(data[0]['Source']),
            'Lien symbolique dans les donnees')
    require(not any(m['Destination'].startswith('/opt/cabinet') for m in api['Mounts']),
            'Code monte dans le conteneur : verification manuelle requise')
    for m in api['Mounts']:
        if m['Destination'].startswith('/run/secrets/'):
            require(inside(Path(m['Source']), BASE), 'Secret hors recette')


def atomically_copy(source, destination):
    temp = destination.with_name(destination.name + '.u2-update-tmp')
    require(not temp.exists(), 'Fichier temporaire deja present')
    try:
        shutil.copy2(source, temp)
        os.replace(str(temp), str(destination))
    finally:
        if temp.exists():
            temp.unlink()


def main(apply):
    require(socket.gethostname() == 'DS224' and os.geteuid() == 0,
            'Executer sur DS224 avec sudo python3')
    require(BASE.resolve() == BASE and PACKAGE.resolve() == PACKAGE, 'Chemin non canonique')
    os.umask(0o077)
    api = inspect(API)
    validate_api(api)
    db = inspect(DB)
    require(db['Config']['Labels'].get('com.docker.compose.project') == PROJECT and
            db['Config']['Labels'].get('com.docker.compose.service') == 'db' and
            db['State']['Health']['Status'] == 'healthy', 'Base U2 non valide')
    require(any(m['Source'] == str(BASE / 'postgres') and
                m['Destination'] == '/var/lib/postgresql/data' for m in db['Mounts']),
            'Volume PostgreSQL hors recette')
    labels = api['Config']['Labels']
    server = Path(labels['com.docker.compose.project.working_dir'])
    compose_file = Path(labels['com.docker.compose.project.config_files'])
    env_file = server / '.env'
    require(inside(server, BASE / 'code') and server.resolve() == server and
            compose_file == server / 'compose.yaml', 'Dossier Compose inattendu')
    for path in (compose_file, env_file):
        require(path.is_file() and path.resolve() == path, 'Configuration non canonique')
    # Ecarter les substitutions heritees de la session de l'operateur.
    compose_env = {k: v for k, v in os.environ.items()
                   if not k.startswith(('CABINET_', 'COMPOSE_', 'DOCKER_'))}
    compose = ['docker', 'compose', '--project-name', PROJECT,
               '--file', str(compose_file), '--env-file', str(env_file)]
    config = json.loads(run(compose + ['config', '--format', 'json'], env=compose_env))
    cfg = config['services']['api']
    require(config.get('name') == PROJECT, 'Nom Compose inattendu')
    require(cfg.get('container_name', API) == API, 'Nom conteneur inattendu')
    require(len(cfg['ports']) == 1 and cfg['ports'][0].get('host_ip') == '127.0.0.1'
            and str(cfg['ports'][0]['published']) == '8766'
            and int(cfg['ports'][0]['target']) == 8765, 'Port Compose hors recette')
    require(str(cfg.get('user')) == api['Config']['User'], 'Utilisateur Compose divergent')
    require(cfg.get('read_only') is True and api['HostConfig']['ReadonlyRootfs'],
            'Protection lecture seule absente')
    current_binds = {(m['Source'], m['Destination']) for m in api['Mounts'] if m['Type'] == 'bind'}
    for mount in cfg.get('volumes', []):
        require(mount.get('type') == 'bind' and
                (mount['source'], mount['target']) in current_binds,
                'Montage Compose divergent')
    require(any(m['target'] == '/data' and m['source'] == '/volume1/CabinetCardioTestU2'
                for m in cfg['volumes']), 'Donnees Compose hors recette')
    current_env = dict(e.split('=', 1) for e in api['Config']['Env'] if '=' in e)
    require(all(str(value) == current_env.get(key) for key, value in cfg.get('environment', {}).items()),
            'Environnement Compose divergent')
    for item in cfg.get('secrets', []):
        source = config['secrets'][item['source']]['file']
        require((source, '/run/secrets/' + item.get('target', item['source'])) in current_binds,
                'Secret Compose divergent')
    image_ref = api['Config']['Image']
    require(image_ref in (PROJECT + '-api', PROJECT + '-api:latest'), 'Image API inattendue')
    require(cfg.get('image', image_ref) in (PROJECT + '-api', PROJECT + '-api:latest'),
            'Image Compose divergente')
    health()

    archive = PACKAGE / 'sources-06506de.zip'
    require(sha(archive.read_bytes()) == ZIP_SHA, 'Archive de sources alteree')
    prefix = 'chatgpt-' + COMMIT + '/versions/cabinet-2026.09.12/'
    with zipfile.ZipFile(archive) as z:
        names = [n for n in z.namelist() if n.startswith(prefix + 'Serveur/cabinet/')
                 and not n.endswith('/')]
        require(all('/' not in n[len(prefix + 'Serveur/cabinet/'):] for n in names),
                'Arborescence source inattendue')
        names += [prefix + 'Build/schemas.json', prefix + 'Serveur/requirements-runtime.txt']
        payload = {n[len(prefix):]: z.read(n) for n in names}
    expected_new = {n: normalized(data) for n, data in payload.items()}
    expected_old = dict(expected_new)
    expected_old.update({'Serveur/cabinet/' + n: h for n, h in OLD.items()})
    # Le controle NAS du 29/09 portait sur 7e3a6a0. Le correctif recovery
    # d'a152465/b40e7d9 a ete qualifie ensuite dans la copie de restauration,
    # sans que cela prouve son installation dans l'API courante.
    expected_7e3 = dict(expected_old)
    expected_7e3['Serveur/cabinet/recovery.py'] = RECOVERY_7E3
    paths = {n: '/opt/cabinet/' + ('Serveur/requirements.txt' if
             n == 'Serveur/requirements-runtime.txt' else n) for n in payload}
    actual = runtime_hashes(['docker', 'exec', API], list(paths.values()))
    deployed = {n: actual[p] for n, p in paths.items()}
    known = {'7e3a6a0': expected_7e3, 'b40e7d9': expected_old, '06506de': expected_new}
    source_before = next((name for name, hashes in known.items() if deployed == hashes), None)
    if source_before is None:
        diagnostic = {'statut': 'ARRET_AVANT_MODIFICATION', 'image': api['Image'],
                      'dossier': str(server), 'fichiers': []}
        for n in sorted(payload):
            local = server.parent / n
            disk_hash = normalized(local.read_bytes()) if local.is_file() and local.resolve() == local else None
            entry = {'fichier': n, 'actif': deployed[n], 'disque': disk_hash,
                     'attendus': {name: hashes[n] for name, hashes in known.items()}}
            diagnostic['fichiers'].append(entry)
            if deployed[n] not in set(entry['attendus'].values()):
                print('ECART_CODE : ' + n, flush=True)
        dest = PACKAGE / 'diagnostic-code.json'
        dest.write_text(json.dumps(diagnostic, indent=2))
        dest.chmod(0o644)
        raise RuntimeError('Code deploye different des versions qualifiees ; diagnostic-code.json cree')
    for n in payload:
        path = server.parent / n
        require(path.is_file() and path.resolve() == path, 'Source locale manquante ou non canonique : ' + n)
        require(normalized(path.read_bytes()) == deployed[n], 'Source disque et image divergentes : ' + n)
    print('CONTROLES_U2_OK ; source=' + source_before + ' ; dossier=' + str(server), flush=True)
    if deployed == expected_new:
        print('DEJA_A_JOUR_06506de ; verifier maintenant les connexions des postes', flush=True)
        return
    if not apply:
        print('VERIFICATION_SEULE ; aucune modification du service', flush=True)
        return

    stamp = time.strftime('%Y%m%dT%H%M%SZ', time.gmtime())
    backup = BASE / 'sauvegardes' / ('avant-fiche-patient-' + stamp)
    require(backup.parent.resolve() == backup.parent, 'Sauvegardes hors chemin attendu')
    backup.mkdir(mode=0o700)
    private_log = backup / 'construction.log'
    report = {'source': COMMIT, 'source_avant': source_before, 'statut': 'EN_COURS', 'sauvegarde': str(backup),
              'api': API, 'source_installee': False, 'validation_postes': False}
    changed_disk = False
    tagged_new = False
    restart_attempted = False
    old_tag = PROJECT + '-api:avant-fiche-' + stamp.lower()
    new_tag = PROJECT + '-api:fiche-06506de-' + stamp.lower()
    protected = {p: sha(p.read_bytes()) for p in (compose_file, env_file)}
    other_ids = run(['docker', 'ps', '-q']).decode().split()
    others = {i: inspect(i) for i in other_ids if i != api['Id'] and not api['Id'].startswith(i)}
    try:
        for n in CHANGED:
            shutil.copy2(server / 'cabinet' / n, backup / n)
        run(['docker', 'image', 'tag', api['Image'], old_tag])
        report['image_avant'] = api['Image']
        report['image_retour'] = old_tag
        print('SAUVEGARDE_BASE_U2', flush=True)
        with (backup / 'postgres.dump').open('wb') as output:
            result = subprocess.run(['docker', 'exec', DB, 'pg_dump', '-U', 'postgres',
                                     '-d', 'cabinet', '-Fc'], stdout=output, stderr=subprocess.PIPE)
        require(result.returncode == 0 and (backup / 'postgres.dump').stat().st_size > 100,
                'Sauvegarde de la base echouee')
        # Aucun fichier de donnees modifie ; ce dump accompagne le retour de code,
        # il ne constitue pas a lui seul une sauvegarde complete base + documents.
        build = backup / 'construction'
        build.mkdir()
        for n in CHANGED:
            dest = build / n
            dest.write_bytes(payload['Serveur/cabinet/' + n])
            dest.chmod(0o644)
        image_user = json.loads(run(['docker', 'image', 'inspect', api['Image']]))[0]['Config']['User']
        require(re.fullmatch(r'[A-Za-z0-9_-]+(?::[A-Za-z0-9_-]+)?', image_user) is not None,
                'Utilisateur image inattendu')
        (build / 'Dockerfile').write_text(
            'FROM ' + old_tag + '\nUSER root\n'
            'COPY --chown=0:0 contract.py service.py recovery.py /opt/cabinet/Serveur/cabinet/\n'
            'USER ' + image_user + '\nLABEL cabinet.source="' + COMMIT + '"\n')
        print('CONSTRUCTION_API_U2 ; le service actuel reste disponible', flush=True)
        with private_log.open('wb') as output:
            result = subprocess.run(['docker', 'build', '--pull=false', '--network=none',
                                     '-t', new_tag, str(build)], stdout=output, stderr=subprocess.STDOUT)
        require(result.returncode == 0, 'Construction echouee ; journal prive dans la sauvegarde')
        candidate = runtime_hashes(['docker', 'run', '--rm', '--network', 'none', '--read-only',
                                    '--entrypoint', '', new_tag], list(paths.values()))
        require({n: candidate[p] for n, p in paths.items()} == expected_new,
                'Image construite non conforme aux sources qualifiees')
        require(all(sha(p.read_bytes()) == h for p, h in protected.items()), 'Configuration modifiee pendant la preparation')
        require(inspect(API)['Id'] == api['Id'] and inspect(DB)['Id'] == db['Id'],
                'Conteneurs modifies pendant la preparation')
        changed_disk = True
        for n in CHANGED:
            atomically_copy(build / n, server / 'cabinet' / n)
        tagged_new = True
        run(['docker', 'image', 'tag', new_tag, image_ref])
        print('REDEMARRAGE_DE_LA_SEULE_API_U2', flush=True)
        restart_attempted = True
        run(compose + ['up', '-d', '--no-deps', '--no-build', '--force-recreate', 'api'], env=compose_env)
        report['sante'] = wait_health()
        current = inspect(API)
        validate_api(current)
        new_id = json.loads(run(['docker', 'image', 'inspect', new_tag]))[0]['Id']
        require(current['Image'] == new_id, 'Nouvelle image non utilisee')
        actual = runtime_hashes(['docker', 'exec', API], list(paths.values()))
        require({n: actual[p] for n, p in paths.items()} == expected_new, 'Code actif non conforme')
        require(all(sha(p.read_bytes()) == h for p, h in protected.items()), 'Configuration alteree')
        for ident, before in others.items():
            after = inspect(ident)
            require(after['Id'] == before['Id'] and after['Image'] == before['Image'] and
                    after['State']['StartedAt'] == before['State']['StartedAt'],
                    'Un autre conteneur a change pendant la mise a jour')
        report.update(statut='SUCCESS', source_installee=True, image_apres=new_id,
                      autres_conteneurs_inchanges=True, configuration_inchangee=True)
        print('MISE_A_JOUR_U2_OK ; validation medecin et secretariat encore necessaire', flush=True)
    except BaseException as exc:
        report['statut'] = 'ECHEC'
        report['erreur'] = str(exc)[:300]
        if changed_disk or tagged_new or restart_attempted:
            try:
                print('RETOUR_AUTOMATIQUE_A_L_ANCIENNE_API_U2', flush=True)
                for n in CHANGED:
                    atomically_copy(backup / n, server / 'cabinet' / n)
                run(['docker', 'image', 'tag', api['Image'], image_ref])
                if restart_attempted:
                    run(compose + ['up', '-d', '--no-deps', '--no-build', '--force-recreate', 'api'], env=compose_env)
                wait_health()
                require(inspect(API)['Image'] == api['Image'], 'Image precedente non restauree')
                report['retour_arriere'] = 'OK'
            except BaseException as rollback_error:
                report['retour_arriere'] = 'A_VERIFIER'
                report['erreur_retour'] = str(rollback_error)[:300]
        raise
    finally:
        report['date_utc'] = time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime())
        content = json.dumps(report, indent=2, ensure_ascii=False)
        (backup / 'resultat.json').write_text(content)
        public = PACKAGE / ('mise-a-jour-' + stamp + '.json')
        public.write_text(content)
        public.chmod(0o644)
        print('RAPPORT=' + str(public), flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--appliquer', action='store_true')
    args = parser.parse_args()
    try:
        main(args.appliquer)
    except Exception as error:
        print('ARRET : ' + str(error), flush=True)
        raise SystemExit(1)
