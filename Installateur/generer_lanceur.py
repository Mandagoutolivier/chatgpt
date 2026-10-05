"""Assembler le CMD autonome. Le commit doit contenir exactement les fichiers de version locaux."""
import argparse
import hashlib
import json
import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VERSION = ROOT / 'versions/cabinet-2026.09.12'
HEADER = r'''@echo off
setlocal
set "CABINET_SELF=%~f0"
set "CABINET_BOOT_DIR=%TEMP%\CabinetBoot-%RANDOM%-%RANDOM%"
if exist "%CABINET_BOOT_DIR%" goto collision
mkdir "%CABINET_BOOT_DIR%"
if errorlevel 1 exit /b 1
set "CABINET_BOOT_PS=%CABINET_BOOT_DIR%\demarrer.ps1"
powershell.exe -NoLogo -NoProfile -Command "$ErrorActionPreference='Stop';$s=[IO.File]::ReadAllText($env:CABINET_SELF);$m=[regex]::Match($s,'(?m)^# CABINET_POWERSHELL_PAYLOAD_V1\r?$');if(-not $m.Success){throw 'Lanceur incomplet'};[IO.File]::WriteAllText($env:CABINET_BOOT_PS,$s.Substring($m.Index+$m.Length),(New-Object Text.UTF8Encoding($true)))"
if errorlevel 1 goto erreur
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "%CABINET_BOOT_PS%"
set "CABINET_RESULT=%ERRORLEVEL%"
goto nettoyage
:erreur
set "CABINET_RESULT=1"
:nettoyage
if exist "%CABINET_BOOT_PS%" del "%CABINET_BOOT_PS%"
rmdir "%CABINET_BOOT_DIR%"
echo.
echo Vous pouvez fermer cette fenetre. En cas de pause, relancez ce meme fichier.
pause
exit /b %CABINET_RESULT%
:collision
echo Relancez le fichier : le dossier temporaire existe deja.
pause
exit /b 1
# CABINET_POWERSHELL_PAYLOAD_V1
'''

def payload(commit):
    if not re.fullmatch('[a-f0-9]{40}', commit):
        raise ValueError('SHA de commit invalide')
    hashes = {}
    for path in sorted(VERSION.rglob('*')):
        if path.is_file() and '__pycache__' not in path.parts and '.pytest_cache' not in path.parts:
            hashes[path.relative_to(VERSION).as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    helpers = (VERSION / 'Build/outils_telechargement.ps1').read_text(encoding='utf-8-sig')
    main = (ROOT / 'Installateur/lanceur.template.ps1').read_text(encoding='utf-8-sig')
    main = main.replace('@@COMMIT@@', commit).replace('@@HASHES@@', json.dumps(hashes, ensure_ascii=True, indent=2))
    return helpers + '\n' + main

def verifier_commit(commit):
    """Les empreintes sont calculees sur l arbre de travail : il doit etre identique au commit publie."""
    head = subprocess.run(['git', 'rev-parse', commit], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip()
    if head != commit:
        raise SystemExit('Le commit indique est introuvable ou abrege : ' + commit)
    actuel = subprocess.run(['git', 'rev-parse', 'HEAD'], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip()
    if actuel != commit:
        raise SystemExit('HEAD (' + actuel[:7] + ') differe du commit a publier (' + commit[:7] + ') : extraire ce commit avant de generer.')
    statut = subprocess.run(['git', 'status', '--porcelain', '--', str(VERSION)], cwd=ROOT, capture_output=True, text=True, check=True).stdout
    if statut.strip():
        raise SystemExit('Modifications locales non commitees dans versions/ : le lanceur ne correspondrait a aucun commit.')
    suivis = set(subprocess.run(['git', 'ls-files', '--', str(VERSION)], cwd=ROOT, capture_output=True, text=True, check=True).stdout.split('\n')) - {''}
    locaux = {p.relative_to(ROOT).as_posix() for p in VERSION.rglob('*') if p.is_file() and '__pycache__' not in p.parts and '.pytest_cache' not in p.parts}
    if locaux - suivis:
        raise SystemExit('Fichiers non suivis par git dans versions/ : ' + ', '.join(sorted(locaux - suivis)))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('commit')
    args = parser.parse_args()
    verifier_commit(args.commit)
    code = payload(args.commit)
    target = ROOT / 'Installateur'
    # Le payload ne comporte que de l ASCII : compatible cmd.exe et PowerShell 5.1.
    (target / 'Demarrer_Installation_Cabinet.ps1').write_bytes(code.replace('\r\n', '\n').replace('\n', '\r\n').encode('ascii'))
    (target / 'Demarrer_Installation_Cabinet.cmd').write_bytes((HEADER + code).replace('\r\n', '\n').replace('\n', '\r\n').encode('ascii'))
    print('Lanceur autonome genere pour', args.commit)

if __name__ == '__main__':
    main()
