"""Assembler le CMD depuis les objets du commit, jamais depuis l'arbre local."""
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
    def git(*args):
        return subprocess.check_output(['git', '-C', str(ROOT), *args])
    if git('rev-parse', commit+'^{commit}').decode().strip() != commit:
        raise ValueError('Le SHA doit identifier un commit exact')
    prefix = VERSION.relative_to(ROOT).as_posix() + '/'
    hashes = {}
    for entry in git('ls-tree', '-rz', commit, '--', prefix).split(b'\0'):
        if not entry:
            continue
        metadata, path = entry.split(b'\t', 1)
        mode, kind, sha = metadata.split()
        if kind != b'blob' or mode not in (b'100644', b'100755'):
            raise ValueError('Seuls les fichiers ordinaires sont admis dans la livraison')
        relative = path.decode()[len(prefix):]
        hashes[relative] = hashlib.sha256(git('cat-file', 'blob', sha.decode())).hexdigest()
    if not hashes:
        raise ValueError('Version absente de ce commit')
    helpers = git('show', commit+':'+prefix+'Build/outils_telechargement.ps1').decode('utf-8-sig')
    main = git('show', commit+':Installateur/lanceur.template.ps1').decode('utf-8-sig')
    main = main.replace('@@COMMIT@@', commit).replace('@@HASHES@@', json.dumps(hashes, ensure_ascii=True, indent=2))
    return helpers + '\n' + main

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('commit')
    args = parser.parse_args()
    code = payload(args.commit)
    target = ROOT / 'Installateur'
    # Le payload ne comporte que de l ASCII : compatible cmd.exe et PowerShell 5.1.
    (target / 'Demarrer_Installation_Cabinet.ps1').write_bytes(code.replace('\r\n', '\n').replace('\n', '\r\n').encode('ascii'))
    (target / 'Demarrer_Installation_Cabinet.cmd').write_bytes((HEADER + code).replace('\r\n', '\n').replace('\n', '\r\n').encode('ascii'))
    print('Lanceur autonome genere pour', args.commit)

if __name__ == '__main__':
    main()
