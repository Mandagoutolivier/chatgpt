@echo off
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -File "%~dp0Installer.ps1" %*
if errorlevel 1 echo Installation interrompue. Consulter le message et le journal.
pause
