@echo off
setlocal
cd /d "%~dp0"
if not "%~1"=="" goto manuel
powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0Build\assistant_installation.ps1"
goto resultat
:manuel
powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0Installer.ps1" %*
:resultat
set "CABINET_RESULT=%ERRORLEVEL%"
if not "%CABINET_RESULT%"=="0" echo Installation interrompue. Consultez le message ci-dessus.
pause
exit /b %CABINET_RESULT%
