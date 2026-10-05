@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Deploy-C5-NP10-AX8-ACCUEIL.ps1" %*
set "C5ExitCode=%errorlevel%"
echo.
if not "%C5ExitCode%"=="0" echo Operation interrompue. Lire le message ci-dessus.
pause
exit /b %C5ExitCode%
