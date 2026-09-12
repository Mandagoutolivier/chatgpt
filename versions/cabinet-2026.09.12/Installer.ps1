[CmdletBinding()]
param(
    [ValidateSet('Domicile','Secretariat','Cabinet')][string]$Profil,
    [ValidateSet('Preparation','Installation')][string]$Mode='Preparation',
    [string]$RacineNas='',
    [string]$UrlService='',
    [string]$DossierPrepare='',
    [string]$DossierGdt='C:\Mandagout',
    [string]$SqliteExe='',
    [string]$FichierJeton=''
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Build\outils_assistant.ps1')
if (-not $Profil) {
    Write-Host 'Installation Cabinet Cardio - version 2026.09.12'
    Write-Host '1 - Domicile : secretariat et medecin sur ce PC'
    Write-Host '2 - Secretariat : agenda, patients, actes et impressions'
    Write-Host '3 - Cabinet : poste medecin, Word, Dragon et ECG'
    do { $choice=Read-Host 'Votre choix (1, 2 ou 3)' } until ($choice -in @('1','2','3'))
    $Profil=@{'1'='Domicile';'2'='Secretariat';'3'='Cabinet'}[$choice]
}
if (-not $RacineNas) {
    $RacineNas=Choisir-RacineNasAssistant '' '' (Join-Path $env:APPDATA 'CabinetCardio\chemin.txt')
    if (-not $RacineNas) { Write-Host 'Installation en pause.';return }
}
$profiles=@{Domicile='Domicile';Secretariat='CabinetSecretariat';Cabinet='CabinetMedecin'}
$arguments=@{Profil=$profiles[$Profil];Mode=$Mode;RacineNas=$RacineNas;DossierGdt=$DossierGdt}
foreach ($key in @('UrlService','DossierPrepare','SqliteExe','FichierJeton')) {
    $value=Get-Variable -Name $key -ValueOnly
    if ($value) { $arguments[$key]=$value }
}
& (Join-Path $PSScriptRoot 'Build\installer_multi_postes.ps1') @arguments
if ($Mode -eq 'Preparation') {
    Write-Host 'Apres compilation et recette Office, activer ces memes fichiers avec :'
    Write-Host ".\Installer.ps1 -Profil $Profil -Mode Installation -DossierPrepare 'CHEMIN_AFFICHE_CI_DESSUS'"
}
