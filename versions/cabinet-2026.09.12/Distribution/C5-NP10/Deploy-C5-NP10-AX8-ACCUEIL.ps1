[CmdletBinding()]
param(
    [switch]$Execute,
    [string]$ConfirmerBascule = '',
    [string]$DossierPaquet = '\\DS224\CabinetCardio-Dev\Deploy\C5-NP10',
    [string]$SourceWord = '',
    [string]$SourceExcel = '',
    [string]$AccueilAdmin = '\\ACCUEIL\C$',
    [string]$AccueilUtilisateur = 'accueil',
    [string]$AncienComplementWord = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $SourceWord) { $SourceWord = Join-Path $PSScriptRoot 'CabinetDragonC5NP10.dotm' }
if (-not $SourceExcel) { $SourceExcel = Join-Path $PSScriptRoot 'Cabinet.xlsm' }

# C5-NP10 : Word a evolue ; Excel est volontairement le binaire C3 qualifie,
# aucun module Excel n'ayant change entre f1e7d2e et f2e74f5.
$Revision = 'C5-NP10'
$Commit = 'f2e74f52e145f3e07480ea9bbeac3bda3fafe346'
$HashWord = 'FF183F42D591BE91CEE2146AF2260439157957C3A5D4FD13D9328B31F19C2C7C'
$HashExcel = '454670F1CB00DD2929EA09D9E6AA0243E68C43E6478B279979E836AFB1FECD24'

function Assert-Hash([string]$Path, [string]$Expected, [string]$Label) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label absent : $Path" }
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($Expected -and $actual -ne $Expected) { throw "$Label : empreinte inattendue ($actual)." }
    return $actual
}
function Assert-OfficeFerme([string]$ProcessName) {
    $session = (Get-Process -Id $PID).SessionId
    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session }) {
        throw "$ProcessName est ouvert dans cette session : fermer Office avant le deploiement."
    }
}
function Write-Shortcut([string]$Path, [string]$Target, [string]$Arguments, [string]$WorkingDirectory) {
    $shell = New-Object -ComObject WScript.Shell
    try {
        $shortcut = $shell.CreateShortcut($Path)
        $shortcut.TargetPath = $Target
        $shortcut.Arguments = $Arguments
        $shortcut.WorkingDirectory = $WorkingDirectory
        $shortcut.IconLocation = "$Target,0"
        $shortcut.Save()
    } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell) }
}

if ((whoami) -ne 'ax8_max\olivi') { throw 'Ce script doit etre lance dans AX8_MAX\\olivi.' }
Assert-OfficeFerme 'WINWORD'
Assert-OfficeFerme 'EXCEL'
# Controle distant en lecture seule : en cas de refus DCOM, arret avant toute copie.
$accProcesses = @(Get-WmiObject -Class Win32_Process -ComputerName 'ACCUEIL' -Filter "Name='EXCEL.EXE' OR Name='WINWORD.EXE'" -ErrorAction Stop)
if ($accProcesses.Count) { throw 'Fermez Word et Excel sur ACCUEIL avant le deploiement.' }

$wordHash = Assert-Hash $SourceWord $HashWord 'Modele Word C5-NP10'
$excelHash = Assert-Hash $SourceExcel $HashExcel 'Classeur Excel qualifie'

if (-not (Test-Path -LiteralPath $AccueilAdmin -PathType Container)) { throw "ACCUEIL inaccessible : $AccueilAdmin" }
$configAccueil = Join-Path $AccueilAdmin ('Users\' + $AccueilUtilisateur + '\AppData\Roaming\CabinetCardio')
foreach ($config in @((Join-Path $env:APPDATA 'CabinetCardio'), $configAccueil)) {
  foreach ($name in 'chemin.txt','service.url','service.token','poste.ini') {
    if (-not (Test-Path -LiteralPath (Join-Path $config $name) -PathType Leaf)) {
        throw "Configuration de production absente : $config / $name. Aucun fichier n'a ete copie."
    }
  }
  if ([IO.File]::ReadAllText((Join-Path $config 'chemin.txt')).Trim().TrimEnd('\') -ine '\\DS224\CabinetCardio') {
    throw "La racine de production attendue n'est pas configuree : $config"
  }
  $url = [IO.File]::ReadAllText((Join-Path $config 'service.url')).Trim()
  if ($url -notmatch '^https://' -or $url.TrimEnd('/') -ieq 'https://DS224:8444') {
    throw "Adresse de production invalide ou adresse de recette : $config"
  }
}

$date = Get-Date -Format 'yyyyMMdd-HHmmss'
$manifest = [ordered]@{
    revision = $Revision; sourceCommit = $Commit; created = (Get-Date).ToString('o')
    word = [ordered]@{ source=$SourceWord; sha256=$wordHash; name='CabinetDragonC5NP10.dotm' }
    excel = [ordered]@{ source=$SourceExcel; sha256=$excelHash; name='Cabinet.xlsm' }
    targets = @('AX8_MAX\olivi', ('ACCUEIL\'+$AccueilUtilisateur))
    execute = [bool]$Execute
}

if (-not $Execute) {
    $manifest | ConvertTo-Json -Depth 6
    Write-Host 'CONTROLE REUSSI. Relancer avec -Execute -ConfirmerBascule DEPLOYER-C5-NP10 pour copier les fichiers.' -ForegroundColor Yellow
    exit 0
}
if ($ConfirmerBascule -cne 'DEPLOYER-C5-NP10') { throw 'Confirmation absente ou incorrecte : aucune installation realisee.' }
if (-not $AncienComplementWord) { throw 'Indiquer -AncienComplementWord avec le chemin exact du complement Word actuellement actif.' }
if (-not (Test-Path -LiteralPath $AncienComplementWord -PathType Leaf)) { throw "Ancien complement Word introuvable : $AncienComplementWord" }
if ($AncienComplementWord -ieq $SourceWord) { throw 'La source C5 ne peut pas etre le complement a sauvegarder.' }

# Paquet NAS : copie versionnee, jamais dans CabinetCardio clinique.
$paquet = Join-Path $DossierPaquet $date
New-Item -ItemType Directory -Path $paquet -Force | Out-Null
Copy-Item -LiteralPath $SourceWord -Destination (Join-Path $paquet 'CabinetDragonC5NP10.dotm') -ErrorAction Stop
Copy-Item -LiteralPath $SourceExcel -Destination (Join-Path $paquet 'Cabinet.xlsm') -ErrorAction Stop
[void](Assert-Hash (Join-Path $paquet 'CabinetDragonC5NP10.dotm') $HashWord 'Copie Word NAS')
[void](Assert-Hash (Join-Path $paquet 'Cabinet.xlsm') $HashExcel 'Copie Excel NAS')
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $paquet 'manifest-deploiement.json') -Encoding UTF8

# AX8 : l'ancien complement est deplace dans une sauvegarde horodatee, puis C5 est
# installe dans un dossier versionne. Aucun fichier Normal.dotm ni configuration API n'est modifie.
$axRoot = 'C:\CabinetCardio\C5-NP10-' + $date
$axBackup = Join-Path $axRoot ('Sauvegardes\'+$date)
New-Item -ItemType Directory -Path $axBackup -Force | Out-Null
Copy-Item -LiteralPath $AncienComplementWord -Destination (Join-Path $axBackup (Split-Path $AncienComplementWord -Leaf)) -ErrorAction Stop
Copy-Item -LiteralPath (Join-Path $paquet 'CabinetDragonC5NP10.dotm') -Destination (Join-Path $axRoot 'CabinetDragonC5NP10.dotm') -Force

$lanceurWord = @'
$ErrorActionPreference='Stop'
$model=Join-Path $PSScriptRoot 'CabinetDragonC5NP10.dotm'
if(-not (Test-Path -LiteralPath $model)){throw 'Modele C5 absent.'}
$word=New-Object -ComObject Word.Application
$word.Visible=$true
foreach($a in $word.AddIns){if($a.Installed -and $a.Name -match 'Cabinet|ModeleCourrierChatGPT'){throw 'Un ancien complement Cabinet est encore charge : fermer Word, le desactiver puis relancer.'}}
$addin=$word.AddIns.Add($model,$true)
$word.Run('modRaccourcis.AutoExec')
[void]$word.Run('modAttenteLocale.SynchroniserAttentes')
[void]$word.Run('modFileArrivees.Unifie_AfficherFileArrivees')
'@
Set-Content -LiteralPath (Join-Path $axRoot 'Lancer-Cabinet-C5NP10.ps1') -Value $lanceurWord -Encoding UTF8
Write-Shortcut (Join-Path $env:USERPROFILE ('Desktop\Cabinet C5-NP10 - Medecin-'+$date+'.lnk')) "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" ('-NoProfile -File "'+(Join-Path $axRoot 'Lancer-Cabinet-C5NP10.ps1')+'"') $axRoot

# ACCUEIL : installation cote a cote. L'ancien classeur et ses raccourcis ne sont pas modifies.
$accLocalRoot = 'C:\CabinetCardio\C5-NP10-' + $date
$accRoot = Join-Path $AccueilAdmin ('CabinetCardio\C5-NP10-'+$date)
New-Item -ItemType Directory -Path $accRoot -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $paquet 'Cabinet.xlsm') -Destination (Join-Path $accRoot 'Cabinet.xlsm') -Force
$accCmd = Join-Path $accRoot 'Lancer-Cabinet-C5NP10.cmd'
Set-Content -LiteralPath $accCmd -Value ('@echo off' + [Environment]::NewLine + 'start "" "'+$accLocalRoot+'\Cabinet.xlsm"' + [Environment]::NewLine) -Encoding ASCII
Write-Shortcut (Join-Path $AccueilAdmin ('Users\'+$AccueilUtilisateur+'\Desktop\Cabinet C5-NP10 - Secretariat-'+$date+'.lnk')) (Join-Path $accLocalRoot 'Lancer-Cabinet-C5NP10.cmd') '' $accLocalRoot
[void](Assert-Hash (Join-Path $axRoot 'CabinetDragonC5NP10.dotm') $HashWord 'Copie Word AX8')
[void](Assert-Hash (Join-Path $accRoot 'Cabinet.xlsm') $HashExcel 'Copie Excel ACCUEIL')

# Desactivation de l'ancien complement uniquement apres verification des deux copies.
$ancienHash = (Get-FileHash -LiteralPath $AncienComplementWord -Algorithm SHA256).Hash
[void](Assert-Hash (Join-Path $axBackup (Split-Path $AncienComplementWord -Leaf)) $ancienHash 'Sauvegarde ancien complement')
Move-Item -LiteralPath $AncienComplementWord -Destination ($AncienComplementWord+'.desactive-'+$date) -ErrorAction Stop
$manifest['ancienComplement'] = $AncienComplementWord
$manifest['ancienComplementDesactive'] = $AncienComplementWord+'.desactive-'+$date
$manifest['sauvegardeWord'] = $axBackup
$manifest['dossierAX8'] = $axRoot
$manifest['dossierAccueil'] = $accLocalRoot

$manifest['deployed'] = (Get-Date).ToString('o')
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $paquet 'manifest-deploiement.json') -Encoding UTF8
Write-Host "DEPLOIEMENT TERMINE. Paquet et journal : $paquet" -ForegroundColor Green
Write-Host 'Ouvrir ensuite les deux raccourcis C5, verifier la connexion et conserver les anciens raccourcis jusqu a validation.' -ForegroundColor Yellow
