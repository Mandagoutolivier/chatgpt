[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Domicile','CabinetSecretariat','CabinetMedecin')]
    [string]$Profil,

    [string]$RacineNas = '\\DS224\CabinetCardio',
    [string]$DossierGdt = 'C:\Mandagout',
    [string]$DossierSources = (Join-Path (Split-Path $PSScriptRoot -Parent) 'ModelesSource'),
    [string]$SqliteExe = '',
    [switch]$ConserverModelesConstruits
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')

$racinePaquet = Split-Path $PSScriptRoot -Parent
$installerWord = $Profil -in @('Domicile','CabinetMedecin')
$installerSecretariat = $Profil -in @('Domicile','CabinetSecretariat')
$horodatage = Get-Date -Format 'yyyyMMdd-HHmmss'
$dossierLocal = Join-Path $env:APPDATA 'CabinetCardio'
$dossierSauvegarde = Join-Path $dossierLocal ("Sauvegardes\" + $horodatage)
$dossierTravail = Join-Path $env:TEMP ("CabinetCardio-Installation-" + [guid]::NewGuid().ToString('N'))
$journal = Join-Path $dossierLocal ("installation-" + $horodatage + '.log')

New-Item -ItemType Directory -Force -Path $dossierLocal,$dossierSauvegarde,$dossierTravail | Out-Null
Start-Transcript -Path $journal -Force | Out-Null

function Etape([string]$texte) { Write-Host ''; Write-Host ("=== " + $texte) -ForegroundColor Cyan }
function Ok([string]$texte) { Write-Host ("  OK  " + $texte) -ForegroundColor Green }
function Sauvegarder-SiPresent([string]$chemin) {
    if (Test-Path $chemin) {
        Copy-Item $chemin (Join-Path $dossierSauvegarde ([IO.Path]::GetFileName($chemin))) -Force
        Ok "sauvegarde de $chemin"
    }
}
function Trouver-Source([string[]]$noms) {
    foreach ($nom in $noms) {
        $candidat = Join-Path $DossierSources $nom
        if (Test-Path $candidat) { return (Resolve-Path $candidat).Path }
    }
    throw "Source absente dans $DossierSources : $($noms -join ' ou ')"
}
function Tester-Office([string]$progId) {
    $application = $null
    try {
        $application = New-Object -ComObject $progId
        $application.Quit()
        Ok "$progId disponible"
    } finally {
        if ($null -ne $application) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) | Out-Null } catch {} }
    }
}
function Obtenir-SqliteOfficiel() {
    Etape 'Telechargement de SQLite depuis sqlite.org'
    $page = (Invoke-WebRequest -UseBasicParsing -Uri 'https://sqlite.org/download.html').Content
    $ligneProduit = ($page -split "`n" | Where-Object {
        $_ -match 'PRODUCT,[0-9.]+,[0-9]{4}/sqlite-tools-win-x64-[0-9]+\.zip,'
    } | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($ligneProduit)) {
        throw 'Impossible de trouver le paquet Windows x64 dans la page officielle SQLite.'
    }
    if ($ligneProduit -notmatch 'PRODUCT,[0-9.]+,(?<url>[0-9]{4}/sqlite-tools-win-x64-[0-9]+\.zip),') {
        throw 'L’adresse du paquet SQLite n’a pas pu etre extraite.'
    }
    $url = 'https://sqlite.org/' + $Matches['url']
    $zipSqlite = Join-Path $dossierTravail 'sqlite-tools.zip'
    $extraction = Join-Path $dossierTravail 'sqlite-tools'
    Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $zipSqlite
    Expand-Archive -Path $zipSqlite -DestinationPath $extraction -Force
    $exe = Get-ChildItem $extraction -Recurse -Filter sqlite3.exe | Select-Object -First 1
    if ($null -eq $exe) { throw 'Le paquet officiel SQLite ne contient pas sqlite3.exe.' }
    Ok "SQLite telecharge depuis $url"
    return $exe.FullName
}

try {
    Etape "Controles prealables — profil $Profil"
    if (Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue) {
        throw 'Fermez completement Word et Excel, puis relancez le script.'
    }
    if (-not (Test-Path $RacineNas)) {
        if ($Profil -eq 'Domicile') {
            throw "NAS inaccessible : $RacineNas. Connectez d'abord le VPN Cabinet Freebox Pro."
        }
        throw "NAS inaccessible : $RacineNas. Verifiez le reseau et le partage Synology."
    }
    $testEcriture = Join-Path $RacineNas ('.test-install-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllText($testEcriture, 'test', [Text.Encoding]::ASCII)
        [IO.File]::Delete($testEcriture)
        Ok "lecture/ecriture NAS : $RacineNas"
    } catch { throw "Le NAS est lisible mais non inscriptible : $($_.Exception.Message)" }
    if ($installerWord) { Tester-Office 'Word.Application' }
    if ($installerSecretariat) { Tester-Office 'Excel.Application' }

    Etape 'Protection des donnees et preparation du NAS'
    # Aucun fichier de base existant n'est remplace. Les dossiers manquants
    # sont crees ; une copie datee de la configuration est prise si elle existe.
    foreach ($relatif in @('Base','Actes','Patients','Config','Modeles','Modeles\Deploy',
                            'Echange','Echange\Arrives','Echange\Arrives\Pris',
                            'Echange\AEnvoyer','Echange\Traites','Sauvegardes','Logs')) {
        New-Item -ItemType Directory -Force -Path (Join-Path $RacineNas $relatif) | Out-Null
    }
    $configNas = Join-Path $RacineNas 'Config\config.ini'
    if (Test-Path $configNas) {
        Copy-Item $configNas (Join-Path $RacineNas ("Sauvegardes\config-avant-installation-$horodatage.ini")) -Force
        Ok 'configuration NAS sauvegardee'
    } else {
        $configDefaut = Join-Path $racinePaquet 'Src\ConfigDefaut\config.ini'
        if (-not (Test-Path $configDefaut)) { throw "Configuration initiale absente : $configDefaut" }
        Copy-Item $configDefaut $configNas
        Ok 'configuration NAS initialisee'
    }
    $RacineNas | Out-File (Join-Path $dossierLocal 'chemin.txt') -Encoding ASCII -Force
    New-Item -ItemType Directory -Force -Path $DossierGdt | Out-Null
    Ok "chemin.txt -> $RacineNas"
    Ok "dossier GDT local : $DossierGdt"

    if ($installerWord) {
        Etape 'Construction et installation de la partie medecin'
        $prod6 = Trouver-Source @('ModeleCourrierChatGPT_PROD(6).dotm','ModeleCourrierChatGPT_PROD6.dotm')
        $cabinet1 = Trouver-Source @('Cabinet(1).dotm','Cabinet1.dotm')
        $modeleConstruit = Join-Path $dossierTravail 'CabinetUnifie_TEST.dotm'
        & (Join-Path $PSScriptRoot 'construire_modele_unifie.ps1') `
            -Prod6 $prod6 -Cabinet1 $cabinet1 -Sortie $modeleConstruit -RacineSources $racinePaquet
        if (-not (Test-Path $modeleConstruit)) { throw 'Le constructeur Word n’a produit aucun modele.' }

        $startupWord = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
        $modeleInstalle = Join-Path $startupWord 'CabinetUnifie.dotm'
        New-Item -ItemType Directory -Force -Path $startupWord | Out-Null
        Sauvegarder-SiPresent $modeleInstalle
        Copy-Item $modeleConstruit $modeleInstalle -Force
        Ok "modele medecin installe : $modeleInstalle"

        if ([string]::IsNullOrWhiteSpace($SqliteExe)) {
            $sourceSqlite = Join-Path $DossierSources 'sqlite3.exe'
            if (Test-Path $sourceSqlite) { $SqliteExe = $sourceSqlite }
            else {
                $commandeSqlite = Get-Command sqlite3.exe -ErrorAction SilentlyContinue
                if ($null -ne $commandeSqlite) { $SqliteExe = $commandeSqlite.Source }
            }
        }
        if ([string]::IsNullOrWhiteSpace($SqliteExe) -or -not (Test-Path $SqliteExe)) { $SqliteExe = Obtenir-SqliteOfficiel }
        & (Join-Path $PSScriptRoot 'installer_sqlite_medecin.ps1') -SqliteExe $SqliteExe
    }

    if ($installerSecretariat) {
        Etape 'Construction et installation de la partie secretariat'
        $cabinetXlsm = Trouver-Source @('Cabinet.xlsm')
        $excelConstruit = Join-Path $dossierTravail 'CabinetSecretariat_TEST.xlsm'
        & (Join-Path $PSScriptRoot 'construire_cabinet_secretariat.ps1') `
            -CabinetXlsm $cabinetXlsm -Sortie $excelConstruit -RacineSources $racinePaquet
        if (-not (Test-Path $excelConstruit)) { throw 'Le constructeur Excel n’a produit aucun classeur.' }

        $dossierExcel = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CabinetCardio'
        $excelInstalle = Join-Path $dossierExcel 'Cabinet.xlsm'
        New-Item -ItemType Directory -Force -Path $dossierExcel | Out-Null
        Sauvegarder-SiPresent $excelInstalle
        Copy-Item $excelConstruit $excelInstalle -Force
        $bureau = [Environment]::GetFolderPath('Desktop')
        $raccourci = Join-Path $bureau 'Cabinet Cardio.lnk'
        $wsh = New-Object -ComObject WScript.Shell
        $lien = $wsh.CreateShortcut($raccourci)
        $lien.TargetPath = $excelInstalle
        $lien.WorkingDirectory = $dossierExcel
        $lien.Save()
        Ok "application secretariat installee : $excelInstalle"
    }

    Etape 'Verification finale'
    if ($installerWord -and -not (Test-Path (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\CabinetUnifie.dotm'))) {
        throw 'Modele Word absent apres installation.'
    }
    if ($installerSecretariat -and -not (Test-Path (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CabinetCardio\Cabinet.xlsm'))) {
        throw 'Application secretariat absente apres installation.'
    }
    if ($ConserverModelesConstruits) {
        $destinationTests = Join-Path $dossierLocal ("ModelesConstruits\" + $horodatage)
        New-Item -ItemType Directory -Force -Path $destinationTests | Out-Null
        Copy-Item (Join-Path $dossierTravail '*') $destinationTests -Force
        Ok "copies de test conservees : $destinationTests"
    }
    Write-Host ''
    Write-Host "INSTALLATION TERMINEE — $Profil" -ForegroundColor Green
    Write-Host "Sauvegardes locales : $dossierSauvegarde"
    Write-Host "Journal : $journal"
    Write-Host 'Avant production : compiler les projets VBA et tester sur un patient fictif.' -ForegroundColor Yellow
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    if (Test-Path $dossierTravail) { Remove-Item $dossierTravail -Recurse -Force -ErrorAction SilentlyContinue }
}
