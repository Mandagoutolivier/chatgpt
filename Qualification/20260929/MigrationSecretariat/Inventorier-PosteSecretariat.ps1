[CmdletBinding()]
param(
    [string]$DossierRapport = 'C:\Users\Public\MigrationSecretariat'
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $DossierRapport -Force | Out-Null

function Mesurer-Dossier([string]$Nom, [string]$Chemin) {
    $resultat = [ordered]@{ Nom = $Nom; Chemin = $Chemin; Existe = $false }
    if (-not (Test-Path -LiteralPath $Chemin)) { return [pscustomobject]$resultat }
    $resultat.Existe = $true
    $fichiers = @(Get-ChildItem -LiteralPath $Chemin -File -Recurse -Force -ErrorAction SilentlyContinue)
    $resultat.Fichiers = $fichiers.Count
    $resultat.Octets = [long](($fichiers | Measure-Object -Property Length -Sum).Sum)
    $resultat.Extensions = @($fichiers | Group-Object Extension | Sort-Object Count -Descending |
        Select-Object -First 20 @{n='Extension';e={ if ($_.Name) { $_.Name } else { '[sans extension]' } }}, Count)
    return [pscustomobject]$resultat
}

function Lire-Version([string]$Nom, [string[]]$Chemins) {
    foreach ($chemin in $Chemins) {
        if (Test-Path -LiteralPath $chemin) {
            $item = Get-Item -LiteralPath $chemin
            return [pscustomobject]@{
                Nom = $Nom
                Present = $true
                Chemin = $chemin
                Version = $item.VersionInfo.FileVersion
            }
        }
    }
    return [pscustomobject]@{ Nom = $Nom; Present = $false; Chemin = $null; Version = $null }
}

$identite = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$nomOrdinateur = [Environment]::MachineName
$ordinateur = Get-CimInstance Win32_ComputerSystem
$systeme = Get-CimInstance Win32_OperatingSystem
$applications = @(Get-ItemProperty `
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', `
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*', `
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' `
    -ErrorAction SilentlyContinue | Where-Object DisplayName |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallLocation |
    Sort-Object DisplayName, DisplayVersion -Unique)

$imprimantes = @(Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue |
    Select-Object Name, DriverName, PortName, Default, Network, Shared)
$lecteursReseau = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=4' -ErrorAction SilentlyContinue |
    Select-Object DeviceID, ProviderName)
$partages = @(Get-SmbShare -ErrorAction SilentlyContinue |
    Where-Object { -not $_.Special } | Select-Object Name, Path, Description)
$servicesSauvegarde = @(Get-Service -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match 'Synology|ActiveBackup' -or $_.DisplayName -match 'Synology|Active Backup' } |
    Select-Object Name, DisplayName, Status, StartType)

$startupWord = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$xlstart = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
$modeles = Join-Path $env:APPDATA 'Microsoft\Templates'
$outlook = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
$thunderbird = Join-Path $env:APPDATA 'Thunderbird'
$firefox = Join-Path $env:APPDATA 'Mozilla\Firefox'
$chrome = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'

$emplacements = @(
    Mesurer-Dossier 'Bureau utilisateur' ([Environment]::GetFolderPath('Desktop'))
    Mesurer-Dossier 'Documents utilisateur' ([Environment]::GetFolderPath('MyDocuments'))
    Mesurer-Dossier 'Word STARTUP' $startupWord
    Mesurer-Dossier 'Excel XLSTART' $xlstart
    Mesurer-Dossier 'Modeles Office' $modeles
    Mesurer-Dossier 'Outlook local' $outlook
    Mesurer-Dossier 'Thunderbird' $thunderbird
    Mesurer-Dossier 'Firefox' $firefox
    Mesurer-Dossier 'Chrome' $chrome
)

$logicielsMetier = @(
    Lire-Version 'Word' @('C:\Program Files\Microsoft Office\Office16\WINWORD.EXE','C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE')
    Lire-Version 'Excel' @('C:\Program Files\Microsoft Office\Office16\EXCEL.EXE','C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE')
    Lire-Version 'Resting 12-Lead' @('C:\Resting12Lead\Resting12Lead.exe','C:\Program Files\Resting12Lead\Resting12Lead.exe')
    Lire-Version 'Quick BP' @('C:\Program Files (x86)\Quick BP\QuickBP.exe','C:\QuickBP\QuickBP.exe')
    Lire-Version 'EasyScope' @('C:\Program Files (x86)\EasyScope\EasyScope.exe','C:\EasyScope\EasyScope.exe')
)

$rapport = [ordered]@{
    Date = (Get-Date -Format o)
    Etape = 'inventaire_source_avant_migration'
    ModificationEffectuee = $false
    DonneesCopiees = $false
    Identite = $identite
    Ordinateur = [ordered]@{
        Nom = $nomOrdinateur
        Fabricant = $ordinateur.Manufacturer
        Modele = $ordinateur.Model
        MemoireOctets = [long]$ordinateur.TotalPhysicalMemory
        Windows = $systeme.Caption
        Version = $systeme.Version
        Architecture = $systeme.OSArchitecture
    }
    AdressesIPv4 = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notlike '169.254.*' } |
        Select-Object IPAddress, PrefixLength, InterfaceAlias)
    Disques = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' |
        Select-Object DeviceID, VolumeName, FileSystem, Size, FreeSpace)
    Applications = $applications
    LogicielsMetier = $logicielsMetier
    Imprimantes = $imprimantes
    LecteursReseau = $lecteursReseau
    Partages = $partages
    ServicesSauvegarde = $servicesSauvegarde
    EmplacementsUtilisateur = $emplacements
    Exclusions = @('aucun mot de passe','aucun jeton','aucun nom de fichier patient','aucune copie de donnees')
}

$json = Join-Path $DossierRapport ('inventaire-source-' + $nomOrdinateur + '.json')
$rapport | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $json -Encoding UTF8
$empreinte = (Get-FileHash -LiteralPath $json -Algorithm SHA256).Hash
$resume = Join-Path $DossierRapport ('inventaire-source-' + $nomOrdinateur + '.txt')
@(
    'INVENTAIRE_SOURCE_OK'
    ('Ordinateur=' + $nomOrdinateur)
    ('Compte=' + $identite)
    ('Applications=' + $applications.Count)
    ('Imprimantes=' + $imprimantes.Count)
    ('Partages=' + $partages.Count)
    ('SHA256=' + $empreinte)
    ('Rapport=' + $json)
) | Set-Content -LiteralPath $resume -Encoding UTF8

Get-Content -LiteralPath $resume
