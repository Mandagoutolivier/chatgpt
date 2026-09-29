[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Destination
)

$ErrorActionPreference = 'Stop'
$sourceUtilisateur = [Environment]::GetFolderPath('Desktop')
$sourcePublic = [Environment]::GetFolderPath('CommonDesktopDirectory')
$racine = Join-Path $Destination ('Bureau-' + [Environment]::MachineName + '-' + (Get-Date -Format 'yyyyMMddHHmmss'))
$destUtilisateur = Join-Path $racine 'Utilisateur'
$destPublic = Join-Path $racine 'Public'
New-Item -ItemType Directory -Path $destUtilisateur, $destPublic -Force | Out-Null

function Copier-Bureau([string]$Source, [string]$Cible) {
    if (-not (Test-Path -LiteralPath $Source)) { return }
    & robocopy.exe $Source $Cible /E /COPY:DAT /DCOPY:DAT /R:1 /W:1 /XJ /NP /NFL /NDL | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "Echec robocopy $Source (code $LASTEXITCODE)." }
}

Copier-Bureau $sourceUtilisateur $destUtilisateur
Copier-Bureau $sourcePublic $destPublic

$cles = @(
    [pscustomobject]@{
        Reg = 'HKCU\Software\Microsoft\Windows\Shell\Bags\1\Desktop'
        PSPath = 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Bags\1\Desktop'
    },
    [pscustomobject]@{
        Reg = 'HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\1\Desktop'
        PSPath = 'Registry::HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\1\Desktop'
    }
)
for ($i = 0; $i -lt $cles.Count; $i++) {
    if (Test-Path -LiteralPath $cles[$i].PSPath) {
        & reg.exe export $cles[$i].Reg (Join-Path $racine ('Disposition-' + ($i + 1) + '.reg')) /y *> $null
    }
}

$fichiers = @(Get-ChildItem -LiteralPath $racine -File -Recurse -Force)
$manifest = [ordered]@{
    Date = (Get-Date -Format o)
    Source = [Environment]::MachineName
    Compte = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    BureauUtilisateur = $sourceUtilisateur
    BureauPublic = $sourcePublic
    Fichiers = @($fichiers | ForEach-Object {
        [pscustomobject]@{
            CheminRelatif = $_.FullName.Substring($racine.Length).TrimStart('\')
            Octets = $_.Length
            SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        }
    })
    Affichage = @(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue |
        Select-Object Name, CurrentHorizontalResolution, CurrentVerticalResolution, CurrentRefreshRate)
    MiseAEchelleDPI = (Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' `
        -Name AppliedDPI -ErrorAction SilentlyContinue).AppliedDPI
}
$manifestPath = Join-Path $racine 'manifest-bureau.json'
$manifest | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

[pscustomobject]@{
    Statut = 'EXPORT_BUREAU_OK'
    Dossier = $racine
    Fichiers = $fichiers.Count
    Octets = [long](($fichiers | Measure-Object Length -Sum).Sum)
    ManifestSHA256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
} | ConvertTo-Json -Compress
