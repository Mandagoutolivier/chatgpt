[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Source,
    [switch]$RestaurerDisposition
)

$ErrorActionPreference = 'Stop'
$machine = [Environment]::MachineName
$identite = [Security.Principal.WindowsIdentity]::GetCurrent().Name
if ($machine -ne 'RDC' -or $identite -ne 'RDC\PATRICIA') {
    throw 'Ce script doit etre execute uniquement dans la session RDC\PATRICIA.'
}

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Execution avec elevation administrateur requise pour restaurer aussi le Bureau public.'
}

$manifestPath = Join-Path $Source 'manifest-bureau.json'
if (-not (Test-Path -LiteralPath $manifestPath)) { throw 'Manifest du Bureau source absent.' }
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($fichier in @($manifest.Fichiers)) {
    $chemin = Join-Path $Source ([string]$fichier.CheminRelatif)
    if (-not (Test-Path -LiteralPath $chemin)) { throw "Fichier source absent : $chemin" }
    if ((Get-FileHash -LiteralPath $chemin -Algorithm SHA256).Hash -ne $fichier.SHA256) {
        throw "Empreinte source incorrecte : $chemin"
    }
}

$date = Get-Date -Format 'yyyyMMddHHmmss'
$racineSauvegarde = Join-Path $env:LOCALAPPDATA ('MigrationSecretariat\AvantBureau-' + $date)
$avantUtilisateur = Join-Path $racineSauvegarde 'Utilisateur'
$avantPublic = Join-Path $racineSauvegarde 'Public'
New-Item -ItemType Directory -Path $avantUtilisateur, $avantPublic -Force | Out-Null
$bureauUtilisateur = [Environment]::GetFolderPath('Desktop')
$bureauPublic = [Environment]::GetFolderPath('CommonDesktopDirectory')

function Copier-Bureau([string]$Origine, [string]$Destination) {
    if (-not (Test-Path -LiteralPath $Origine)) { return }
    & robocopy.exe $Origine $Destination /E /COPY:DAT /DCOPY:DAT /R:1 /W:1 /XJ /NP /NFL /NDL | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "Echec robocopy $Origine (code $LASTEXITCODE)." }
}

Copier-Bureau $bureauUtilisateur $avantUtilisateur
Copier-Bureau $bureauPublic $avantPublic
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
        & reg.exe export $cles[$i].Reg (Join-Path $racineSauvegarde ('Disposition-avant-' + ($i + 1) + '.reg')) /y *> $null
    }
}

Copier-Bureau (Join-Path $Source 'Utilisateur') $bureauUtilisateur
Copier-Bureau (Join-Path $Source 'Public') $bureauPublic

$erreursCopie = @()
foreach ($fichier in @($manifest.Fichiers | Where-Object { $_.CheminRelatif -notlike 'Disposition-*.reg' })) {
    $relatif = [string]$fichier.CheminRelatif
    if ($relatif -like 'Utilisateur\*') {
        $cible = Join-Path $bureauUtilisateur $relatif.Substring('Utilisateur\'.Length)
    } elseif ($relatif -like 'Public\*') {
        $cible = Join-Path $bureauPublic $relatif.Substring('Public\'.Length)
    } else { continue }
    if (-not (Test-Path -LiteralPath $cible) -or
        (Get-FileHash -LiteralPath $cible -Algorithm SHA256).Hash -ne $fichier.SHA256) {
        $erreursCopie += $relatif
    }
}
if ($erreursCopie.Count) { throw ('Copie non conforme : ' + ($erreursCopie -join ', ')) }

$raccourcis = @()
$shell = New-Object -ComObject WScript.Shell
foreach ($lnk in @(Get-ChildItem -LiteralPath $bureauUtilisateur, $bureauPublic -Filter '*.lnk' -File -Force -ErrorAction SilentlyContinue)) {
    $r = $shell.CreateShortcut($lnk.FullName)
    $cibleExiste = $true
    if ($r.TargetPath -and $r.TargetPath -notmatch '^[a-z]+:' -and $r.TargetPath -notlike '\\*') {
        $cibleExiste = Test-Path -LiteralPath $r.TargetPath
    } elseif ($r.TargetPath -and ($r.TargetPath -like '*:\*' -or $r.TargetPath -like '\\*')) {
        $cibleExiste = Test-Path -LiteralPath $r.TargetPath
    }
    $raccourcis += [pscustomobject]@{ Nom=$lnk.Name; Cible=$r.TargetPath; CibleAccessible=$cibleExiste }
}

$resolutionSource = @($manifest.Affichage | Select-Object -First 1)
$resolutionCible = @(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue |
    Select-Object -First 1 Name, CurrentHorizontalResolution, CurrentVerticalResolution, CurrentRefreshRate)
$dpiCible = (Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name AppliedDPI -ErrorAction SilentlyContinue).AppliedDPI
$affichageCompatible = $resolutionSource.Count -eq 1 -and $resolutionCible.Count -eq 1 -and
    $resolutionSource[0].CurrentHorizontalResolution -eq $resolutionCible[0].CurrentHorizontalResolution -and
    $resolutionSource[0].CurrentVerticalResolution -eq $resolutionCible[0].CurrentVerticalResolution -and
    $manifest.MiseAEchelleDPI -eq $dpiCible
$dispositionRestauree = $false
if ($RestaurerDisposition -and $affichageCompatible) {
    foreach ($reg in @(Get-ChildItem -LiteralPath $Source -Filter 'Disposition-*.reg' -File)) {
        & reg.exe import $reg.FullName *> $null
        if ($LASTEXITCODE -ne 0) { throw "Echec de restauration de $($reg.Name)." }
    }
    $dispositionRestauree = $true
}

$rapport = [ordered]@{
    Statut = 'RESTAURATION_BUREAU_OK'
    Date = (Get-Date -Format o)
    Source = $Source
    SauvegardeAvant = $racineSauvegarde
    FichiersVerifies = @($manifest.Fichiers).Count
    Raccourcis = $raccourcis
    AffichageCompatible = $affichageCompatible
    DispositionDemandee = [bool]$RestaurerDisposition
    DispositionRestauree = $dispositionRestauree
    DeconnexionRequise = $dispositionRestauree
}
$rapportPath = Join-Path $env:LOCALAPPDATA ('MigrationSecretariat\restauration-bureau-' + $date + '.json')
$rapport | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $rapportPath -Encoding UTF8
$rapport | ConvertTo-Json -Compress -Depth 7
