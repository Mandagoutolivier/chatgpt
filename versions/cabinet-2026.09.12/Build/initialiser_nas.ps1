[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$RacineNas,
    [string]$RacineSources = (Split-Path $PSScriptRoot -Parent)
)
. (Join-Path $PSScriptRoot 'outils_construction.ps1')
if ($RacineNas -notmatch '^\\\\[^\\]+\\[^\\]+') { throw 'La racine doit etre un chemin UNC du Synology.' }
if (-not (Test-Path -LiteralPath $RacineNas -PathType Container)) { throw "NAS inaccessible : $RacineNas" }
foreach ($dir in @('Base','Base\locks','Actes','Patients','Config','Config\DDE','Modeles','Echange\Arrives',
                   'Echange\Arrives\EnCours','Echange\Arrives\Pris','Echange\Arrives\Annules',
                   'Echange\AEnvoyer','Echange\Traites','Sauvegardes','Logs')) {
    [void][IO.Directory]::CreateDirectory((Join-Path $RacineNas $dir))
}
$lockPath = Join-Path $RacineNas 'Base\locks\installation-nas.lock'
$lock = [IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
$excel = $null; $wb = $null
try {
    $assets = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'donnees_initiales.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($asset in $assets.PSObject.Properties) {
        $source = Join-Path (Join-Path $RacineSources 'DonneesInitiales') $asset.Name
        if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant() -ne $asset.Value) { throw "Ressource modifiee : $($asset.Name)" }
        $target = Join-Path $RacineNas $asset.Name
        if (-not (Test-Path -LiteralPath $target)) {
            [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
            [IO.File]::Copy($source,$target,$false)
        }
    }
    $config = Join-Path $RacineNas 'Config\config.ini'
    if (-not (Test-Path -LiteralPath $config)) {
        [IO.File]::Copy((Join-Path $RacineSources 'Src\ConfigDefaut\config.ini'),$config,$false)
    }
    # Les donnees metier sont initialisees/migrees par le service PostgreSQL.
    # Ne pas modifier Patients.xlsx ni les classeurs historiques pendant l'installation.
    Write-Host 'Ressources de support preparees. Bases historiques conservees ; migration PostgreSQL distincte.'
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    $lock.Dispose()
}
