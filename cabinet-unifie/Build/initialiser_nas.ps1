[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$RacineNas,
    [string]$RacineSources = (Split-Path $PSScriptRoot -Parent)
)
. (Join-Path $PSScriptRoot 'outils_construction.ps1')
throw 'INSTALLATEUR ARCHIVE ET DESACTIVE : utilisez uniquement la version service NAS apres publication et recette du nouveau lanceur.'
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
    # Creation ou ajout de colonnes, jamais remplacement des lignes existantes.
    $patientPath = Join-Path $RacineNas 'Base\Patients.xlsx'
    $patientLock = [IO.File]::Open((Join-Path $RacineNas 'Base\locks\Patients.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
        $existing = Test-Path -LiteralPath $patientPath
        if ($existing) {
            $backup = Join-Path $RacineNas ('Sauvegardes\Patients-avant-schema-' + [guid]::NewGuid().ToString('N') + '.xlsx')
            [IO.File]::Copy($patientPath,$backup,$false)
            $wb = $excel.Workbooks.Open($patientPath,0,$false)
            if ($wb.ReadOnly) { throw 'Patients.xlsx est en lecture seule. Fermez-le sur les autres postes.' }
        } else { $wb = $excel.Workbooks.Add(-4167) }
        $schemas = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'schemas.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $changed = -not $existing
        foreach ($schema in $schemas.PSObject.Properties) {
            $ws = $null
            foreach ($sheet in $wb.Worksheets) { if ($sheet.Name -eq $schema.Name) { $ws = $sheet; break } }
            if ($null -eq $ws) {
                if (-not $existing -and $wb.Worksheets.Count -eq 1 -and $wb.Worksheets.Item(1).Name -ne 'PATIENTS') { $ws=$wb.Worksheets.Item(1) }
                else { $ws=$wb.Worksheets.Add() }
                $ws.Name=$schema.Name; $changed=$true
            }
            $headers=@{}; $last=$ws.Cells.Item(1,$ws.Columns.Count).End(-4159).Column
            for ($c=1; $c -le $last; $c++) {
                $value=[string]$ws.Cells.Item(1,$c).Value2
                if (-not [string]::IsNullOrWhiteSpace($value)) { $headers[$value]=$c }
            }
            if ($headers.Count -eq 0) { $last=0 }
            foreach ($header in $schema.Value) {
                if (-not $headers.ContainsKey($header)) { $last++;$ws.Cells.Item(1,$last).Value2=$header;$changed=$true }
            }
            $ws.Rows.Item(1).Font.Bold=$true
        }
        if ($changed) {
            if ($existing) { $wb.Save() }
            else {
                $temp = Join-Path $RacineNas ('Base\Patients-' + [guid]::NewGuid().ToString('N') + '.tmp.xlsx')
                $wb.SaveAs($temp,51)
                $wb.Close($false);$wb=$null
                [IO.File]::Move($temp,$patientPath)
            }
        }
        if ($null -ne $wb) { $wb.Close($false);$wb=$null }
        $excel.Quit();$excel=$null
    } finally {
        # Garder le verrou tant que le classeur peut encore etre ouvert.
        try {
            if ($null -ne $wb) { try { $wb.Close($false) } finally { $wb=$null } }
        } finally {
            try {
                if ($null -ne $excel) { try { $excel.Quit() } finally { $excel=$null } }
            } finally { $patientLock.Dispose() }
        }
    }
    Write-Host 'Ressources NAS et schema prepares. Bases, dictionnaires et configuration existants conserves.'
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    $lock.Dispose()
}
