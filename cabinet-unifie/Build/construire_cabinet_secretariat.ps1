[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$CabinetXlsm,
    [Parameter(Mandatory=$true)][string]$Sortie,
    [string]$RacineSources = (Split-Path $PSScriptRoot -Parent)
)
. (Join-Path $PSScriptRoot 'outils_construction.ps1')
$manifest = Lire-Manifeste $RacineSources
Verifier-ModeleSource $CabinetXlsm 'Cabinet.xlsm' $manifest
$CabinetXlsm = (Resolve-Path -LiteralPath $CabinetXlsm).Path
$Sortie = [IO.Path]::GetFullPath($Sortie)
if ($Sortie -eq $CabinetXlsm) { throw 'La sortie doit etre distincte du classeur source.' }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Fermez completement Excel.' }
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('CabinetExcel-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tmp)
$built = Join-Path $tmp 'Cabinet.xlsm'
[IO.File]::Copy($CabinetXlsm, $built)
$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3; $excel.EnableEvents = $false
    $wb = $excel.Workbooks.Open($built, 0, $false)
    Installer-SourcesVba $wb.VBProject $manifest.excel $RacineSources
    $wb.Save()
    $wb.Close($false); $wb = $null
    $excel.Quit(); $excel = $null
    Publier-FichierConstruit $built $Sortie
    Write-Host "Classeur construit : $Sortie. Compilation et essais Excel requis."
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
