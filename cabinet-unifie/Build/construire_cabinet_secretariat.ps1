param(
    [Parameter(Mandatory=$true)][string]$CabinetXlsm,
    [Parameter(Mandatory=$true)][string]$Sortie,
    [string]$RacineSources = (Split-Path $PSScriptRoot -Parent)
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $CabinetXlsm)) { throw "Fichier introuvable : $CabinetXlsm" }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Fermez completement Excel.' }
$CabinetXlsm = (Resolve-Path $CabinetXlsm).Path
$Sortie = [IO.Path]::GetFullPath($Sortie)
New-Item -ItemType Directory -Force -Path (Split-Path $Sortie -Parent) | Out-Null
Copy-Item $CabinetXlsm $Sortie -Force
$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 3
    $wb = $excel.Workbooks.Open($Sortie)
    $projet = $wb.VBProject
    if ($null -eq $projet) { throw "Activez l'acces approuve au modele d'objet VBA dans Excel." }
    $sources = @{
        'modGdt' = 'Src\Commun\modGdt.bas'
        'modEchange' = 'Src\Excel\modEchange.bas'
        'modUI' = 'Src\Excel\modUI.bas'
    }
    foreach ($nom in $sources.Keys) {
        try { $projet.VBComponents.Remove($projet.VBComponents.Item($nom)) } catch {}
        $f = Join-Path $RacineSources $sources[$nom]
        if (-not (Test-Path $f)) { throw "Source absente : $f" }
        [void]$projet.VBComponents.Import($f)
    }
    $wb.Save()
    $wb.Close($false); $wb = $null
    Write-Host "Application secretariat construite : $Sortie" -ForegroundColor Green
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
}
