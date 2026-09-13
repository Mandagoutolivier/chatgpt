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
. (Join-Path $PSScriptRoot 'outils_assistant.ps1')
Attendre-FermetureOffice -Noms 'EXCEL'
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
    Lier-FeuillesVba $wb $manifest.excel
    Installer-SourcesVba $wb.VBProject $manifest.excel $RacineSources
    $ws=$wb.Worksheets.Item(1)
    $top=80
    foreach ($shape in $ws.Shapes) { $top=[Math]::Max($top,[double]$shape.Top+[double]$shape.Height+20) }
    $n=0
    foreach ($item in @(@('Assure / ayant droit','UI_AssurePatient'),@('Parametres des destinataires','UI_ParametresDestinataire'),@('Encaisser une seance','UI_EncaisserSeance'))) {
        $button=$ws.Buttons().Add(20,$top+45*$n,210,32)
        $button.Caption=$item[0];$button.OnAction=$item[1]
        $n++
    }
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
