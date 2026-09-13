[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$DossierPrepare,
    [switch]$CompilationWordValidee,
    [switch]$CompilationExcelValidee,
    [switch]$RecetteValidee
)
. (Join-Path $PSScriptRoot 'outils_construction.ps1')
. (Join-Path $PSScriptRoot 'outils_installation.ps1')
if ($env:OS -ne 'Windows_NT') { throw 'Validation Office requise sur Windows.' }
. (Join-Path $PSScriptRoot 'outils_assistant.ps1')
Attendre-FermetureOffice
$root=Split-Path $PSScriptRoot -Parent
$stage=(Resolve-Path -LiteralPath $DossierPrepare).Path
$receipt=Get-Content -LiteralPath (Join-Path $stage 'preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$current=Empreintes-Sources $root
if ($current.Count -ne @($receipt.sources.PSObject.Properties).Count) { throw 'Sources modifiees : refaire la preparation.' }
foreach ($p in $receipt.sources.PSObject.Properties) { if ($current[$p.Name] -ne $p.Value) { throw 'Sources modifiees : refaire la preparation.' } }
$medecin=$receipt.profil -in @('Domicile','CabinetMedecin')
$secretariat=$receipt.profil -in @('Domicile','CabinetSecretariat')
if (-not $RecetteValidee -or ($medecin -and -not $CompilationWordValidee) -or ($secretariat -and -not $CompilationExcelValidee)) {
    throw 'Effectuez la compilation VBA et RECETTE_WINDOWS.md, puis renseignez les commutateurs correspondant aux controles vraiment realises.'
}
$manifest=Lire-Manifeste $root
foreach ($hostName in @('word','excel')) {
    if (($hostName -eq 'word' -and -not $medecin) -or ($hostName -eq 'excel' -and -not $secretariat)) { continue }
    $app=$null;$document=$null
    try {
        if ($hostName -eq 'word') {
            $app=New-Object -ComObject Word.Application;$app.AutomationSecurity=3
            $document=$app.Documents.Open((Join-Path $stage 'CabinetUnifie.dotm'),$false,$true)
        } else {
            $app=New-Object -ComObject Excel.Application;$app.AutomationSecurity=3;$app.EnableEvents=$false
            $document=$app.Workbooks.Open((Join-Path $stage 'Cabinet.xlsm'),0,$true)
        }
        foreach ($item in $manifest.$hostName) {
            $component=$document.VBProject.VBComponents.Item($item.name)
            $code=Lire-ModuleVba $component.CodeModule
            $expected=Lire-CodeVba (Join-Path $root $item.path)
            if ($code.Replace("`r`n","`n").Trim() -ne $expected.Replace("`r`n","`n").Trim()) { throw "Source Office differente : $($item.name)" }
        }
        foreach ($reference in $document.VBProject.References) { if ($reference.IsBroken) { throw "Reference Office manquante : $($reference.Name)" } }
    } finally {
        if ($null -ne $document) { $document.Close($false) }
        if ($null -ne $app) { $app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
    }
}
# La compilation peut modifier le p-code : figer les nouveaux binaires dont les sources viennent d etre comparees.
Ecrire-Preparation $stage $receipt.profil $root
$receipt=Get-Content -LiteralPath (Join-Path $stage 'preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$receipt.compilationOffice='Confirmee par operateur ; sources relues dans Office'
$receipt | Add-Member -NotePropertyName recetteValidee -NotePropertyValue $true
$receipt | Add-Member -NotePropertyName dateValidation -NotePropertyValue ([DateTime]::UtcNow.ToString('o'))
$receipt | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $stage 'preparation.json') -Encoding UTF8
Write-Host 'Preparation validee et empreintes figees. Vous pouvez activer ce dossier avec Installer.ps1 -Mode Installation.'
