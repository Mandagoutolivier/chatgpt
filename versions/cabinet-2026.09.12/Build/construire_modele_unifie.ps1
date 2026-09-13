[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Prod6,
    [Parameter(Mandatory=$true)][string]$Cabinet1,
    [Parameter(Mandatory=$true)][string]$Sortie,
    [string]$RacineSources = ''
)
. (Join-Path $PSScriptRoot 'outils_construction.ps1')
if ([string]::IsNullOrWhiteSpace($RacineSources)) {
    $RacineSources = Split-Path $PSScriptRoot -Parent
}
$manifest = Lire-Manifeste $RacineSources
Verifier-ModeleSource $Prod6 'ModeleCourrierChatGPT_PROD(6).dotm' $manifest
Verifier-ModeleSource $Cabinet1 'Cabinet(1).dotm' $manifest
$Prod6 = (Resolve-Path -LiteralPath $Prod6).Path
$Cabinet1 = (Resolve-Path -LiteralPath $Cabinet1).Path
$Sortie = [IO.Path]::GetFullPath($Sortie)
if ($Sortie -in @($Prod6,$Cabinet1)) { throw 'La sortie doit etre distincte des modeles sources.' }
. (Join-Path $PSScriptRoot 'outils_assistant.ps1')
Attendre-FermetureOffice -Noms 'WINWORD'
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('CabinetBuild-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tmp)
$built = Join-Path $tmp 'CabinetUnifie.dotm'
[IO.File]::Copy($Prod6, $built)
$word = $null; $src = $null; $dst = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false; $word.DisplayAlerts = 0; $word.AutomationSecurity = 3
    $src = $word.Documents.Open($Cabinet1, $false, $true, $false)
    $dst = $word.Documents.Open($built, $false, $false, $false)
    # L'acces VBA doit etre autorise par la configuration Office de l'utilisateur.
    $srcProj = $src.VBProject; $dstProj = $dst.VBProject
    if ($srcProj.Protection -ne 0 -or $dstProj.Protection -ne 0) { throw 'Projet VBA protege.' }
    foreach ($component in $srcProj.VBComponents) {
        if ($component.Type -eq 100) { continue }
        $existing = @($dstProj.VBComponents | Where-Object { $_.Name -eq $component.Name })
        if ($existing.Count -gt 0) { throw "Collision de composants : $($component.Name)" }
        $ext = switch ($component.Type) { 1 {'.bas'} 2 {'.cls'} 3 {'.frm'} default {throw 'Type de composant inattendu.'} }
        $file = Join-Path $tmp ($component.Name + $ext)
        $component.Export($file)
        [void]$dstProj.VBComponents.Import($file)
    }
    Installer-SourcesVba $dstProj $manifest.word $RacineSources
    $dst.Save()
    $dst.Close(0); $dst = $null
    $src.Close(0); $src = $null
    $word.Quit(); $word = $null
    Installer-Ruban $built (Join-Path $PSScriptRoot 'ruban_unifie.xml')
    Publier-FichierConstruit $built $Sortie
    Write-Host "Modele construit : $Sortie. Compilation et essais Word requis avant utilisation clinique."
} finally {
    if ($null -ne $dst) { try { $dst.Close(0) } catch {} }
    if ($null -ne $src) { try { $src.Close(0) } catch {} }
    if ($null -ne $word) { try { $word.Quit() } catch {} }
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
