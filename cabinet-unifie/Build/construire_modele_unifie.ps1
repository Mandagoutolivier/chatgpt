param(
    [Parameter(Mandatory=$true)][string]$Prod6,
    [Parameter(Mandatory=$true)][string]$Cabinet1,
    [Parameter(Mandatory=$true)][string]$Sortie,
    [string]$RacineSources = (Split-Path $PSScriptRoot -Parent)
)
$ErrorActionPreference = 'Stop'

foreach ($p in @($Prod6, $Cabinet1)) {
    if (-not (Test-Path $p)) { throw "Fichier introuvable : $p" }
}
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Fermez completement Word avant de construire le modele.'
}

$Prod6 = (Resolve-Path $Prod6).Path
$Cabinet1 = (Resolve-Path $Cabinet1).Path
$Sortie = [IO.Path]::GetFullPath($Sortie)
New-Item -ItemType Directory -Force -Path (Split-Path $Sortie -Parent) | Out-Null
Copy-Item $Prod6 $Sortie -Force
$tmp = Join-Path $env:TEMP ('CabinetUnifie_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

function Copier-EntreeZip([string]$sourceZip, [string]$cibleZip, [string]$nomEntree) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $srcZip = [IO.Compression.ZipFile]::OpenRead($sourceZip)
    try {
        $entree = $srcZip.GetEntry($nomEntree)
        if ($null -eq $entree) { throw "Entree OOXML absente : $nomEntree" }
        $memoire = New-Object IO.MemoryStream
        $flux = $entree.Open()
        try { $flux.CopyTo($memoire) } finally { $flux.Dispose() }
        $octets = $memoire.ToArray(); $memoire.Dispose()
    } finally { $srcZip.Dispose() }
    $dstZip = [IO.Compression.ZipFile]::Open($cibleZip, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $ancienne = $dstZip.GetEntry($nomEntree)
        if ($null -ne $ancienne) { $ancienne.Delete() }
        $nouvelle = $dstZip.CreateEntry($nomEntree)
        $flux = $nouvelle.Open()
        try { $flux.Write($octets, 0, $octets.Length) } finally { $flux.Dispose() }
    } finally { $dstZip.Dispose() }
}

$word = $null; $src = $null; $dst = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $src = $word.Documents.Open($Cabinet1, $false, $true, $false)
    $dst = $word.Documents.Open($Sortie, $false, $false, $false)
    $srcProj = $src.VBProject
    $dstProj = $dst.VBProject
    if ($null -eq $srcProj -or $null -eq $dstProj) {
        throw "Activez dans Word : Fichier > Options > Centre de gestion de la confidentialite > Parametres > Parametres des macros > Acces approuve au modele d'objet du projet VBA."
    }

    # Cabinet(1) est ajoute au moteur PROD(6). ThisDocument reste celui du modele cible.
    $exports = @()
    foreach ($c in @($srcProj.VBComponents)) {
        if ($c.Type -eq 100) { continue }
        $ext = switch ($c.Type) { 1 {'.bas'} 2 {'.cls'} 3 {'.frm'} default {'.bas'} }
        $f = Join-Path $tmp ($c.Name + $ext)
        $c.Export($f)
        $exports += $f
    }
    foreach ($f in $exports) { [void]$dstProj.VBComponents.Import($f) }

    foreach ($f in @(
        (Join-Path $RacineSources 'Src\Integration\modPowerMicUnifie.bas'),
        (Join-Path $RacineSources 'Src\Integration\modIntegrationUnifie.bas'),
        (Join-Path $RacineSources 'Src\Integration\modAttenteLocale.bas')
    )) {
        if (-not (Test-Path $f)) { throw "Module d'integration absent : $f" }
        [void]$dstProj.VBComponents.Import($f)
    }

    # Un seul moteur de gras : celui de Cabinet(1), avec ses dictionnaires.
    $cm = $dstProj.VBComponents.Item('modProdRapide').CodeModule
    $ligne = 1
    while ($ligne -le $cm.CountOfLines -and $cm.Lines($ligne,1) -notmatch 'If Not NGL_NormaliserDocument\(docPrincipal\) Then') { $ligne++ }
    if ($ligne -gt $cm.CountOfLines) { throw 'Bloc de gras PROD(6) introuvable.' }
    $cm.DeleteLines($ligne, 3)
    $cm.InsertLines($ligne, '    Call modGras.AppliquerGras(docPrincipal)')

    # Publication du courrier final dans la file du secretariat avant fermeture.
    $ligne = 1
    while ($ligne -le $cm.CountOfLines -and $cm.Lines($ligne,1) -notmatch 'If SD_EnregistrerCourrierFinal\(') { $ligne++ }
    if ($ligne -gt $cm.CountOfLines) { throw 'Bloc SaveAs2 PROD(6) introuvable.' }
    while ($ligne -le $cm.CountOfLines -and $cm.Lines($ligne,1) -notmatch 'cheminSortieDragon\) Then') { $ligne++ }
    if ($ligne -gt $cm.CountOfLines) { throw 'Fin du bloc SaveAs2 PROD(6) introuvable.' }
    $cm.InsertLines($ligne + 1, "`r`n            modIntegrationUnifie.TransmettreSecretariat _`r`n                docPrincipal, cheminSortieDragon")

    $dst.Save()
    $dst.Close(0)
    $dst = $null
    $src.Close(0)
    $src = $null
    $word.Quit()
    $word = $null
    # Le ruban Cabinet(1) fait partie du modele unique. Les relations et types
    # OOXML correspondants sont recopies apres fermeture de Word.
    Copier-EntreeZip $Cabinet1 $Sortie '[Content_Types].xml'
    Copier-EntreeZip $Cabinet1 $Sortie '_rels/.rels'
    Copier-EntreeZip $Cabinet1 $Sortie 'customUI/customUI14.xml'
    Write-Host "Modele unifie construit : $Sortie" -ForegroundColor Green
    Write-Host 'Macros PowerMic : Unifie_A_NouvelleLettre / Unifie_B_FormuleAppel / Unifie_C_InsererPatient / Unifie_D_Finaliser'
}
finally {
    if ($null -ne $dst) { try { $dst.Close(0) } catch {} }
    if ($null -ne $src) { try { $src.Close(0) } catch {} }
    if ($null -ne $word) { try { $word.Quit() } catch {} }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
