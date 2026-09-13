$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Build/outils_construction.ps1')
# Doublures des objets COM : verifier le traitement des feuilles sans Excel installe.
class ModuleEssai {
    [int]$CountOfLines=0
    [string]$Texte=''
    [int]$AppelsLecture=0
    [void] DeleteLines([int]$Start,[int]$Count) { $this.Texte='';$this.CountOfLines=0 }
    [void] AddFromString([string]$Code) { $this.Texte=$Code;$this.CountOfLines=@($Code -split '\r?\n').Count }
    [string] Lines([int]$Start,[int]$Count) {
        $this.AppelsLecture++
        if ($Count -eq 0) { throw 'Lecture interdite avec zero lignes.' }
        return $this.Texte
    }
}
class FeuillesEssai {
    [hashtable]$ParNom=@{}
    [object] Item([string]$Nom) {
        if (-not $this.ParNom.ContainsKey($Nom)) { throw 'Feuille absente.' }
        return $this.ParNom[$Nom]
    }
}
$count=0
function Verifier([bool]$Condition,[string]$Nom) { if (-not $Condition) { throw "ECHEC : $Nom" };$script:count++;Write-Host "PASS : $Nom" }
function Refuser([scriptblock]$Action,[string]$Nom) { $rejected=$false;try { & $Action } catch { $rejected=$true };Verifier $rejected $Nom }
function Composant([string]$Nom,[int]$Type=100) { return [pscustomobject]@{Name=$Nom;Type=$Type;CodeModule=[ModuleEssai]::new()} }
function Classeur([string]$CodeAgenda) {
    $components=New-Object Collections.ArrayList
    foreach ($name in @('Feuil1',$CodeAgenda,'ThisWorkbook')) { [void]$components.Add((Composant $name)) }
    $sheets=[FeuillesEssai]::new()
    $sheets.ParNom=@{Accueil=[pscustomobject]@{CodeName='Feuil1'};Agenda=[pscustomobject]@{CodeName=$CodeAgenda}}
    return [pscustomobject]@{Worksheets=$sheets;VBProject=[pscustomobject]@{Protection=0;VBComponents=$components;References=@()}}
}
$manifest=Lire-Manifeste $root
$documents=@($manifest.excel | Where-Object { $_.kind -eq 'document' })
# Le modele reel contient deux feuilles : chacune doit avoir une declaration du manifeste.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip=[IO.Compression.ZipFile]::OpenRead((Join-Path $root 'ModelesSource/Cabinet.xlsm'))
try {
    [xml]$xml=Lire-EntreeZip $zip 'xl/workbook.xml'
    $ns=New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('x','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $sheets=@($xml.SelectNodes('/x:workbook/x:sheets/x:sheet',$ns))
    foreach ($sheet in $sheets) {
        $mapped=@($documents | Where-Object { $null -ne $_.PSObject.Properties['sheet'] -and $_.sheet -eq $sheet.GetAttribute('name') })
        Verifier ($mapped.Count -eq 1) ('feuille du modele declaree : '+$sheet.GetAttribute('name'))
    }
} finally { $zip.Dispose() }
foreach ($automatic in @('Feuil2','Sheet2')) {
    $wb=Classeur $automatic
    $agenda=$wb.VBProject.VBComponents[1]
    Lier-FeuillesVba $wb $documents
    Verifier ($agenda.Name -eq 'FeuilAgenda' -and $wb.Worksheets.ParNom.Count -eq 2) "Agenda lie sans supprimer de feuille : $automatic"
    Installer-SourcesVba $wb.VBProject $documents $root
    Verifier ($wb.VBProject.VBComponents.Count -eq 3 -and $agenda.CodeModule.Texte -match 'Option Explicit') "construction accepte les documents attendus : $automatic"
    Verifier ($wb.VBProject.VBComponents[2].CodeModule.Texte -match 'Workbook_SheetBeforeDoubleClick') 'evenement agenda du classeur conserve'
}
$wb=Classeur 'Feuil2';$wb.Worksheets.ParNom.Remove('Agenda')
Refuser { Lier-FeuillesVba $wb $documents } 'feuille Agenda absente refusee'
$wb=Classeur 'Feuil2';[void]$wb.VBProject.VBComponents.Add((Composant 'FeuilAgenda' 1))
Refuser { Lier-FeuillesVba $wb $documents } 'collision de nom VBA refusee'
$wb=Classeur 'Feuil2';Lier-FeuillesVba $wb $documents
[void]$wb.VBProject.VBComponents.Add((Composant 'FeuilleInconnue'))
Refuser { Installer-SourcesVba $wb.VBProject $documents $root } 'autre module document inconnu reste refuse'
$empty=[ModuleEssai]::new()
Verifier ((Lire-ModuleVba $empty) -eq '' -and $empty.AppelsLecture -eq 0) 'validation du module vide sans lecture Lines 1 0'
$empty.AddFromString('Option Explicit')
Verifier ((Lire-ModuleVba $empty) -eq 'Option Explicit' -and $empty.AppelsLecture -eq 1) 'validation du module non vide lit le code'
Write-Host "$count controles Excel reussis. Objets COM simules ; aucune compilation Office executee."
