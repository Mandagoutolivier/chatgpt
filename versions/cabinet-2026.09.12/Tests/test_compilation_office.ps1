$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Build/outils_assistant.ps1')
# Simulation de l interface Office : aucune application Office lancee, aucune compilation VBA pretendue.
class VoletCompilationEssai {
    [bool]$Echec=$false
    [int]$Affichages=0
    [void] Show() { $this.Affichages++;if ($this.Echec) { throw 'Affichage du volet indisponible.' } }
}
class ComposantsCompilationEssai {
    [hashtable]$Modules=@{}
    [string]$Demande=''
    [object] Item([string]$Nom) {
        $this.Demande=$Nom
        if (-not $this.Modules.ContainsKey($Nom)) { throw 'Module absent du fichier prepare.' }
        return $this.Modules[$Nom]
    }
}
class DocumentsCompilationEssai {
    [object]$Document
    [string]$Chemin=''
    [bool]$LectureSeule=$true
    [bool]$Echec=$false
    [object] Open([string]$Fichier,[object]$Options,[bool]$ReadOnly) {
        $this.Chemin=$Fichier;$this.LectureSeule=$ReadOnly
        if ($this.Echec) { throw 'Ouverture impossible.' }
        return $this.Document
    }
}
$count=0
function Verifier([bool]$Condition,[string]$Nom) { if (-not $Condition) { throw "ECHEC : $Nom" };$script:count++;Write-Host "PASS : $Nom" }
function Refuser([scriptblock]$Action,[string]$Message,[string]$Nom) {
    $errorText='';try { & $Action | Out-Null } catch { $errorText=$_.Exception.Message }
    Verifier ($errorText -like $Message) $Nom
}
function OfficeEssai([string]$Hote) {
    $pane=[VoletCompilationEssai]::new()
    $components=[ComposantsCompilationEssai]::new()
    $moduleName=if ($Hote -eq 'Word') { 'ThisDocument' } else { 'ThisWorkbook' }
    $components.Modules[$moduleName]=[pscustomobject]@{CodeModule=[pscustomobject]@{CodePane=$pane}}
    $project=[pscustomobject]@{Name='ProjetPrepare';VBComponents=$components}
    $doc=[pscustomobject]@{FullName=('C:\Preparation\'+$Hote);VBProject=$project;Enregistrements=0;Fermetures=0;EchecSauvegarde=$false}
    $doc | Add-Member ScriptMethod Save { if ($this.EchecSauvegarde) { throw 'Sauvegarde impossible.' };$this.Enregistrements++ }
    $doc | Add-Member ScriptMethod Close { param($SaveChanges);if ($SaveChanges) { throw 'Fermeture ne doit pas enregistrer sans confirmation.' };$this.Fermetures++ }
    $documents=[DocumentsCompilationEssai]::new();$documents.Document=$doc
    # Projet actif distinct et propriete non modifiable, comme dans le contrat documente de VBE.
    $vbe=[pscustomobject]@{MainWindow=[pscustomobject]@{Visible=$false}}
    $vbe | Add-Member ScriptProperty ActiveVBProject { 'AutreProjet' }
    $app=[pscustomobject]@{VBE=$vbe;Documents=$documents;Workbooks=$documents;AutomationSecurity=0;EnableEvents=$true;Visible=$false;Arrets=0}
    $app | Add-Member ScriptMethod Quit { $this.Arrets++ }
    return [pscustomobject]@{App=$app;Doc=$doc;Documents=$documents;Components=$components;Pane=$pane;Hote=$Hote}
}
function New-Object { param([string]$ComObject)
    if ($ComObject -ne ($script:office.Hote+'.Application')) { throw 'Mauvaise application Office.' }
    return $script:office.App
}
function Read-Host { param([string]$Prompt)
    $script:prompts++
    if ($script:reponses.Count -eq 0) { throw 'Confirmation inattendue.' }
    return $script:reponses.Dequeue()
}
function Preparer([string]$Hote,[string[]]$Reponses=@('OUI')) {
    $script:office=OfficeEssai $Hote
    $script:prompts=0
    $script:reponses=[Collections.Queue]::new()
    foreach ($r in $Reponses) { $script:reponses.Enqueue($r) }
}
try {
    foreach ($hostName in @('Word','Excel')) {
        Preparer $hostName @('invalide','oui')
        Refuser { $office.App.VBE.ActiveVBProject=$office.Doc.VBProject } '*ActiveVBProject*' 'doublure reproduit la propriete non modifiable'
        $result=@(Compiler-ProjetAssistant $office.Doc.FullName $hostName)
        Verifier ($result.Count -eq 1 -and $result[0] -eq $true -and $prompts -eq 2) "$hostName : confirmation explicite, reponse invalide redemandee"
        $expected=if ($hostName -eq 'Word') { 'ThisDocument' } else { 'ThisWorkbook' }
        Verifier ($office.Components.Demande -eq $expected -and $office.Pane.Affichages -eq 1 -and $office.App.VBE.MainWindow.Visible -and $office.App.VBE.ActiveVBProject -eq 'AutreProjet') "$hostName : ouverture du module du bon fichier sans modifier ActiveVBProject"
        Verifier ($office.Documents.Chemin -eq $office.Doc.FullName -and -not $office.Documents.LectureSeule -and $office.App.AutomationSecurity -eq 3 -and ($hostName -eq 'Word' -or -not $office.App.EnableEvents)) "$hostName : fichier exact, macros automatiques desactivees"
        Verifier ($office.Doc.Enregistrements -eq 1 -and $office.Doc.Fermetures -eq 1 -and $office.App.Arrets -eq 1) "$hostName : enregistrement apres OUI et fermeture"
    }
    Preparer 'Excel' @('NON')
    Refuser { Compiler-ProjetAssistant $office.Doc.FullName 'Excel' } '*Compilation Excel non validee*' 'NON bloque la validation'
    Verifier ($office.Doc.Enregistrements -eq 0 -and $office.Doc.Fermetures -eq 1 -and $office.App.Arrets -eq 1) 'NON ferme sans enregistrer'
    foreach ($answer in @('OUI','NON')) {
        Preparer 'Excel' @($answer);$office.Pane.Echec=$true
        if ($answer -eq 'OUI') {
            Verifier ((Compiler-ProjetAssistant $office.Doc.FullName 'Excel') -eq $true) 'affichage impossible permet la compilation manuelle confirmee'
        } else {
            Refuser { Compiler-ProjetAssistant $office.Doc.FullName 'Excel' } '*Compilation Excel non validee*' 'affichage impossible ne valide pas automatiquement'
        }
        Verifier ($prompts -eq 1 -and $office.Doc.Enregistrements -eq [int]($answer -eq 'OUI')) 'repli manuel exige toujours la confirmation'
    }
    Preparer 'Excel'
    $office.Doc | Add-Member ScriptProperty VBProject { throw 'Acces VBA interdit.' } -Force
    Refuser { Compiler-ProjetAssistant $office.Doc.FullName 'Excel' } '*Compilation Excel interrompue*' 'refus acces VBA reste bloquant avec contexte Excel'
    Verifier ($prompts -eq 0 -and $office.Doc.Enregistrements -eq 0 -and $office.App.Arrets -eq 1) 'refus acces ne devient pas un repli manuel'
    Preparer 'Excel';$office.Components.Modules.Clear()
    Refuser { Compiler-ProjetAssistant $office.Doc.FullName 'Excel' } '*Module absent*' 'module prepare absent reste bloquant'
    Verifier ($prompts -eq 0 -and $office.App.Arrets -eq 1) 'aucune confirmation sur un projet incorrect'
    Preparer 'Word';$office.Doc.EchecSauvegarde=$true
    Refuser { Compiler-ProjetAssistant $office.Doc.FullName 'Word' } '*Compilation Word interrompue*Sauvegarde impossible*' 'echec sauvegarde ne renvoie aucun succes'
    Verifier ($office.Doc.Fermetures -eq 1 -and $office.App.Arrets -eq 1) 'echec sauvegarde ferme Office'
    Preparer 'Excel';$office.Documents.Echec=$true
    Refuser { Compiler-ProjetAssistant $office.Doc.FullName 'Excel' } '*Ouverture impossible*' 'echec ouverture conserve le diagnostic'
    Verifier ($prompts -eq 0 -and $office.App.Arrets -eq 1) 'echec ouverture libere application sans confirmation'
    Write-Host "$count controles de compilation guidee reussis. Office simule ; compilation VBA reelle non testee."
} finally {
    Remove-Item Function:\New-Object
    Remove-Item Function:\Read-Host
}
