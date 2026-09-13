$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'Build/outils_assistant.ps1')
$count=0
function Verifier([bool]$Condition,[string]$Nom) { if (-not $Condition) { throw "ECHEC $Nom" };$script:count++;Write-Host "PASS : $Nom" }
# Scenarios de processus simules : aucune application Office ni aucun arret force.
& {
    $ctx=[pscustomobject]@{sondes=0;demandes=0;sommeils=0;disparition=0;reponse='';noms=@();messages=@()}
    function Get-Process {
        param($Name,$ErrorAction)
        $ctx.sondes++;$ctx.noms=@($Name)
        if ($ctx.sondes -le $ctx.disparition) { [pscustomobject]@{ProcessName='WINWORD';Id=12345;SessionId=2} }
    }
    function Start-Sleep { param($Milliseconds);$ctx.sommeils++ }
    function Read-Host { param($Prompt);$ctx.demandes++;return $ctx.reponse }
    function Write-Host { param($Object);$ctx.messages+=@([string]$Object) }
    function Stop-Process { throw 'INTERDIT : aucun processus ne doit etre termine.' }
    function Reinitialiser([int]$Disparition,[string]$Reponse='') {
        $ctx.sondes=0;$ctx.demandes=0;$ctx.sommeils=0;$ctx.disparition=$Disparition;$ctx.reponse=$Reponse;$ctx.messages=@()
    }
    Reinitialiser 0
    Attendre-FermetureOffice
    Verifier ($ctx.demandes -eq 0 -and $ctx.sommeils -eq 0) 'aucune attente quand Office est ferme'
    Reinitialiser 2
    Attendre-FermetureOffice
    Verifier ($ctx.demandes -eq 0 -and $ctx.sommeils -eq 2) 'fermeture asynchrone absorbee sans interrompre installation'
    Reinitialiser 1
    Attendre-FermetureOffice -DelaiSecondes 0
    Verifier ($ctx.demandes -eq 1) 'processus persistant attend intervention puis reprend'
    Verifier (($ctx.messages -join ' ') -match 'WINWORD.*PID 12345.*session Windows 2') 'processus identifie sans titre de document'
    Reinitialiser 3
    Attendre-FermetureOffice -DelaiSecondes 0
    Verifier ($ctx.demandes -eq 3) 'fermeture reellement controlee apres chaque Entree'
    Reinitialiser 999 'Q'
    $paused=$false
    try { Attendre-FermetureOffice -DelaiSecondes 0 } catch { $paused=$_.Exception.Message -like 'Installation en pause*' }
    Verifier ($paused -and $ctx.demandes -eq 1) 'pause volontaire sans fermeture forcee'
    Reinitialiser 999 'Q'
    try { Attendre-FermetureOffice -DelaiSecondes 2 } catch {}
    Verifier ($ctx.sommeils -eq 8 -and $ctx.demandes -eq 1) 'delai automatique borne avant question'
    Reinitialiser 0
    Attendre-FermetureOffice -Noms 'EXCEL'
    Verifier ($ctx.noms.Count -eq 1 -and $ctx.noms[0] -eq 'EXCEL') 'constructeur peut limiter le controle a son application'
}
Write-Host "$count controles de l attente Office reussis. Processus simules."
