$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'Build/outils_assistant.ps1')
$count=0
function Verifier([bool]$Condition,[string]$Nom) { if (-not $Condition) { throw "ECHEC : $Nom" };$script:count++;Write-Host "PASS : $Nom" }
function Refuser([scriptblock]$Action,[string]$Nom) { $failed=$false;try { & $Action } catch { $failed=$true };Verifier $failed $Nom }
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('choix-nas-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tmp)
try {
    Verifier ((Normaliser-CheminNasAssistant ' "\\ds224\home\Donnees Cabinet\" ') -eq '\\ds224\home\Donnees Cabinet') 'chemin UNC avec espaces et guillemets conserve'
    foreach ($path in @('C:\CabinetCardio','Z:\claude','\\ds224','\\ds224\home\..\secret','\\ds224\home\x:flux','\\ds224\home\x.')) {
        Refuser { Normaliser-CheminNasAssistant $path } "chemin invalide refuse : $path"
    }
    Tester-EcritureNasAssistant $tmp
    Verifier (@(Get-ChildItem -LiteralPath $tmp -Force).Count -eq 0) 'sonde ecriture supprimee sans residu'
    $file=Join-Path $tmp 'existant.txt';[IO.File]::WriteAllText($file,'ne pas modifier')
    Tester-EcritureNasAssistant $tmp
    Verifier ([IO.File]::ReadAllText($file) -eq 'ne pas modifier') 'fichiers existants conserves'
    [IO.File]::WriteAllText((Join-Path $tmp 'Installer.ps1'),'source fictive')
    Verifier (Tester-DossierCodeAssistant $tmp) 'dossier du programme identifie'
    # Simulation des reponses utilisateur et du reseau : aucune connexion SMB reelle.
    & {
        $ctx=[pscustomobject]@{reponses=$null;existants=@();creations=0;sondes=0;code=$false;refuserEcriture=$false}
        function Read-Host {
            param($Prompt)
            if ($ctx.reponses.Count -eq 0) { throw 'Question supplementaire inattendue dans le parcours.' }
            return $ctx.reponses.Dequeue()
        }
        function Test-Path { param($LiteralPath,$PathType);return ($LiteralPath -in $ctx.existants) }
        function Tester-DossierCodeAssistant { param($Dossier);return $ctx.code }
        function Creer-DossierNasAssistant { param($Dossier);$ctx.creations++;$ctx.existants+=@($Dossier) }
        function Tester-EcritureNasAssistant { param($Dossier);$ctx.sondes++;if ($ctx.refuserEcriture) { throw 'acces refuse simule' } }
        function Reponses([string[]]$Valeurs) {
            $ctx.reponses=New-Object 'System.Collections.Generic.Queue[string]'
            foreach ($value in $Valeurs) { $ctx.reponses.Enqueue($value) }
        }
        Reponses @('','Q')
        $chosen=Choisir-RacineNasAssistant '' '' 'absent.txt'
        Verifier ($null -eq $chosen -and $ctx.sondes -eq 0 -and $ctx.creations -eq 0) 'aucun partage impose au premier lancement'
        Reponses @('','Q')
        $chosen=Choisir-RacineNasAssistant '' '\\ds224\CabinetCardio' 'absent.txt'
        Verifier ($null -eq $chosen -and $ctx.creations -eq 0) 'ancien partage absent ne cree pas un partage SMB'
        $ctx.existants=@('\\nas\commun\donnees')
        Reponses @('')
        $chosen=Choisir-RacineNasAssistant '' '\\nas\commun\donnees' 'absent.txt'
        Verifier ($chosen -eq '\\nas\commun\donnees' -and $ctx.sondes -eq 1) 'reprise du choix memorise avec test ecriture'
        $ctx.existants=@('\\nas\commun');$ctx.creations=0
        Reponses @('\\nas\commun\neuf','','Q')
        $chosen=Choisir-RacineNasAssistant '' '' 'absent.txt'
        Verifier ($null -eq $chosen -and $ctx.creations -eq 0) 'pas de creation sans selection explicite'
        Reponses @('\\nas\commun\neuf','CREER')
        $chosen=Choisir-RacineNasAssistant '' '' 'absent.txt'
        Verifier ($chosen -eq '\\nas\commun\neuf' -and $ctx.creations -eq 1) 'sous dossier neuf cree sur demande'
        $ctx.code=$true
        Reponses @('','Q')
        $chosen=Choisir-RacineNasAssistant '\\nas\commun\neuf' '' 'absent.txt'
        Verifier ($null -eq $chosen) 'dossier contenant le code refuse comme base'
        $ctx.code=$false;$ctx.existants=@('\\nas\home\donnees')
        Reponses @('','','Q')
        $chosen=Choisir-RacineNasAssistant '\\nas\home\donnees' '' 'absent.txt'
        Verifier ($null -eq $chosen) 'home exige la verification de l acces commun'
        Reponses @('','COMMUN')
        $chosen=Choisir-RacineNasAssistant '\\nas\home\donnees' '' 'absent.txt'
        Verifier ($chosen -eq '\\nas\home\donnees') 'home accepte apres verification explicite'
        $ctx.existants=@('\\nas\commun\donnees');$ctx.refuserEcriture=$true
        Reponses @('','Q')
        $chosen=Choisir-RacineNasAssistant '\\nas\commun\donnees' '' 'absent.txt'
        Verifier ($null -eq $chosen) 'dossier sans droits ecriture non retenu'
    }
    Write-Host "$count controles du choix NAS reussis. Reseau simule ; aucun acces au NAS du cabinet."
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force }
