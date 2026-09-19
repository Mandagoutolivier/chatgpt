# Tests de l instrumentation des copies ; aucun COM et aucun Office lance.
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Build/outils_recette_isolee.ps1')
. (Join-Path $root 'Build/outils_recette_u1.ps1')
class ModuleRecetteFictif {
    [string]$Texte
    [int]$CountOfLines
    ModuleRecetteFictif([string]$code) { $this.Texte=$code; $this.CountOfLines=($code -split "`n").Count }
    [string] Lines([int]$premiere,[int]$nombre) { return $this.Texte }
    [void] DeleteLines([int]$premiere,[int]$nombre) { $this.Texte='';$this.CountOfLines=0 }
    [void] AddFromString([string]$code) { $this.Texte=$code;$this.CountOfLines=($code -split "`n").Count }
}
$n=0
function Verifier([bool]$ok,[string]$nom) { if (-not $ok) { throw ('ECHEC : '+$nom) };$script:n++;Write-Host ('PASS : '+$nom) }
function Refuser([scriptblock]$action,[string]$motif) { $message='';try { & $action|Out-Null }catch{$message=$_.Exception.Message};Verifier ($message -like $motif) $motif }
$temp=Join-Path ([IO.Path]::GetTempPath()) ('Cabinet-Qualification-Test-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
try {
    $manifest=Get-Content -LiteralPath (Join-Path $root 'Build/manifest.json') -Raw|ConvertFrom-Json
    foreach ($hote in @('Word','Excel')) {
        $entries=@($manifest.($hote.ToLowerInvariant()))+@($manifest.($hote.ToLowerInvariant()+'_recette'))
        $components=@();$hashes=@{}
        foreach ($entry in $entries) {
            $path=Join-Path $root $entry.path
            $hashes[$path]=(Get-FileHash $path -Algorithm SHA256).Hash
            $code=([IO.File]::ReadAllText($path) -split '\r?\n'|Where-Object {$_ -notmatch '^Attribute\s'}) -join "`r`n"
            $components+=@([pscustomobject]@{Name=$entry.name;CodeModule=[ModuleRecetteFictif]::new($code)})
        }
        $sortie=Join-Path $temp $hote
        [void][IO.Directory]::CreateDirectory($sortie)
        Preparer-CopieRecetteIsolee ([pscustomobject]@{VBComponents=$components}) $sortie $hote
        $byName=@{};foreach($c in $components){$byName[$c.Name]=$c.CodeModule.Texte}
        Verifier ($byName.modServiceNas -notmatch 'WinHttp|service\.token|service\.url' -and $byName.modServiceNas -match 'Acces externe interdit') ($hote+' RPC bloque avant lecture du jeton')
        Verifier (($components|Where-Object {$_.CodeModule.Texte -match 'Environ\$?\("(APPDATA|LOCALAPPDATA|TEMP|TMP|OPENAI_API_KEY)"\)'}).Count -eq 0) ($hote+' aucun chemin de poste ou cle environnement reel')
        Verifier ($byName.modConfig -match 'DonneesFictives' -and $byName.modConfig -match 'poste.ini') ($hote+' configuration dirigee vers fixture')
        if ($hote -eq 'Word') {
            Verifier ($byName.modOpenAI_v22_corrige -notmatch 'WinHttp' -and $byName.modApiConfiguration -notmatch 'Environ\$\("OPENAI_API_KEY"\)') 'OpenAI sans transport ni lecture de cle'
            Verifier ($byName.modControleCourrier -match 'MSXML2.DOMDocument.6.0') 'DOM XML local preserve pour empreinte U1'
            Verifier ($byName.modRaccourcis -match 'Evenement neutralise') 'AutoExec neutralise'
        } else {
            Verifier ($byName.modCerfaPrint -notmatch '(?im)^\s*\w+\.PrintOut\b') 'Impression physique retiree de la copie'
            Verifier ($byName.ThisWorkbook -notmatch 'DemarrerScrutation|ArreterScrutation') 'Evenements du classeur sans scrutation'
        }
        foreach ($path in $hashes.Keys) { if ((Get-FileHash $path -Algorithm SHA256).Hash -ne $hashes[$path]) { throw 'Source originale modifiee.' } }
        Verifier $true ($hote+' sources originales inchangees')
    }
    Refuser { Remplacer-ProcedureRecette 'Option Explicit' 'Appeler' 'Exit Function' } '*non unique*'
    Refuser { Verifier-ResultatRecetteSimple '{"reussis":0,"echec":false}' 'Suite' 1 } '*incomplete*'
    Refuser { Verifier-ResultatRecetteSimple '{"reussis":67,"echec":"false"}' 'Suite' 67 } '*incomplet*'
    Refuser { Verifier-ResultatRecetteSimple '{"reussis":66,"echec":false}' 'Suite' 67 } '*incomplete*'
    $r=Verifier-ResultatRecetteSimple '{"reussis":67,"echec":false}' 'Suite' 67
    Verifier ($r.reussis -eq 67) 'Compte rendu attendu accepte'
    Write-Host "$n controles d instrumentation reussis ; aucune compilation VBA reelle revendiquee."
} finally { Remove-Item -LiteralPath $temp -Recurse -Force }
