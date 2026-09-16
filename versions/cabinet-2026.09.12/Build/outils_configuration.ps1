function Lire-ConfigurationIni([string]$Contenu) {
    $courante=''
    $sortieVue=$false
    $valeurs=@{}
    foreach($brute in ($Contenu -split "\r\n|\r|\n")){
        $ligne=$brute.Trim([char[]]" `t")
        if($ligne.Length -eq 0 -or $ligne.StartsWith(';') -or $ligne.StartsWith('#')){continue}
        if($ligne -match '^\[([^\]]*)\]'){
            $courante=$matches[1].Trim([char[]]" `t").ToLowerInvariant()
            if($courante -eq 'sortie'){
                if($sortieVue){throw 'Configuration ambigue : section [SORTIE] en double.'}
                $sortieVue=$true
            }
            continue
        }
        if($ligne -match '^([^=]*)=(.*)$'){
            $cle=$matches[1].Trim([char[]]" `t").ToLowerInvariant()
            $valeur=$matches[2].Trim([char[]]" `t")
            $index=$courante+'|'+$cle
            if($courante -eq 'sortie' -and $cle -in @('exportactif','dossier','nomfichier') -and $valeurs.ContainsKey($index)){
                throw ('Configuration ambigue : cle SORTIE/'+$cle+' en double.')
            }
            $valeurs[$index]=$valeur
        }
    }
    return $valeurs
}

function Lire-ValeurIni([string]$Contenu,[string]$Section,[string]$Cle) {
    $valeurs=Lire-ConfigurationIni $Contenu
    return $valeurs[$Section.Trim([char[]]" `t").ToLowerInvariant()+'|'+$Cle.Trim([char[]]" `t").ToLowerInvariant()]
}

function Verifier-ConfigurationSortie([string]$Chemin) {
    $contenu=[IO.File]::ReadAllText($Chemin)
    $valeurs=Lire-ConfigurationIni $contenu
    $actif=$valeurs['sortie|exportactif']
    $dossier=$valeurs['sortie|dossier']
    $nom=$valeurs['sortie|nomfichier']
    if($null -eq $actif){throw 'Configuration historique detectee : renseignez explicitement SORTIE/ExportActif=0 ou 1 avant de poursuivre.'}
    if($actif -notin @('0','1')){throw 'SORTIE/ExportActif doit valoir 0 ou 1.'}
    if([string]::IsNullOrWhiteSpace($dossier)){throw 'SORTIE/Dossier doit etre renseigne explicitement.'}
    if($nom -notin @('PublicationID','IdentitePublication')){throw 'SORTIE/NomFichier doit valoir PublicationID ou IdentitePublication.'}
}
