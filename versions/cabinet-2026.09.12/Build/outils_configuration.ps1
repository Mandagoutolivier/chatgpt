function Lire-ValeurIni([string]$Contenu,[string]$Section,[string]$Cle) {
    $courante=''
    foreach($brute in ($Contenu -split "\r?\n")){
        $ligne=$brute.Trim()
        if($ligne -match '^\[([^\]]+)\]$'){$courante=$matches[1];continue}
        if($courante -ieq $Section -and $ligne -match '^([^=]+)=(.*)$' -and $matches[1].Trim() -ieq $Cle){return $matches[2].Trim()}
    }
    return $null
}

function Verifier-ConfigurationSortie([string]$Chemin) {
    $contenu=[IO.File]::ReadAllText($Chemin)
    $actif=Lire-ValeurIni $contenu 'SORTIE' 'ExportActif'
    $dossier=Lire-ValeurIni $contenu 'SORTIE' 'Dossier'
    $nom=Lire-ValeurIni $contenu 'SORTIE' 'NomFichier'
    if($null -eq $actif){throw 'Configuration historique detectee : renseignez explicitement SORTIE/ExportActif=0 ou 1 avant de poursuivre.'}
    if($actif -notin @('0','1')){throw 'SORTIE/ExportActif doit valoir 0 ou 1.'}
    if([string]::IsNullOrWhiteSpace($dossier)){throw 'SORTIE/Dossier doit etre renseigne explicitement.'}
    if($nom -notin @('PublicationID','IdentitePublication')){throw 'SORTIE/NomFichier doit valoir PublicationID ou IdentitePublication.'}
}

