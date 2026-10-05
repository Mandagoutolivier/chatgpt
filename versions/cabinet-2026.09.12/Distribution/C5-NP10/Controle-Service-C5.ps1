# Fonctions sans action a l'import. Les jetons ne sont jamais affiches ou journalises.
function Read-C5Text([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Configuration manquante : $Path" }
    return [IO.File]::ReadAllText($Path).Trim()
}
function Test-C5Service([string]$ConfigPath, [string]$Role, [string]$Revision) {
    $root = Read-C5Text (Join-Path $ConfigPath 'chemin.txt')
    if ($root.TrimEnd('\') -ine '\\DS224\CabinetCardio') { throw "Racine de production incorrecte : $ConfigPath" }
    $urlText = Read-C5Text (Join-Path $ConfigPath 'service.url')
    if ([IO.File]::ReadAllText((Join-Path $ConfigPath 'service.url')) -match '[\r\n]') {
        throw "service.url doit tenir sur une ligne, sans retour final : $ConfigPath"
    }
    $url = $null
    if (-not [uri]::TryCreate($urlText,[UriKind]::Absolute,[ref]$url) -or $url.Scheme -ine 'https' -or $url.Port -eq 8444 -or $url.UserInfo -or $url.Query -or $url.Fragment -or $urlText -match '\s') {
        throw "URL HTTPS de production invalide ou URL de recette : $ConfigPath"
    }
    $ini = Read-C5Text (Join-Path $ConfigPath 'poste.ini')
    $section = ''; $profile = ''
    foreach ($line in ($ini -split '\r?\n')) {
        if ($line -match '^\s*\[([^]]+)\]\s*$') { $section = $matches[1] }
        elseif ($section -ieq 'POSTE' -and $line -match '^\s*Profil\s*=\s*(.*?)\s*$') { $profile = $matches[1] }
    }
    $expectedProfile = if ($Role -eq 'medecin') { 'CabinetMedecin' } else { 'CabinetSecretariat' }
    if ($profile -ine $expectedProfile) { throw "Profil $expectedProfile requis dans $ConfigPath\poste.ini" }
    $token = Read-C5Text (Join-Path $ConfigPath 'service.token')
    if ([IO.File]::ReadAllText((Join-Path $ConfigPath 'service.token')) -match '[\r\n]') {
        throw "service.token contient des retours de ligne incompatibles avec VBA : $ConfigPath"
    }
    if ($token.Length -lt 32 -or $token -match '\s') { throw "Jeton local invalide : $ConfigPath" }
    $http = $null
    try {
        $http = New-Object -ComObject WinHttp.WinHttpRequest.5.1
        $http.SetTimeouts(5000,5000,15000,15000); $http.Option(6) = $false
        $http.Open('POST',$url.AbsoluteUri.TrimEnd('/')+'/v1/rpc',$false)
        $http.SetRequestHeader('Content-Type','application/json')
        $http.SetRequestHeader('Authorization','Bearer '+$token)
        $http.Send('{"operation":"whoami","params":{}}')
        if ($http.Status -ne 200) { throw ('Authentification/service refuse : HTTP '+$http.Status) }
        $envelope = ([Text.Encoding]::UTF8.GetString([byte[]]$http.ResponseBody) | ConvertFrom-Json)
        $who = $envelope.result
        if (@($who.roles).Count -ne 1 -or $who.roles[0] -ine $Role -or $who.schema -ne 2 -or $who.protocole -ne 2 -or ($Revision -and $who.revision -cne $Revision)) {
            throw 'Role strict, revision, protocole ou schema du service incompatibles.'
        }
        return [pscustomobject]@{ Url=$url.AbsoluteUri.TrimEnd('/'); Revision=$who.revision; ID=$who.ID; Root=$root.TrimEnd('\') }
    } finally { $token=$null; if ($http) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($http) } }
}
