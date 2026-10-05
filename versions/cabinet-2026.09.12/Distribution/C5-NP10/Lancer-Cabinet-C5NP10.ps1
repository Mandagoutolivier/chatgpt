#requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory=$true)][ValidateSet('Medecin','Secretariat')][string]$Profil)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$app=$null; $book=$null; $addin=$null
function Invoke-C5Office($Object,[string]$Method,[object[]]$Values) {
    $Object.GetType().InvokeMember($Method,[Reflection.BindingFlags]::InvokeMethod,$null,$Object,$Values)
}
try {
    $expectedIdentity = if ($Profil -eq 'Medecin') { 'AX8_MAX\olivi' } else { 'ACCUEIL\accueil' }
    if ([Security.Principal.WindowsIdentity]::GetCurrent().Name -ine $expectedIdentity) { throw "Lancement reserve a $expectedIdentity" }
    if (@(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq (Get-Process -Id $PID).SessionId }).Count) {
        throw 'Fermer Office dans cette session avant de lancer C5.'
    }
    $receipt = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'installation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    . (Join-Path $PSScriptRoot 'Controle-Service-C5.ps1')
    $role = if ($Profil -eq 'Medecin') { 'medecin' } else { 'secretariat' }
    $config = Join-Path $env:APPDATA 'CabinetCardio'
    $who = Test-C5Service $config $role $receipt.revision
    if ($Profil -eq 'Medecin') {
        $binary = Join-Path $PSScriptRoot 'CabinetUnifie.dotm'
        if ((Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash -ine $receipt.word) { throw 'Modele Word modifie.' }
        # Charger les correspondants AVANT AutoExec. Sans ce cache, ChargerBases echoue.
        $items = New-Object 'System.Collections.Generic.List[object]'; $offset=0; $http=$null
        $token=Read-C5Text (Join-Path $config 'service.token')
        try {
            $http=New-Object -ComObject WinHttp.WinHttpRequest.5.1
            $http.SetTimeouts(5000,5000,15000,15000); $http.Option(6)=$false
            for ($page=0;$page -lt 5000;$page++) {
                $http.Open('POST',$who.Url+'/v1/rpc',$false)
                $http.SetRequestHeader('Content-Type','application/json');$http.SetRequestHeader('Authorization','Bearer '+$token)
                $body=@{operation='table.read';params=@{genre='CORRESPONDANTS';q='';year='';date='';offset=$offset}} | ConvertTo-Json -Compress
                $http.Send($body)
                if ($http.Status -ne 200) { throw 'Chargement des correspondants refuse.' }
                $r=([Text.Encoding]::UTF8.GetString([byte[]]$http.ResponseBody) | ConvertFrom-Json).result
                foreach ($item in @($r.items)) { $items.Add($item) }
                if ($r.PSObject.Properties.Name -notcontains 'next') { throw 'Pagination incomplete.' }
                if ($null -eq $r.next) { break }
                if ([int]$r.next -le $offset) { throw 'Pagination invalide.' };$offset=[int]$r.next
            }
            if ($page -ge 5000) { throw 'Trop de pages de correspondants.' }
            $cache = Join-Path $env:LOCALAPPDATA 'CabinetCardio'
            [void][IO.Directory]::CreateDirectory($cache)
            [IO.File]::WriteAllText((Join-Path $cache 'correspondants-session.json'),(ConvertTo-Json -InputObject @($items.ToArray()) -Depth 10 -Compress),(New-Object Text.UTF8Encoding($false)))
        } finally { $token=$null;if($http){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($http)} }
        $app=New-Object -ComObject Word.Application
        $app.Visible=$true
        # Respecter le niveau de securite Office existant ; aucune baisse globale.
        foreach ($a in $app.AddIns) { if ($a.Installed -and $a.Name -match 'Cabinet|ModeleCourrierChatGPT') { throw 'Un autre complement Cabinet est charge.' } }
        $app.AutomationSecurity=2 # msoAutomationSecurityByUI
        $addin=Invoke-C5Office $app.AddIns 'Add' @([string]$binary,[bool]$true)
        [void](Invoke-C5Office $app 'Run' @('modBase.ChargerBases',[bool]$true))
        if ((Invoke-C5Office $app 'Run' @('modConfig.racine')) -ine $who.Root) { throw 'Racine effectivement lue par Word incorrecte.' }
        [void](Invoke-C5Office $app 'Run' @('modRaccourcis.AutoExec'))
        [void](Invoke-C5Office $app 'Run' @('modAttenteLocale.SynchroniserAttentes'))
        [void](Invoke-C5Office $app 'Run' @('modFileArrivees.Unifie_AfficherFileArrivees'))
    } else {
        $binary=Join-Path $PSScriptRoot 'Cabinet.xlsm'
        # Excel peut sauvegarder sa vue d'agenda ; hash integral verifie a l'installation.
        if (-not (Test-Path -LiteralPath $binary -PathType Leaf)) { throw 'Classeur absent.' }
        $app=New-Object -ComObject Excel.Application;$app.Visible=$true;$app.EnableEvents=$true
        $app.AutomationSecurity=2 # msoAutomationSecurityByUI
        $book=$app.Workbooks.Open($binary,0,$false)
        if ($book.ReadOnly) { throw 'Classeur en lecture seule.' }
        $prefix="'"+$book.Name+"'!"
        if ($app.Run($prefix+'modConfig.racine') -ine $who.Root) { throw 'Racine effectivement lue par Excel incorrecte.' }
        [void]$app.Run($prefix+'modEchange.NombreEnAttente')
        [void]$app.Run($prefix+'modEchange.DemarrerScrutation')
        [void]$app.Run($prefix+'modAgendaVue.AgendaAujourdhui')
        $book.Worksheets.Item('Accueil').Activate()
    }
} catch {
    Write-Host ('ECHEC : '+$_.Exception.Message) -ForegroundColor Red
    if ($app) {
        try {
            if ($Profil -eq 'Medecin') { [void](Invoke-C5Office $app 'Quit' @(0)) }
            else { if($book){$book.Close($false)};$app.Quit() }
        } catch {}
    }
    [void](Read-Host 'Appuyez sur Entree pour fermer')
    exit 1
} finally {
    if($book){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($addin){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin)}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
