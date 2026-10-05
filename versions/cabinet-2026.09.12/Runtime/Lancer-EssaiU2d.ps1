[CmdletBinding()]
param([switch]$ControleDemarrage)
$ErrorActionPreference='Stop'
$app=$null; $book=$null; $success=$false
$result=[ordered]@{Date=(Get-Date -Format o);Statut='EN_COURS';RecetteClinique=$false}
function Invoke-Office($obj,[string]$name,[object[]]$values) {
    $obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::InvokeMethod,$null,$obj,$values)
}
try {
    $receipt=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'installation-essai.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $identity=[Security.Principal.WindowsIdentity]::GetCurrent().Name
    if($identity -ne $receipt.Compte -or $identity -notin @('AX8_MAX\CabinetU2Medecin','RDC\CabinetU2Test')){throw 'Ce raccourci est reserve au profil Windows de recette.'}
    $doctor=$identity -eq 'AX8_MAX\CabinetU2Medecin'
    $expectedRole=if($doctor){'medecin'}else{'secretariat'}
    $hostProcess=if($doctor){'WINWORD'}else{'EXCEL'}
    $session=(Get-Process -Id $PID).SessionId
    if(@(Get-Process $hostProcess -ErrorAction SilentlyContinue | Where-Object {$_.SessionId -eq $session}).Count){throw ('Fermez '+$hostProcess+' dans ce profil avant de relancer cet essai.')}
    $cfg=Join-Path $env:APPDATA 'CabinetCardio'
    $root=[IO.File]::ReadAllText((Join-Path $cfg 'chemin.txt')).Trim().TrimEnd('\')
    $url=[IO.File]::ReadAllText((Join-Path $cfg 'service.url')).Trim().TrimEnd('/')
    if($root -ne '\\DS224\CabinetCardioTestU2' -or $url -ne 'https://DS224:8444'){throw 'Configuration hors de la recette U2d : lancement refuse.'}
    $token=[IO.File]::ReadAllText((Join-Path $cfg 'service.token')).Trim()
    try {
        $http=New-Object -ComObject WinHttp.WinHttpRequest.5.1
        $http.SetTimeouts(5000,5000,15000,15000)
        $http.Open('POST',$url+'/v1/rpc',$false)
        $http.SetRequestHeader('Content-Type','application/json')
        $http.SetRequestHeader('Authorization','Bearer '+$token)
        $http.Send((@{operation='whoami';params=@{}}|ConvertTo-Json -Compress))
        if($http.Status -ne 200){throw 'Connexion au service de recette refusee.'}
        $who=([Text.Encoding]::UTF8.GetString([byte[]]$http.ResponseBody)|ConvertFrom-Json).result
        $correspondants=@()
        $offset=0
        do {
            $http.Open('POST',$url+'/v1/rpc',$false)
            $http.SetRequestHeader('Content-Type','application/json')
            $http.SetRequestHeader('Authorization','Bearer '+$token)
            $params=@{genre='CORRESPONDANTS';q='';year='';date='';offset=$offset}
            $http.Send((@{operation='table.read';params=$params}|ConvertTo-Json -Compress))
            if($http.Status -ne 200){throw 'Chargement des correspondants de recette refuse.'}
            $page=([Text.Encoding]::UTF8.GetString([byte[]]$http.ResponseBody)|ConvertFrom-Json).result
            foreach($item in @($page.items)){$correspondants += $item}
            if($page.PSObject.Properties.Name -notcontains 'next'){throw 'Pagination des correspondants incomplete.'}
            $next=$page.next
            if($null -ne $next){if([int]$next -le $offset){throw 'Pagination des correspondants invalide.'};$offset=[int]$next}
        } while($null -ne $next)
        $cacheDir=Join-Path $env:LOCALAPPDATA 'CabinetCardio'
        [void][IO.Directory]::CreateDirectory($cacheDir)
        $cachePath=Join-Path $cacheDir 'correspondants-session.json'
        $json=ConvertTo-Json -InputObject ([object[]]$correspondants) -Depth 10 -Compress
        [IO.File]::WriteAllText($cachePath,$json,(New-Object Text.UTF8Encoding($false)))
    }finally{$token=$null;if($http){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($http);$http=$null}}
    if(@($who.roles).Count -ne 1 -or $who.roles[0] -ne $expectedRole -or $who.schema -ne 2 -or $who.revision -ne '2026.10.01-u2d-rc1'){throw 'Role strict ou version du service de recette non conforme.'}
    $result.Compte=$identity;$result.Role=$expectedRole;$result.Revision=$who.revision
    if($doctor){
        $binary=[string](Join-Path $PSScriptRoot 'CabinetDragonC5NP10.dotm')
        if((Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash -ne $receipt.SHA256){throw 'Le modele Word de recette a change.'}
        $app=New-Object -ComObject Word.Application
        $app.AutomationSecurity=1
        $app.Visible=$true
        foreach($a in $app.AddIns){if($a.Installed -and $a.Name -match 'Cabinet|ModeleCourrierChatGPT'){throw 'Un autre complement Cabinet est deja charge.'}}
        $result.Etape='chargement_modele';$addin=Invoke-Office $app.AddIns 'Add' @([string]$binary,[bool]$true)
        $macro='modRaccourcis.AutoExec'
        $result.Etape='raccourcis';[void](Invoke-Office $app 'Run' @($macro))
        $result.Etape='racine_vba';$actualRoot=Invoke-Office $app 'Run' @('modConfig.racine')
        if($actualRoot -ne $root){throw 'La racine effectivement lue par Word est incorrecte.'}
        $result.Etape='bases_vba'
        [void](Invoke-Office $app 'Run' @('modBase.ChargerBases',[bool]$true))
        $result.Etape='connexion_vba'
        [void](Invoke-Office $app 'Run' @('modAttenteLocale.SynchroniserAttentes'))
        $result.ConnexionDepuisOffice=$true
        [void](Invoke-Office $app 'Run' @('modFileArrivees.Unifie_AfficherFileArrivees'))
        $result.Modele=$binary;$result.ModeleCharge=$addin.Installed
        $result.Documents=$app.Documents.Count
        $whoOffice=$null;$dict=$null;$addin=$null
    }else{
        $binary=Join-Path $PSScriptRoot 'Cabinet.xlsm'
        if(-not [IO.File]::Exists($binary)){throw 'Classeur de recette absent.'}
        $app=New-Object -ComObject Excel.Application
        $app.AutomationSecurity=2
        $app.EnableEvents=$true
        $app.Visible=$true
        $book=$app.Workbooks.Open($binary,0,$false)
        $macroPrefix="'"+$book.Name+"'!"
        $actualRoot=$app.Run($macroPrefix+'modConfig.racine')
        if($actualRoot -ne $root){throw 'La racine effectivement lue par Excel est incorrecte.'}
        $count=$app.Run($macroPrefix+'modEchange.NombreEnAttente')
        $app.Run($macroPrefix+'modEchange.DemarrerScrutation')
        $book.Worksheets.Item('Accueil').Activate()
        $result.ConnexionDepuisOffice=$true
        $result.Classeur=$book.FullName;$result.LectureSeule=$book.ReadOnly
        $result.CourriersEnAttente=$count
        $result.Accueil=[string]$book.Worksheets.Item('Accueil').Range('B4').Value2
        if($book.ReadOnly){throw 'Le classeur local est en lecture seule.'}
    }
    $result.Statut='SUCCES';$success=$true
}catch{
    $result.Statut='ECHEC';$result.Erreur=$_.Exception.Message;$result.Pile=$_.ScriptStackTrace
    if(-not $ControleDemarrage){Add-Type -AssemblyName System.Windows.Forms;[void][Windows.Forms.MessageBox]::Show($result.Erreur,'Cabinet - ESSAI U2d')}
}finally{
    if($app -and ($ControleDemarrage -or -not $success)){
        if($doctor){try{[void](Invoke-Office $app 'Run' @('modFileArrivees.FermerFileArrivees'))}catch{};try{[void](Invoke-Office $app 'Quit' @(0))}catch{}}
        else{try{if($book){$app.Run("'"+$book.Name+"'!modEchange.ArreterScrutation");$book.Close($false)}}catch{};try{$app.Quit()}catch{}}
    }
    if($book){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    $book=$null;$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
    $report=Join-Path $PSScriptRoot $(if($ControleDemarrage){'controle-demarrage.json'}else{'dernier-demarrage.json'})
    $result|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $report -Encoding UTF8
    $result|ConvertTo-Json -Depth 5 -Compress
}
if(-not $success){exit 1}
