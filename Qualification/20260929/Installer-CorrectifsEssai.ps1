[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$identity=[Security.Principal.WindowsIdentity]::GetCurrent().Name
if($identity -notin @('AX8_MAX\CabinetU2Medecin','RDC\CabinetU2Test')){throw 'Profil de recette obligatoire'}
$session=(Get-Process -Id $PID).SessionId
if(@(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue|Where-Object{$_.SessionId -eq $session}).Count){throw 'Fermer Office dans le profil de recette'}
$doctor=$identity -eq 'AX8_MAX\CabinetU2Medecin'
$role=if($doctor){'medecin'}else{'secretariat'}
$file=if($doctor){'CabinetUnifie.dotm'}else{'Cabinet.xlsm'}
$old=if($doctor){'AC941632E4C5A16EC202E6067BB4E78CF978154FB4A7FDA4BB7FA6FD96770DF2'}else{'3A975BDF3879016DE64C3495EED0A6C23050EAA1D774B2389D34D48FEC4355FD'}
$new=if($doctor){'C8DC1F1EF7040FC6DCBAEC3DE553C5BFC64CDD38153846529C1FDC2BE21A8A67'}else{'5293E7E3D9AB416BDCCE5680F22B9F044F17540879BF3594993FA3476926DC79'}
$sourceCommit='06506de063c76cc6811da46b5c503e5b549fc571'
$root='\\DS224\CabinetCardioTestU2'
$source=Join-Path $root 'Patients\_Qualification20260929\CorrectifsFichePatient-06506de'
$dest=Join-Path $env:LOCALAPPDATA 'CabinetCardioTestU2\EssaiU2c20260929'
$binary=Join-Path $dest $file
$receiptFile=Join-Path $dest 'installation-essai.json'
$receipt=Get-Content $receiptFile -Raw -Encoding UTF8|ConvertFrom-Json
if($receipt.Compte -ne $identity -or $receipt.Role -ne $role -or $receipt.Dossier -ne $dest -or $receipt.ActivationClinique){throw 'Recu installation incorrect'}
if((Get-FileHash $binary).Hash -ne $old -or $receipt.SHA256 -ne $old){throw 'Version installee inattendue'}
if((Get-FileHash (Join-Path $source $file)).Hash -ne $new){throw 'Binaire source non conforme'}
$qualification=Get-Content (Join-Path $source 'qualification-binaires.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($qualification.Statut -ne 'SUCCESS' -or $qualification.CommitSource -ne $sourceCommit -or $qualification.ActivationEffectuee){throw 'Qualification binaire incorrecte'}
$nas=@(Get-ChildItem (Join-Path $source 'mise-a-jour-*.json')|ForEach-Object{Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object{$_.statut -eq 'SUCCESS' -and $_.source -eq $sourceCommit})
if($nas.Count -ne 1 -or -not $nas[0].source_installee){throw 'Mise a jour NAS non confirmee'}
$cfg=Join-Path $env:APPDATA 'CabinetCardio'
if([IO.File]::ReadAllText((Join-Path $cfg 'chemin.txt')).Trim().TrimEnd('\') -ne $root){throw 'Racine incorrecte'}
$url=[IO.File]::ReadAllText((Join-Path $cfg 'service.url')).Trim().TrimEnd('/')
if($url -ne 'https://DS224:8444'){throw 'URL incorrecte'}
$http=New-Object -ComObject WinHttp.WinHttpRequest.5.1
try{
 $http.SetTimeouts(5000,5000,15000,15000);$http.Open('POST',$url+'/v1/rpc',$false)
 $token=[IO.File]::ReadAllText((Join-Path $cfg 'service.token')).Trim()
 $http.SetRequestHeader('Authorization','Bearer '+$token);$token=$null
 $http.SetRequestHeader('Content-Type','application/json');$http.Send('{"operation":"whoami","params":{}}')
 if($http.Status -ne 200){throw 'Connexion refusee'}
 $who=($http.ResponseText|ConvertFrom-Json).result
 if(@($who.roles).Count -ne 1 -or $who.roles[0] -ne $role -or $who.schema -ne 2 -or $who.revision -ne '2026.09.21-u2c'){throw 'Role ou revision incorrecte'}
}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($http)}
$normal=Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
$protected=@{}
foreach($p in @($normal,(Join-Path $dest 'Lancer-EssaiU2c.ps1')) + @($receipt.Raccourcis) + @('chemin.txt','poste.ini','service.url','service.token','openai.key'|ForEach-Object{Join-Path $cfg $_})){
 if(Test-Path -LiteralPath $p){$protected[$p]=(Get-FileHash -LiteralPath $p).Hash}
}
$trustBefore=Get-ItemProperty $receipt.EmplacementApprouve
if($trustBefore.Path.TrimEnd('\') -ne $dest -or $trustBefore.AllowSubfolders -ne 0){throw 'Emplacement approuve incorrect'}
$backup=Join-Path $env:LOCALAPPDATA ('U2Q29\AvantCorrectifs-'+(Get-Date -Format 'yyyyMMdd-HHmmss'))
if(Test-Path $backup){throw 'Sauvegarde deja presente'}
[void][IO.Directory]::CreateDirectory($backup)
Copy-Item $binary (Join-Path $backup $file)
Copy-Item $receiptFile (Join-Path $backup 'installation-essai.json')
Copy-Item (Join-Path $dest 'Lancer-EssaiU2c.ps1') (Join-Path $backup 'Lancer-EssaiU2c.ps1')
if(Test-Path $normal){Copy-Item $normal (Join-Path $backup 'Normal-avant.dotm')}
$report=[ordered]@{Statut='EN_COURS';Compte=$identity;Role=$role;SourceCommit=$sourceCommit;Sauvegarde=$backup;ActivationClinique=$false;RecetteClinique=$false}
try{
 Copy-Item (Join-Path $source $file) $binary -Force
 if((Get-FileHash $binary).Hash -ne $new){throw 'Copie non conforme'}
 $receipt.SHA256=$new;$receipt.SourceCommit=$sourceCommit
 $receipt|Add-Member -NotePropertyName CorrectifsDate -NotePropertyValue (Get-Date -Format o) -Force
 $receipt|Add-Member -NotePropertyName SauvegardeCorrectifs -NotePropertyValue $backup -Force
 $receipt|ConvertTo-Json -Depth 8|Set-Content $receiptFile -Encoding UTF8
 & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $dest 'Lancer-EssaiU2c.ps1') -ControleDemarrage
 if($LASTEXITCODE -ne 0){throw 'Controle de demarrage echoue'}
 $check=Get-Content (Join-Path $dest 'controle-demarrage.json') -Raw|ConvertFrom-Json
 if($check.Statut -ne 'SUCCES' -or -not $check.ConnexionDepuisOffice){throw 'Connexion depuis Office non validee'}
 foreach($p in $protected.Keys){if((Get-FileHash -LiteralPath $p).Hash -ne $protected[$p]){throw ('Fichier preserve modifie : '+[IO.Path]::GetFileName($p))}}
 if((Get-FileHash $binary).Hash -ne $new){throw 'Binaire altere au demarrage'}
 $trustAfter=Get-ItemProperty $receipt.EmplacementApprouve
 if($trustAfter.Path -ne $trustBefore.Path -or $trustAfter.AllowSubfolders -ne $trustBefore.AllowSubfolders){throw 'Parametres de confiance modifies'}
 $report.Statut='SUCCESS';$report['SHA256']=$new;$report['ControleOffice']=$check
 $report['ConfigurationEtRaccourcisConserves']=$true
}catch{
 $report.Statut='ECHEC';$report['Erreur']=$_.Exception.Message
 if(@(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue|Where-Object{$_.SessionId -eq $session}).Count){$report['RetourArriere']='A_VERIFIER_OFFICE_ENCORE_OUVERT'}
 else{
  Copy-Item (Join-Path $backup $file) $binary -Force
  Copy-Item (Join-Path $backup 'installation-essai.json') $receiptFile -Force
  $report['RetourArriere']=if((Get-FileHash $binary).Hash -eq $old){'OK'}else{'A_VERIFIER'}
 }
 throw
}finally{
 $report['Date']=(Get-Date -Format o)
 $report|ConvertTo-Json -Depth 8|Set-Content (Join-Path $dest 'installation-correctifs.json') -Encoding UTF8
 $report|ConvertTo-Json -Depth 8
}
