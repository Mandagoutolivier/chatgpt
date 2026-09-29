[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Racine,[Parameter(Mandatory=$true)][string]$Sortie,[Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{40}$')][string]$CommitSource)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
if([Security.Principal.WindowsIdentity]::GetCurrent().Name -notin @('AX8_MAX\CabinetU2Medecin','RDC\CabinetU2Test')){throw 'Compte de recette requis'}
if(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue){throw 'Office deja ouvert'}
if(Test-Path $Sortie){throw 'Sortie neuve requise'}
[void][IO.Directory]::CreateDirectory($Sortie)
. (Join-Path $Racine 'Build\outils_construction.ps1')
. (Join-Path $Racine 'Build\outils_assistant.ps1')
. (Join-Path $Racine 'Build\outils_recette_u1.ps1')
. (Join-Path $Racine 'Build\outils_installation.ps1')
$normal=Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
$hadNormal=Test-Path $normal
$normalHash=if($hadNormal){(Get-FileHash $normal).Hash}else{''}
if($hadNormal){Copy-Item $normal (Join-Path $Sortie 'Normal-avant.dotm')}
$journal=Join-Path $Sortie 'acces-vba-a-restaurer.json'
$report=[ordered]@{CommitSource=$CommitSource;Statut='EN_COURS';Binaires=@();OfficeFerme=$false;NormalRestaure=$false;AccesVbaRestaure=$false;ActivationEffectuee=$false;RecetteClinique=$false}
function Finaliser { [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers() }
function Canon([string]$s){$parts=[regex]::Split(($s -replace '\r','').Trim(),'("(?:[^"]|"")*")');for($i=0;$i -lt $parts.Length;$i+=2){$parts[$i]=$parts[$i].ToLowerInvariant()};return ($parts -join '')}
try{
 Autoriser-AccesVbaAssistant $journal
 & (Join-Path $Racine 'Build\construire_modele_unifie.ps1') -Prod6 (Join-Path $Racine 'ModelesSource\ModeleCourrierChatGPT_PROD(6).dotm') -Cabinet1 (Join-Path $Racine 'ModelesSource\Cabinet(1).dotm') -Sortie (Join-Path $Sortie 'CabinetUnifie.dotm') -RacineSources $Racine
 & (Join-Path $Racine 'Build\construire_cabinet_secretariat.ps1') -CabinetXlsm (Join-Path $Racine 'ModelesSource\Cabinet.xlsm') -Sortie (Join-Path $Sortie 'Cabinet.xlsm') -RacineSources $Racine
 Finaliser;Attendre-FermetureOffice
 $manifest=Lire-Manifeste $Racine
 foreach($hostName in @('word','excel')){
  $app=$null;$doc=$null;$components=$null;$refs=$null;$project=$null
  $filename=if($hostName -eq 'word'){'CabinetUnifie.dotm'}else{'Cabinet.xlsm'}
  $path=Join-Path $Sortie $filename
  try{
   if($hostName -eq 'word'){$app=New-Object -ComObject Word.Application;$app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3;$app.WordBasic.DisableAutoMacros(1);$doc=$app.Documents.Open($path,$false,$false,$false)}
   else{$app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3;$doc=$app.Workbooks.Open($path,0,$false)}
   $project=$doc.VBProject
   Compiler-ProjetU1 $app $project
   $doc.Save();$doc.Close(0);$doc=$null;$project=$null
   if($hostName -eq 'word'){$doc=$app.Documents.Open($path,$false,$true,$false)}else{$doc=$app.Workbooks.Open($path,0,$true)}
   $project=$doc.VBProject;$names=@($manifest.$hostName|ForEach-Object{$_.name});$count=0
   foreach($component in $project.VBComponents){if($component.Name -notin $names){throw ('Composant non livre prevu : '+$component.Name)};$count++}
   if($count -ne $names.Count){throw 'Nombre de composants different'}
   foreach($item in $manifest.$hostName){$component=$project.VBComponents.Item($item.name);$actual=Lire-ModuleVba $component.CodeModule;$expected=Lire-CodeVba (Join-Path $Racine $item.path);if((Canon $actual) -cne (Canon $expected)){throw ('Source differente : '+$item.name)}}
   $references=@();foreach($reference in $project.References){if($reference.IsBroken){throw ('Reference manquante : '+$reference.Name)};$references+=@([pscustomobject]@{Nom=$reference.Name;Guid=$reference.Guid;Major=$reference.Major;Minor=$reference.Minor})}
   $report.Binaires+=@([pscustomobject]@{Fichier=$filename;Compilation=$true;Reouverture=$true;Composants=$count;SourcesConformes=$true;References=$references})
   $doc.Close(0);$doc=$null;$app.Quit();$app=$null
  }finally{
   if($null -ne $doc){try{$doc.Close(0)}catch{}}
   if($null -ne $app){try{$app.Quit()}catch{}}
   $doc=$null;$app=$null;$project=$null;$component=$null;$reference=$null;Finaliser;Attendre-FermetureOffice
  }
 }
 $env:CABINET_SOURCE_COMMIT=$report.CommitSource
 Ecrire-Preparation $Sortie 'Domicile' $Racine
 Verifier-Preparation $Sortie 'Domicile' $Racine
 foreach($item in $report.Binaires){$item|Add-Member -NotePropertyName SHA256 -NotePropertyValue ((Get-FileHash (Join-Path $Sortie $item.Fichier)).Hash)}
 $report.Statut='SUCCESS'
}catch{$report.Statut='ECHEC';$report['Erreur']=$_.Exception.Message;throw}
finally{
 Finaliser;Attendre-FermetureOffice;$report.OfficeFerme=$true
 Restaurer-AccesVbaAssistant $journal;$report.AccesVbaRestaure=$true
 if($hadNormal){Copy-Item (Join-Path $Sortie 'Normal-avant.dotm') $normal -Force;if((Get-FileHash $normal).Hash -ne $normalHash){throw 'Normal non restaure'}}elseif(Test-Path $normal){Copy-Item $normal (Join-Path $Sortie 'Normal-cree.dotm');Remove-Item $normal}
 $report.NormalRestaure=$true;$report['DateUTC']=[DateTime]::UtcNow.ToString('o')
 $report|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Sortie 'qualification-binaires.json') -Encoding UTF8
 $report|ConvertTo-Json -Depth 8
}
