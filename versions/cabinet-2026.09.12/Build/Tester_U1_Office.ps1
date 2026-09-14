[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Sortie)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'outils_assistant.ps1')
. (Join-Path $PSScriptRoot 'outils_recette_u1.ps1')
if(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue){throw 'Office deja ouvert : aucun processus existant ne sera utilise.'}
[void][IO.Directory]::CreateDirectory($Sortie)
$journal=Join-Path $Sortie 'acces-vba-a-restaurer.json'
$rapport=Join-Path $Sortie 'validation-office-en-cours.log'
$word=$null;$doc=$null;$excel=$null;$wb=$null
function Trace-U1($text){$text|Add-Content -LiteralPath $rapport -Encoding UTF8}
try{
    Trace-U1 'Debut de qualification sur copies separees ; aucune activation.'
    Autoriser-AccesVbaAssistant $journal
    & (Join-Path $PSScriptRoot 'construire_modele_unifie.ps1') -Prod6 (Join-Path $root 'ModelesSource\ModeleCourrierChatGPT_PROD(6).dotm') -Cabinet1 (Join-Path $root 'ModelesSource\Cabinet(1).dotm') -Sortie (Join-Path $Sortie 'CabinetUnifie.dotm') -RacineSources $root
    Trace-U1 'Modele Word construit.'
    & (Join-Path $PSScriptRoot 'construire_cabinet_secretariat.ps1') -CabinetXlsm (Join-Path $root 'ModelesSource\Cabinet.xlsm') -Sortie (Join-Path $Sortie 'Cabinet.xlsm') -RacineSources $root
    Trace-U1 'Classeur Excel construit.'
    Attendre-FermetureOffice
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3
    $doc=$word.Documents.Open((Join-Path $Sortie 'CabinetUnifie.dotm'),$false,$false,$false)
    Trace-U1 'Compilation Word demandee.'
    Compiler-ProjetU1 $word $doc.VBProject
    Trace-U1 'Compilation Word terminee.'
    $doc.Save()
    $noSave=[object]0
    $doc.Close([ref]$noSave);$doc=$null
    # Execution limitee au modele de recette copie ci-dessus.
    $word.AutomationSecurity=1
    $doc=$word.Documents.Open((Join-Path $Sortie 'CabinetUnifie.dotm'),$false,$false,$false)
    $testArgument=[object]$Sortie
    $result=$word.Run('modRecetteU1.Executer',[ref]$testArgument)
    Trace-U1 ('Tests Word : '+[string]$result)
    $recette=ConvertFrom-Json -InputObject ([string]$result) -ErrorAction Stop
    if($null -eq $recette -or $recette.echec -ne $false -or [int]$recette.reussis -lt 26){throw ('Recette Word incomplete ou en echec : '+[string]$recette.description)}
    $doc.Close([ref]$noSave);$doc=$null;$word.Quit([ref]$noSave);$word=$null
    Attendre-FermetureOffice
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    $wb=$excel.Workbooks.Open((Join-Path $Sortie 'Cabinet.xlsm'),0,$false)
    Trace-U1 'Compilation Excel demandee.'
    Compiler-ProjetU1 $excel $wb.VBProject
    $wb.Save();Trace-U1 'Compilation Excel terminee.'
    $wb.Close($false);$wb=$null;$excel.Quit();$excel=$null
    Trace-U1 'SUCCES ; aucune validation clinique ou impression papier effectuee.'
}catch{
    Trace-U1 ('ECHEC : '+$_.Exception.Message)
    throw
}finally{
    # Chaque nettoyage est independant ; aucun ne peut empecher la restauration.
    $noSave=[object]0
    if($null -ne $doc){try{$doc.Close([ref]$noSave)}catch{Trace-U1 'Fermeture copie Word a verifier.'}}
    if($null -ne $wb){try{$wb.Close($false)}catch{Trace-U1 'Fermeture copie Excel a verifier.'}}
    if($null -ne $word){try{$word.Quit([ref]$noSave)}catch{Trace-U1 'Fermeture instance Word a verifier.'}}
    if($null -ne $excel){try{$excel.Quit()}catch{Trace-U1 'Fermeture instance Excel a verifier.'}}
    try{Restaurer-AccesVbaAssistant $journal;Trace-U1 'Acces VBA restaure.'}
    finally{[IO.File]::Copy($rapport,(Join-Path $Sortie ('validation-office-finale-'+[guid]::NewGuid().ToString('N')+'.log')),$false)}
}
