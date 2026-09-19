[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Sortie, [switch]$RecetteU2,
      [string]$RacineSources='', [switch]$HorsReseau)
$ErrorActionPreference='Stop'
$root=if ($RacineSources) { [IO.Path]::GetFullPath($RacineSources) } else { Split-Path $PSScriptRoot -Parent }
. (Join-Path $PSScriptRoot 'outils_assistant.ps1')
. (Join-Path $PSScriptRoot 'outils_recette_u1.ps1')
if ($HorsReseau) {
    if ($env:CABINET_QUALIFICATION_ISOLEE -cne [IO.Path]::GetFullPath($Sortie)) { throw 'Qualification isolee : utiliser Tester_U2_Office_Isole.ps1.' }
    . (Join-Path $PSScriptRoot 'outils_recette_isolee.ps1')
}
if(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue){throw 'Office deja ouvert : aucun processus existant ne sera utilise.'}
[void][IO.Directory]::CreateDirectory($Sortie)
$journal=Join-Path $Sortie 'acces-vba-a-restaurer.json'
$rapport=Join-Path $Sortie 'validation-office-en-cours.log'
$word=$null;$doc=$null;$excel=$null;$wb=$null;$wordUpdateLinksAvant=$null
$normalPath=Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
$normalBackup=Join-Path $Sortie 'Normal-avant-recette.dotm'
$normalSurveiller=$false;$normalExistait=$false;$normalHash='';$testsValides=$false;$accesRestaure=$false
function Trace-U1($text){$text|Add-Content -LiteralPath $rapport -Encoding UTF8}
function Finaliser-ObjetsOfficeRecette {
    # Apres le retour des scripts/fonctions, finaliser aussi leurs collections et enumerateurs COM.
    # Quit() peut laisser Office actif tant que ces objets devenus inaccessibles ne sont pas liberes.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
try{
    Trace-U1 'Debut de qualification sur copies separees ; aucune activation.'
    $normalExistait=Test-Path -LiteralPath $normalPath
    if($normalExistait){
        [IO.File]::Copy($normalPath,$normalBackup,$false)
        $normalHash=(Get-FileHash -LiteralPath $normalPath -Algorithm SHA256).Hash
        if((Get-FileHash -LiteralPath $normalBackup -Algorithm SHA256).Hash -ne $normalHash){throw 'Sauvegarde Normal non conforme.'}
    }
    $normalSurveiller=$true
    Autoriser-AccesVbaAssistant $journal
    & (Join-Path $PSScriptRoot 'construire_modele_unifie.ps1') -Prod6 (Join-Path $root 'ModelesSource\ModeleCourrierChatGPT_PROD(6).dotm') -Cabinet1 (Join-Path $root 'ModelesSource\Cabinet(1).dotm') -Sortie (Join-Path $Sortie 'CabinetUnifie.dotm') -RacineSources $root -InclureRecette
    Trace-U1 'Modele Word construit.'
    & (Join-Path $PSScriptRoot 'construire_cabinet_secretariat.ps1') -CabinetXlsm (Join-Path $root 'ModelesSource\Cabinet.xlsm') -Sortie (Join-Path $Sortie 'Cabinet.xlsm') -RacineSources $root -InclureRecette
    Trace-U1 'Classeur Excel construit.'
    Finaliser-ObjetsOfficeRecette
    Attendre-FermetureOffice
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3
    $doc=$word.Documents.Open((Join-Path $Sortie 'CabinetUnifie.dotm'),$false,$false,$false)
    if ($HorsReseau) {
        Preparer-CopieRecetteIsolee $doc.VBProject $Sortie 'Word'
        $word.WordBasic.DisableAutoMacros(1)
        $wordUpdateLinksAvant=$word.Options.UpdateLinksAtOpen
        $word.Options.UpdateLinksAtOpen=$false
    }
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
    if($null -eq $recette -or $recette.echec -ne $false -or [int]$recette.reussis -lt 50){throw ('Recette Word incomplete ou en echec : '+[string]$recette.description)}
    if($RecetteU2){
        $resultU2=$word.Run('modRecetteU2.ExecuterU2',[ref]$testArgument)
        Trace-U1 ('Tests U2 : '+[string]$resultU2)
        $u2=Verifier-ResultatRecetteOffice ([string]$resultU2) 'Recette U2 Word'
    }
    $resultModele=$word.Run('modRecetteModeleCourrier.ExecuterModeleCourrier',[ref]$testArgument)
    Trace-U1 ('Tests modele courrier : '+[string]$resultModele)
    $modele=Verifier-ResultatRecetteSimple ([string]$resultModele) 'Recette modele courrier' 38
    if ($modele.modele_reel_verifie -ne $false) { throw 'Modele prive non autorise dans la recette consolidee.' }
    $resultPresentation=$word.Run('modRecettePresentationAnnexe.ExecuterPresentationAnnexe')
    Trace-U1 ('Tests presentation annexe : '+[string]$resultPresentation)
    $presentation=Verifier-ResultatRecetteSimple ([string]$resultPresentation) 'Recette presentation annexe' 67
    if ([int]$presentation.reussis -ne 67) { throw 'Nombre de controles presentation inattendu.' }
    $doc.Close([ref]$noSave);$doc=$null
    if ($null -ne $wordUpdateLinksAvant) { $word.Options.UpdateLinksAtOpen=$wordUpdateLinksAvant;$wordUpdateLinksAvant=$null }
    $word.Quit([ref]$noSave);$word=$null
    Finaliser-ObjetsOfficeRecette
    Attendre-FermetureOffice
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    $wb=$excel.Workbooks.Open((Join-Path $Sortie 'Cabinet.xlsm'),0,$false)
    if ($HorsReseau) { Preparer-CopieRecetteIsolee $wb.VBProject $Sortie 'Excel' }
    Trace-U1 'Compilation Excel demandee.'
    Compiler-ProjetU1 $excel $wb.VBProject
    $wb.Save();Trace-U1 'Compilation Excel terminee.'
    $wb.Close($false);$wb=$null
    $excel.AutomationSecurity=1
    $wb=$excel.Workbooks.Open((Join-Path $Sortie 'Cabinet.xlsm'),0,$false)
    $resultExcel=$excel.Run("'"+$wb.Name+"'!modRecetteU1Excel.Executer")
    Trace-U1 ('Tests Excel : '+[string]$resultExcel)
    $recetteExcel=Verifier-ResultatRecetteOffice ([string]$resultExcel) 'Recette Excel'
    $wb.Close($false);$wb=$null;$excel.Quit();$excel=$null
    $resultats=[ordered]@{WordU1=$recette;WordU2=$null;ModeleCourrier=$modele;PresentationAnnexe=$presentation;Excel=$recetteExcel;CopiesInstrumentees=[bool]$HorsReseau;ActivationEffectuee=$false;RecetteClinique=$false}
    if ($RecetteU2) { $resultats.WordU2=$u2 }
    [IO.File]::WriteAllText((Join-Path $Sortie 'resultats-suites-office.json'),($resultats|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
    $testsValides=$true
}catch{
    Trace-U1 ('ECHEC : '+$_.Exception.Message)
    throw
}finally{
    # Tenter chaque fermeture independamment ; Office doit avoir quitte avant la restauration.
    $noSave=[object]0
    if($null -ne $doc){try{$doc.Close([ref]$noSave)}catch{Trace-U1 'Fermeture copie Word a verifier.'}}
    if($null -ne $wb){try{$wb.Close($false)}catch{Trace-U1 'Fermeture copie Excel a verifier.'}}
    if($null -ne $word){
        try { if ($null -ne $wordUpdateLinksAvant) { $word.Options.UpdateLinksAtOpen=$wordUpdateLinksAvant } }
        catch { Trace-U1 'Restauration UpdateLinksAtOpen a verifier.' }
        finally { try { $word.Quit([ref]$noSave) } catch { Trace-U1 'Fermeture instance Word a verifier.' } }
    }
    if($null -ne $excel){try{$excel.Quit()}catch{Trace-U1 'Fermeture instance Excel a verifier.'}}
    $doc=$null;$wb=$null;$word=$null;$excel=$null
    try{
        Finaliser-ObjetsOfficeRecette
        Attendre-FermetureOffice
        # Office peut reecrire sa configuration en quittant. Garder le journal jusqu a sa fermeture.
        Trace-U1 'Fermeture complete Office confirmee avant restauration.'
        try{Restaurer-AccesVbaAssistant $journal;$accesRestaure=$true;Trace-U1 'Acces VBA restaure.'}
        finally{
            if($normalSurveiller){
                $normalApres=''
                if(Test-Path -LiteralPath $normalPath){$normalApres=(Get-FileHash -LiteralPath $normalPath -Algorithm SHA256).Hash}
                if($normalApres -ne $normalHash){
                    if($normalApres){
                        $copie=Join-Path $Sortie 'Normal-apres-recette.dotm'
                        [IO.File]::Copy($normalPath,$copie,$false)
                        if((Get-FileHash -LiteralPath $copie -Algorithm SHA256).Hash -ne $normalApres){throw 'Conservation de Normal apres recette non conforme.'}
                    }
                    if($normalExistait){[IO.File]::Copy($normalBackup,$normalPath,$true)}
                    elseif(Test-Path -LiteralPath $normalPath){[IO.File]::Delete($normalPath)}
                }
                if($normalExistait){
                    if((Get-FileHash -LiteralPath $normalPath -Algorithm SHA256).Hash -ne $normalHash){throw 'Normal initial non restaure.'}
                }elseif(Test-Path -LiteralPath $normalPath){throw 'Normal cree pendant recette non retire.'}
                Trace-U1 'Normal initial restaure ; toute copie modifiee par Office est conservee.'
            }
            if($testsValides -and $accesRestaure){Trace-U1 'SUCCES ; aucune validation clinique ou impression papier effectuee.'}
        }
    }catch{
        Trace-U1 ('NETTOYAGE INCOMPLET : '+$_.Exception.Message)
        throw
    }finally{[IO.File]::Copy($rapport,(Join-Path $Sortie ('validation-office-finale-'+[guid]::NewGuid().ToString('N')+'.log')),$false)}
}
