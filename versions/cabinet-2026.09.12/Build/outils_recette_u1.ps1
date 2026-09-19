function Verifier-ResultatRecetteSimple([string]$Json,[string]$Nom,[int]$Minimum) {
    try { $r=ConvertFrom-Json -InputObject $Json -ErrorAction Stop } catch { throw ($Nom+' : resultat JSON invalide.') }
    if ($null -eq $r -or $r -is [Array]) { throw ($Nom+' : resultat absent ou non objet.') }
    $props=@($r.PSObject.Properties.Name)
    if ('echec' -notin $props -or 'reussis' -notin $props -or $r.echec -isnot [bool]) { throw ($Nom+' : compte rendu incomplet.') }
    [int]$nombre=0
    if (-not [int]::TryParse([string]$r.reussis,[ref]$nombre) -or $r.echec -or $nombre -lt $Minimum) {
        $detail=if ('description' -in $props) { [string]$r.description } else { '' }
        throw ($Nom+' en echec ou incomplete : '+$detail)
    }
    return $r
}

function Verifier-ResultatRecetteOffice([string]$Json,[string]$Nom) {
    try { $r=ConvertFrom-Json -InputObject $Json -ErrorAction Stop } catch { throw ($Nom+' : resultat JSON invalide.') }
    if ($null -eq $r -or $r -is [Array]) { throw ($Nom+' : resultat absent ou non objet.') }
    $props=@($r.PSObject.Properties.Name)
    foreach ($propriete in @('reussis','attendus','echec')) {
        if ($propriete -notin $props) { throw ($Nom+' : champ '+$propriete+' absent.') }
    }
    if ($r.echec -isnot [bool]) { throw ($Nom+' : champ echec invalide.') }
    [int]$reussis=0
    [int]$attendus=0
    if (-not [int]::TryParse([string]$r.reussis,[ref]$reussis) -or -not [int]::TryParse([string]$r.attendus,[ref]$attendus)) {
        throw ($Nom+' : compteurs invalides.')
    }
    $detail=if ('description' -in $props) { [string]$r.description } else { '' }
    if ($r.echec -or $attendus -lt 1 -or $reussis -ne $attendus) {
        throw ($Nom+' incomplete ou en echec ('+$reussis+'/'+$attendus+') : '+$detail)
    }
    return $r
}

function Compiler-ProjetU1($application, $projet) {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    if (-not ('CabinetU1Native' -as [type])) {
        Add-Type 'using System;using System.Runtime.InteropServices;public class CabinetU1Native { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p); [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l); }'
    }
    $vbe=$application.VBE
    $vbe.MainWindow.Visible=$true
    $projet.VBComponents.Item(1).CodeModule.CodePane.Show()
    if([string]$vbe.ActiveVBProject.FileName -ne [string]$projet.FileName){throw 'Le projet a compiler n est pas celui de la copie attendue.'}
    $ownerPid=[uint32]0
    [void][CabinetU1Native]::GetWindowThreadProcessId([IntPtr]$vbe.MainWindow.HWnd,[ref]$ownerPid)
    if($ownerPid -eq 0){throw 'Instance Office de recette non identifiee.'}
    $compile=$vbe.CommandBars.FindControl(1,578)
    if($null -eq $compile){throw 'Commande de compilation absente.'}
    if($compile.Enabled){$compile.Execute()}
    Start-Sleep -Milliseconds 200
    if($compile.Enabled){
        $messages=@()
        $windows=[System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children,[System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$ownerPid))
        foreach($window in $windows){
            $elements=$window.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
            $dialogs=@($elements|Where-Object{$_.Current.ClassName -eq '#32770' -and $_.Current.Name -eq 'Microsoft Visual Basic pour Applications'})
            foreach($dialog in $dialogs){
                $children=$dialog.FindAll([System.Windows.Automation.TreeScope]::Children,[System.Windows.Automation.Condition]::TrueCondition)
                $message=@($children|Where-Object{$_.Current.ClassName -eq 'Static'}|ForEach-Object{$_.Current.Name}) -join ' '
                $messages+=$message
                if($message -like '*compilation*'){
                    foreach($button in $children){
                        if($button.Current.ProcessId -eq [int]$ownerPid -and $button.Current.Name -eq 'OK' -and $button.Current.ClassName -eq 'Button'){
                            [void][CabinetU1Native]::PostMessage([IntPtr]$button.Current.NativeWindowHandle,245,[IntPtr]::Zero,[IntPtr]::Zero)
                        }
                    }
                }
            }
        }
        Start-Sleep -Milliseconds 500
        $details=''
        $pane=$vbe.ActiveCodePane
        if($null -ne $pane){
            $line=0;$column=0;$lastLine=0;$lastColumn=0
            $pane.GetSelection([ref]$line,[ref]$column,[ref]$lastLine,[ref]$lastColumn)
            $details=$pane.CodeModule.Name+':'+$line+':'+$column+' '+$pane.CodeModule.Lines([Math]::Max(1,$line),1)
        }
        throw ('Compilation refusee : '+($messages -join ' ; ')+' '+$details)
    }
    $vbe.MainWindow.Visible=$false
}
