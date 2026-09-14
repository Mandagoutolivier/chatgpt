Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'

function Ecrire-EtatAssistant([string]$Chemin,$Etat) {
    [void][IO.Directory]::CreateDirectory((Split-Path $Chemin -Parent))
    $tmp=$Chemin+'.'+[guid]::NewGuid().ToString('N')+'.tmp'
    try {
        [IO.File]::WriteAllText($tmp,($Etat | ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $Chemin) { [IO.File]::Replace($tmp,$Chemin,[NullString]::Value) }
        else { [IO.File]::Move($tmp,$Chemin) }
    } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp } }
}

function Lire-EtatAssistant([string]$Chemin) {
    if (-not (Test-Path -LiteralPath $Chemin)) { return $null }
    Get-Content -LiteralPath $Chemin -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Choisir-ProfilAssistant {
    Write-Host ''
    Write-Host '1 - DOMICILE : medecin et secretariat sur ce PC'
    Write-Host '2 - SECRETARIAT : patients, agenda, actes et impressions'
    Write-Host '3 - CABINET : poste medecin'
    do { $choice=Read-Host 'Votre choix (1, 2 ou 3)' } until ($choice -in @('1','2','3'))
    return @{'1'='Domicile';'2'='CabinetSecretariat';'3'='CabinetMedecin'}[$choice]
}

function Attendre-FermetureOffice {
    param(
        [ValidateSet('WINWORD','EXCEL')][string[]]$Noms=@('WINWORD','EXCEL'),
        [ValidateRange(0,30)][int]$DelaiSecondes=30
    )
    $running=@(Get-Process -Name $Noms -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) { return }
    Write-Host 'Attente de la fermeture complete des processus Office...'
    # Quit() peut revenir avant la disparition du processus. Laisser ce delai avant de solliciter l operateur.
    for ($i=0;$i -lt ($DelaiSecondes*4) -and $running.Count -gt 0;$i++) {
        Start-Sleep -Milliseconds 250
        $running=@(Get-Process -Name $Noms -ErrorAction SilentlyContinue)
    }
    while ($running.Count -gt 0) {
        Write-Host 'Enregistrez vos documents et fermez Word et Excel, y compris les fenetres reduites.'
        foreach ($process in $running) {
            Write-Host ('Encore actif : '+$process.ProcessName+' ; PID '+$process.Id+' ; session Windows '+$process.SessionId)
        }
        Write-Host 'Si aucune fenetre n est visible, ouvrez le Gestionnaire des taches pour verifier les processus indiques.'
        Write-Host 'Le lanceur ne force aucune fermeture et ne supprime aucun document.'
        $choice=Read-Host 'Entree apres fermeture pour continuer ; Q pour reprendre plus tard'
        if ($choice -eq 'Q') { throw 'Installation en pause. Relancez le meme fichier pour reprendre.' }
        $running=@(Get-Process -Name $Noms -ErrorAction SilentlyContinue)
    }
}

function Tester-RegleAccesVba([object[]]$Valeurs) {
    foreach ($value in $Valeurs) {
        if ($null -ne $value -and [int]$value -eq 0) { throw 'Une strategie administrateur interdit l acces au projet VBA. Elle ne sera pas modifiee par ce lanceur.' }
    }
}

function Restaurer-AccesVbaAssistant([string]$Journal) {
    if (-not (Test-Path -LiteralPath $Journal)) { return }
    $saved=Lire-EtatAssistant $Journal
    if ($saved.sid -ne [Security.Principal.WindowsIdentity]::GetCurrent().User.Value) { throw 'Le journal Office appartient a une autre session Windows.' }
    foreach ($item in @($saved.items)) {
        $current=Get-ItemProperty -LiteralPath $item.path -Name AccessVBOM -ErrorAction SilentlyContinue
        if ($null -ne $current -and [int]$current.AccessVBOM -eq 1) {
            if ($item.existed) { Set-ItemProperty -LiteralPath $item.path -Name AccessVBOM -Value $item.value }
            else { Remove-ItemProperty -LiteralPath $item.path -Name AccessVBOM }
        }
    }
    Remove-Item -LiteralPath $Journal
}

function Autoriser-AccesVbaAssistant([string]$Journal) {
    Attendre-FermetureOffice
    Restaurer-AccesVbaAssistant $Journal
    $saved=[pscustomobject]@{sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;items=@()}
    try {
    foreach ($hostName in @('Word','Excel')) {
        $app=$null
        try {
            $app=New-Object -ComObject ($hostName+'.Application')
            $version=[string]$app.Version
        } finally {
            if ($null -ne $app) {
                try { $app.Quit() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
            }
        }
        if ($version -notmatch '^\d+\.\d+$') { throw 'Version Office inattendue.' }
        $policies=@()
        foreach ($base in @('HKCU:\Software\Policies','HKLM:\Software\Policies','HKLM:\Software\WOW6432Node\Policies')) {
            $path=Join-Path $base ("Microsoft\Office\$version\$hostName\Security")
            $p=Get-ItemProperty -LiteralPath $path -Name AccessVBOM -ErrorAction SilentlyContinue
            if ($null -ne $p) { $policies+=$p.AccessVBOM }
        }
        Tester-RegleAccesVba $policies
        $path="HKCU:\Software\Microsoft\Office\$version\$hostName\Security"
        $old=Get-ItemProperty -LiteralPath $path -Name AccessVBOM -ErrorAction SilentlyContinue
        if ($null -ne $old -and [int]$old.AccessVBOM -eq 1) { continue }
        $value=$null;if ($null -ne $old) { $value=$old.AccessVBOM }
        $saved.items+=@([pscustomobject]@{path=$path;existed=($null -ne $old);value=$value})
        Ecrire-EtatAssistant $Journal $saved
        if (-not (Test-Path -LiteralPath $path)) { [void](New-Item -Path $path -Force) }
        New-ItemProperty -LiteralPath $path -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
    }
    } catch {
        $cause=$_
        try { Restaurer-AccesVbaAssistant $Journal } catch { Write-Warning 'Restauration AccessVBOM incomplete : journal conserve.' }
        throw $cause
    }
}

function Tester-DeconnexionOfficeAssistant([Exception]$Exception) {
    # PowerShell enveloppe les erreurs COM dans plusieurs InnerException.
    # Classer par HRESULT uniquement, sans dependance a la langue du message.
    $current=$Exception
    for ($depth=0;$depth -lt 16 -and $null -ne $current;$depth++) {
        if ($current.HResult.ToString('X8') -in @('800706BA','80010108','80010006','80010007','80010012')) { return $true }
        $current=$current.InnerException
    }
    return $false
}

function Compiler-ProjetAssistant([string]$Fichier,[ValidateSet('Word','Excel')][string]$Hote) {
    while ($true) {
        try { return (Compiler-ProjetAssistantUneTentative $Fichier $Hote) }
        catch {
            if (-not (Tester-DeconnexionOfficeAssistant $_.Exception)) { throw }
            $erreurCom=$_.Exception
            Write-Host ''
            Write-Host "$Hote ne repond plus au lanceur. L application a pu etre fermee ou s arreter."
            Write-Host ('Detail : '+$erreurCom.Message)
            Write-Host 'La compilation et l enregistrement ne sont pas confirmes pour cette tentative.'
            Write-Host 'R rouvre le meme fichier prepare et redemande la compilation. Q conserve la preparation pour plus tard.'
            do { $choice=Read-Host 'R = reprendre la compilation ; Q = mettre en pause' } until ($choice -in @('R','Q'))
            if ($choice -eq 'Q') { throw [InvalidOperationException]::new("Compilation $Hote en pause ; preparation conservee, aucune validation de cette tentative.",$erreurCom) }
            $processName=if ($Hote -eq 'Word') { 'WINWORD' } else { 'EXCEL' }
            Attendre-FermetureOffice -Noms @($processName)
        }
    }
}

function Compiler-ProjetAssistantUneTentative([string]$Fichier,[ValidateSet('Word','Excel')][string]$Hote) {
    # Office ne fournit pas de compilateur VBA en ligne de commande documente.
    # Ouvrir le bon projet et faire confirmer la commande native, sans SendKeys.
    $app=$null;$document=$null;$project=$null;$components=$null;$component=$null;$code=$null;$pane=$null;$vbe=$null;$window=$null
    try {
        Write-Host "Preparation de la compilation $Hote : $Fichier"
        $app=New-Object -ComObject ($Hote+'.Application')
        $app.AutomationSecurity=3
        if ($Hote -eq 'Word') { $document=$app.Documents.Open($Fichier,$false,$false) }
        else { $app.EnableEvents=$false;$document=$app.Workbooks.Open($Fichier,0,$false) }
        $app.Visible=$true
        # ActiveVBProject est en lecture seule dans le modele VBA documente.
        # Le volet de code appartient au document prepare, jamais au projet actif par defaut.
        $project=$document.VBProject
        if ($null -eq $project) { throw 'Projet VBA inaccessible pour le fichier prepare.' }
        $components=$project.VBComponents
        $moduleName=if ($Hote -eq 'Word') { 'ThisDocument' } else { 'ThisWorkbook' }
        $component=$components.Item($moduleName)
        $code=$component.CodeModule
        if ($null -eq $code) { throw "Module de code introuvable : $moduleName" }
        Write-Host ''
        Write-Host ('Fichier a compiler : '+$document.FullName)
        Write-Host ('Projet VBA : '+$project.Name+' ; module : '+$moduleName)
        try {
            $pane=$code.CodePane
            [void]$pane.Show()
            $vbe=$app.VBE
            $window=$vbe.MainWindow
            $window.Visible=$true
        } catch {
            # Un echec d affichage n est ni un succes de compilation ni un refus d acces au projet.
            Write-Host ('Affichage automatique de l editeur indisponible : '+$_.Exception.Message)
            Write-Host "Dans $Hote, appuyez sur Alt+F11, puis Ctrl+R pour afficher les projets."
        }
        Write-Host 'Dans l editeur VBA, selectionnez le module indique sous le fichier a compiler.'
        Write-Host 'Choisissez Debogage > Compiler.'
        Write-Host "LAISSEZ $Hote ET LE FICHIER OUVERTS. Revenez dans cette console avec Alt+Tab."
        Write-Host "Apres votre OUI, le lanceur enregistrera le fichier puis fermera $Hote lui-meme."
        Write-Host 'Si une erreur apparait, notez-la et repondez NON.'
        do { $answer=Read-Host "Compilation $Hote sans erreur ? OUI / NON" } until ($answer -in @('OUI','NON'))
        if ($answer -ne 'OUI') { throw "Compilation $Hote non validee. La preparation est conservee ; aucune activation." }
        if (-not [string]::Equals([string]$document.FullName,$Fichier,[StringComparison]::OrdinalIgnoreCase)) {
            throw 'Le fichier ouvert a change de nom ou de dossier. Reprendre la compilation du fichier prepare.'
        }
        [void]$document.Save()
        if (-not $document.Saved) { throw 'Office ne confirme pas l enregistrement du fichier prepare.' }
        return $true
    } catch {
        throw [InvalidOperationException]::new(("Compilation $Hote interrompue pour $Fichier : "+$_.Exception.Message),$_.Exception)
    } finally {
        if ($null -ne $document) { try { [void]$document.Close($false) } catch {} }
        if ($null -ne $app) { try { [void]$app.Quit() } catch {} }
        foreach ($comObject in @($window,$vbe,$pane,$code,$component,$components,$project,$document,$app)) {
            if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                # Un objet deja deconnecte ne doit pas masquer l erreur initiale lors du nettoyage.
                try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch {}
            }
        }
    }
}

function Confirmer-RecetteAssistant([string]$Guide,[string]$Dossier) {
    Write-Host ''
    Write-Host 'La compilation est terminee. Les essais fonctionnels ne peuvent pas etre inventes.'
    Write-Host 'Le guide de recette et le dossier des fichiers prepares vont s ouvrir.'
    Start-Process -FilePath notepad.exe -ArgumentList ('"'+$Guide+'"')
    Start-Process -FilePath explorer.exe -ArgumentList ('"'+$Dossier+'"')
    Write-Host 'Utilisez uniquement un service et des patients fictifs pour ces essais.'
    Write-Host 'RECETTE : vous avez termine les essais requis pour ce poste.'
    Write-Host 'PAUSE : conserver la preparation et reprendre plus tard.'
    do { $answer=Read-Host 'Votre choix : RECETTE ou PAUSE' } until ($answer -in @('RECETTE','PAUSE'))
    return ($answer -eq 'RECETTE')
}

function Empreinte-SourcesAssistant([string]$Racine) {
    $json=Empreintes-Sources $Racine | ConvertTo-Json -Compress
    $hash=[Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($json)))).Replace('-','') }
    finally { $hash.Dispose() }
}

function Actualiser-CompilationAssistant($Etat) {
    foreach ($item in @(@('wordCompile','wordHash','CabinetUnifie.dotm'),@('excelCompile','excelHash','Cabinet.xlsm'))) {
        if (-not $Etat.($item[0])) { continue }
        $file=Join-Path $Etat.dossierPrepare $item[2]
        if (-not (Test-Path -LiteralPath $file) -or (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ne $Etat.($item[1])) {
            $Etat.($item[0])=$false;$Etat.($item[1])='';$Etat.recette=$false
        }
    }
}


function Normaliser-CheminNasAssistant([string]$Chemin) {
    $value=$Chemin.Trim().Trim('"').TrimEnd('\')
    if ($value -notmatch '^\\\\[^\\]+\\[^\\]+(?:\\[^\\]+)*$' -or $value -match '[<>:"|?*\x00-\x1f]') {
        throw 'Saisissez un chemin reseau complet : \\serveur\partage\dossier. Une lettre comme Z: ne suffit pas.'
    }
    foreach ($part in $value.Substring(2).Split('\')) {
        if ($part -in @('.','..') -or $part -match '[. ]$') { throw 'Chemin reseau invalide.' }
    }
    return $value
}

function Tester-DossierCodeAssistant([string]$Dossier) {
    foreach ($marker in @('.git','Build\manifest.json','Installer.ps1','ModelesSource')) {
        if (Test-Path -LiteralPath (Join-Path $Dossier $marker)) { return $true }
    }
    return $false
}

function Choisir-RacineNasAssistant([string]$Demandee,[string]$Memorisee,[string]$FichierChemin) {
    $candidate=$Demandee
    if (-not $candidate) { $candidate=$Memorisee }
    if (-not $candidate -and (Test-Path -LiteralPath $FichierChemin -PathType Leaf)) {
        $candidate=([IO.File]::ReadAllText($FichierChemin,[Text.Encoding]::UTF8)).Trim()
    }
    Write-Host ''
    Write-Host 'DOSSIER DES DONNEES SUR LE NAS'
    Write-Host 'Choisissez le dossier de donnees commun aux postes, ou un dossier neuf pour une nouvelle installation.'
    Write-Host 'Un dossier contenant le depot GitHub ne constitue pas une base patients.'
    Write-Host 'Tous les postes devront acceder aux memes donnees avec leurs comptes NAS respectifs.'
    while ($true) {
        if ($candidate) { Write-Host ('Chemin propose ou saisi : '+$candidate) }
        Write-Host 'B = parcourir le reseau ; Q = reprendre plus tard.'
        $answer=Read-Host 'Chemin UNC complet, B, Q ou Entree pour conserver le chemin affiche'
        if ($answer -eq 'Q') { return $null }
        if ($answer -eq 'B') {
            Add-Type -AssemblyName System.Windows.Forms
            $picker=New-Object Windows.Forms.FolderBrowserDialog
            try {
                $picker.Description='Dossier des donnees du cabinet sur le NAS (pas le dossier du code GitHub)'
                $picker.ShowNewFolderButton=$true
                $picker.SelectedPath='\\DS224'
                if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Container)) { $picker.SelectedPath=$candidate }
                if ($picker.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { continue }
                $candidate=$picker.SelectedPath
            } finally { $picker.Dispose() }
        } elseif ($answer) { $candidate=$answer }
        if (-not $candidate) { Write-Host 'Aucun dossier n a encore ete choisi.';continue }
        try { $candidate=Normaliser-CheminNasAssistant $candidate }
        catch { Write-Host $_.Exception.Message;continue }
        if (-not (Test-Path -LiteralPath $candidate -PathType Container)) {
            # Creer un sous-dossier seulement dans un parent existant, jamais un partage SMB.
            $parent=$candidate.Substring(0,$candidate.LastIndexOf('\'))
            if ($parent -match '^\\\\[^\\]+\\[^\\]+' -and (Test-Path -LiteralPath $parent -PathType Container)) {
                Write-Host ('Ce dossier n existe pas : '+$candidate)
                Write-Host 'Une creation prepare un emplacement neuf ; elle ne deplace ni ne retrouve vos anciennes bases.'
                $creation=Read-Host 'Tapez CREER pour creer ce dossier, ou Entree pour choisir un autre chemin'
                if ($creation -ne 'CREER') { continue }
                try { Creer-DossierNasAssistant $candidate }
                catch { Write-Host ('Creation impossible : '+$_.Exception.Message);continue }
            } else {
                Write-Host ('Dossier inaccessible : '+$candidate)
                Write-Host 'Verifiez le partage existant et la connexion au NAS. A domicile, connectez le VPN Cabinet Freebox Pro.'
                continue
            }
        }
        if (Tester-DossierCodeAssistant $candidate) { Write-Host 'Ce dossier contient du code de deploiement. Choisissez le dossier des donnees, ou un sous-dossier neuf distinct.';continue }
        if ($candidate -match '^\\\\[^\\]+\\homes?(?:\\|$)') {
            Write-Host 'Attention : le dossier home depend du compte NAS. Verifiez que les comptes des deux postes accedent au MEME dossier physique.'
            $shared=Read-Host 'Tapez COMMUN si cet acces commun est verifie, sinon Entree pour choisir un autre dossier'
            if ($shared -ne 'COMMUN') { continue }
        }
        try { Tester-EcritureNasAssistant $candidate }
        catch { Write-Host ('Ecriture impossible dans ce dossier : '+$_.Exception.Message);continue }
        Write-Host ('Dossier NAS retenu : '+$candidate)
        return $candidate
    }
}


function Creer-DossierNasAssistant([string]$Dossier) {
    [void][IO.Directory]::CreateDirectory($Dossier)
}

function Tester-EcritureNasAssistant([string]$Dossier) {
    $probe=Join-Path $Dossier ('.cabinet-ecriture-'+[guid]::NewGuid().ToString('N')+'.tmp')
    $stream=$null;$created=$false
    try {
        $stream=[IO.File]::Open($probe,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        $created=$true;$stream.WriteByte(0)
    } finally { if ($null -ne $stream) { $stream.Dispose() };if ($created) { [IO.File]::Delete($probe) } }
}
