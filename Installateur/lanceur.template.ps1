# Ce modele est assemble avec les outils et les empreintes par generer_lanceur.py.
# Aucun mot de passe ni jeton GitHub n est demande ou conserve.
$Commit='@@COMMIT@@'
$Empreintes=@'
@@HASHES@@
'@ | ConvertFrom-Json
$bootLock=$null
try {
    if ($env:OS -ne 'Windows_NT') { throw 'Ce lanceur necessite Windows.' }
    Write-Host ('Cabinet Cardio - installation guidee - version '+$Commit.Substring(0,7))
    Write-Host 'Utilisez votre session Windows habituelle, sans Executer en tant qu administrateur.'
    $cache=Join-Path $env:LOCALAPPDATA 'CabinetCardio\Installation\Sources'
    [void][IO.Directory]::CreateDirectory($cache)
    $bootLock=[IO.File]::Open((Join-Path $cache 'telechargement.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $package=Join-Path $cache $Commit
    [void](Isoler-PaquetInvalide $package $Empreintes)
    if (-not (Test-Path -LiteralPath $package)) {
        $url="https://github.com/Mandagoutolivier/chatgpt/archive/$Commit.zip"
        Write-Host ''
        Write-Host 'Le navigateur va telecharger une version precise du depot prive.'
        Write-Host 'Si GitHub demande une connexion, utilisez votre compte ayant acces au depot.'
        Write-Host 'Si une page 404 apparait : connectez-vous a github.com, puis rouvrez le lien ci-dessous.'
        Write-Host $url
        $downloads=Join-Path $env:USERPROFILE 'Downloads'
        $known=Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name '{374DE290-123F-4565-9164-39C4925E467B}' -ErrorAction SilentlyContinue
        if ($null -ne $known) { $downloads=[Environment]::ExpandEnvironmentVariables($known.'{374DE290-123F-4565-9164-39C4925E467B}') }
        Start-Process $url
        Write-Host 'Recherche du ZIP dans Telechargements pendant 45 secondes...'
        $archive=$null;$deadline=[DateTime]::UtcNow.AddSeconds(45)
        do {
            $candidate=Get-ChildItem -LiteralPath $downloads -Filter "chatgpt-$Commit*.zip" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
            if ($null -ne $candidate) {
                $probe=$null
                try {
                    $probe=[IO.File]::Open($candidate.FullName,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
                    $archive=$candidate.FullName
                } catch {} finally { if ($null -ne $probe) { $probe.Dispose() } }
            }
            if (-not $archive) { Start-Sleep -Seconds 2 }
        } until ($archive -or [DateTime]::UtcNow -ge $deadline)
        if (-not $archive) {
            Write-Host 'Selectionnez le ZIP telecharge. Annuler permet de reprendre plus tard.'
            Add-Type -AssemblyName System.Windows.Forms
            $picker=New-Object Windows.Forms.OpenFileDialog
            try {
                $picker.Title='Selectionner le ZIP Cabinet Cardio telecharge depuis GitHub'
                $picker.Filter='Archive ZIP (*.zip)|*.zip';$picker.InitialDirectory=$downloads
                if ($picker.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { Write-Host 'Installation en pause.';exit 0 }
                $archive=$picker.FileName
            } finally { $picker.Dispose() }
        }
        $temporary=$package+'.'+[guid]::NewGuid().ToString('N')
        Write-Host 'Extraction et verification de chaque fichier...'
        Extraire-PaquetCabinet $archive $temporary $Commit $Empreintes
        [IO.Directory]::Move($temporary,$package)
    }
    Tester-PaquetCabinet $package $Empreintes
    # Le deblocage porte uniquement sur les fichiers verifies ; aucune politique globale n est modifiee.
    Get-ChildItem -LiteralPath $package -Recurse -File | Unblock-File
    $bootLock.Dispose();$bootLock=$null
    Write-Host ('Sources utilisees : '+$package)
    # Transmission au recu de preparation ; les sources viennent d'etre verifiees.
    $env:CABINET_SOURCE_COMMIT=$Commit
    & (Join-Path $package 'Build\assistant_installation.ps1')
    exit 0
} catch {
    Write-Host ''
    Write-Host ('Installation interrompue : '+$_.Exception.Message) -ForegroundColor Red
    if ($_.InvocationInfo -and $_.InvocationInfo.ScriptName) {
        Write-Host ('Etape : '+[IO.Path]::GetFileName($_.InvocationInfo.ScriptName)+' ; ligne '+$_.InvocationInfo.ScriptLineNumber)
    }
    Write-Host 'Relancez ce meme fichier apres correction. Les eventuels fichiers deja prepares sont conserves.'
    exit 1
} finally { if ($null -ne $bootLock) { $bootLock.Dispose() } }
