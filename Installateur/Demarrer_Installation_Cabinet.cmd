@echo off
setlocal
set "CABINET_SELF=%~f0"
set "CABINET_BOOT_DIR=%TEMP%\CabinetBoot-%RANDOM%-%RANDOM%"
if exist "%CABINET_BOOT_DIR%" goto collision
mkdir "%CABINET_BOOT_DIR%"
if errorlevel 1 exit /b 1
set "CABINET_BOOT_PS=%CABINET_BOOT_DIR%\demarrer.ps1"
powershell.exe -NoLogo -NoProfile -Command "$ErrorActionPreference='Stop';$s=[IO.File]::ReadAllText($env:CABINET_SELF);$m=[regex]::Match($s,'(?m)^# CABINET_POWERSHELL_PAYLOAD_V1\r?$');if(-not $m.Success){throw 'Lanceur incomplet'};[IO.File]::WriteAllText($env:CABINET_BOOT_PS,$s.Substring($m.Index+$m.Length),(New-Object Text.UTF8Encoding($true)))"
if errorlevel 1 goto erreur
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "%CABINET_BOOT_PS%"
set "CABINET_RESULT=%ERRORLEVEL%"
goto nettoyage
:erreur
set "CABINET_RESULT=1"
:nettoyage
if exist "%CABINET_BOOT_PS%" del "%CABINET_BOOT_PS%"
rmdir "%CABINET_BOOT_DIR%"
echo.
echo Vous pouvez fermer cette fenetre. En cas de pause, relancez ce meme fichier.
pause
exit /b %CABINET_RESULT%
:collision
echo Relancez le fichier : le dossier temporaire existe deja.
pause
exit /b 1
# CABINET_POWERSHELL_PAYLOAD_V1
Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'

function Tester-PaquetCabinet([string]$Dossier,$Empreintes) {
    if (-not (Test-Path -LiteralPath $Dossier -PathType Container)) { throw 'Paquet absent.' }
    if ((Get-Item -LiteralPath $Dossier).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Lien de dossier refuse.' }
    $entries=@(Get-ChildItem -LiteralPath $Dossier -Recurse -Force)
    foreach ($file in $entries) {
        if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Lien dans le paquet refuse.' }
    }
    $files=@($entries | Where-Object { -not $_.PSIsContainer })
    if ($files.Count -ne @($Empreintes.PSObject.Properties).Count) { throw 'Liste des fichiers du paquet differente de la version attendue.' }
    foreach ($file in $files) {
        $relative=$file.FullName.Substring($Dossier.TrimEnd([char[]]@('\','/')).Length+1).Replace('\','/')
        $p=$Empreintes.PSObject.Properties[$relative]
        if ($null -eq $p -or (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash -ne $p.Value) { throw "Fichier absent du manifeste ou modifie : $relative" }
    }
}

function Extraire-PaquetCabinet([string]$Archive,[string]$Destination,[string]$Commit,$Empreintes) {
    if ($Commit -notmatch '^[a-f0-9]{40}$') { throw 'Identifiant de version invalide.' }
    if (Test-Path -LiteralPath $Destination) { throw 'Le dossier de destination doit etre nouveau.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=$null;$created=$false
    try {
        $zip=[IO.Compression.ZipFile]::OpenRead($Archive)
        if ($zip.Entries.Count -gt 10000) { throw 'Archive trop volumineuse.' }
        $prefix="chatgpt-$Commit/versions/cabinet-2026.09.12/"
        $selected=New-Object 'System.Collections.Generic.List[object]'
        $seen=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        [long]$total=0
        foreach ($entry in $zip.Entries) {
            if (-not $entry.FullName.StartsWith($prefix,[StringComparison]::Ordinal)) { continue }
            $name=$entry.FullName.Substring($prefix.Length)
            if (-not $name -or $name.EndsWith('/')) { continue }
            $segments=$name.Split('/')
            if ($name -match '[\\:\x00-\x1f]' -or @($segments | Where-Object { $_ -in @('','..','.') -or $_ -match '[. ]$' }).Count) { throw 'Chemin invalide dans le ZIP.' }
            if (-not $seen.Add($name)) { throw 'Doublon dans le ZIP.' }
            $kind=($entry.ExternalAttributes -shr 16) -band 0xF000
            if ($kind -ne 0 -and $kind -ne 0x8000) { throw 'Entree non reguliere dans le ZIP.' }
            if ($null -eq $Empreintes.PSObject.Properties[$name]) { throw "Fichier non attendu dans le ZIP : $name" }
            $total+=$entry.Length
            if ($entry.Length -gt 100MB -or $total -gt 300MB) { throw 'Taille decompressee excessive.' }
            $selected.Add([pscustomobject]@{entry=$entry;name=$name})
        }
        if ($selected.Count -ne @($Empreintes.PSObject.Properties).Count) { throw 'ZIP incomplet ou version incorrecte. Telechargez le lien indique par le lanceur.' }
        [void][IO.Directory]::CreateDirectory($Destination);$created=$true
        foreach ($item in $selected) {
            $target=Join-Path $Destination $item.name
            [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
            $inputStream=$item.entry.Open();$outputStream=$null
            try {
                $outputStream=[IO.File]::Open($target,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
                $inputStream.CopyTo($outputStream)
            } finally { $inputStream.Dispose();if ($null -ne $outputStream) { $outputStream.Dispose() } }
        }
        Tester-PaquetCabinet $Destination $Empreintes
    } catch {
        if ($created -and (Test-Path -LiteralPath $Destination)) { Remove-Item -LiteralPath $Destination -Recurse -Force }
        throw
    } finally { if ($null -ne $zip) { $zip.Dispose() } }
}

# Ce modele est assemble avec les outils et les empreintes par generer_lanceur.py.
# Aucun mot de passe ni jeton GitHub n est demande ou conserve.
$Commit='395e39b106cabdb01925ea2e34e0d480d74c8368'
$Empreintes=@'
{
  ".dockerignore": "30eae15fc1fa5b14ff10f03e58ed06cb18f519802455a73df93d0de4bf6d807f",
  "AUDIT.md": "d28f2a32618dea4ab44bf848879a848ab55e0788cc723346bc2570a2e35517a9",
  "Build/assistant_installation.ps1": "d7be691963c57e092a3dd627edb28d57a34cc33d91490d90444dee226e989b06",
  "Build/construire_cabinet_secretariat.ps1": "a4e271f3c26a5dab8ce53120b07d4201a61f6dd484275f032206ec6528da8645",
  "Build/construire_modele_unifie.ps1": "63b1a4c28f0d50bd29294950914bbc9514b78ee9020ee4c796487f07d594680b",
  "Build/donnees_initiales.json": "2c6a672f018d9db815c4f3fcdc91892db9b9a43bb3c7477bc83f49cb4b23b790",
  "Build/initialiser_nas.ps1": "4598eac4dc1a80addb0aefd486f109d82087000c64912abea66c854bc6f507f1",
  "Build/installer_multi_postes.ps1": "ce4243c2b52cf346ea9e494506d2d9c6b9d59c2d72b412b0b54883bae8bbfea1",
  "Build/installer_sqlite_medecin.ps1": "0d6221fa88894d767fdad5d8a6eb290a52c27830629b5557fc73eed52673ebc7",
  "Build/manifest.json": "0b2384ffe12c10ee35c13fa5875777e5239f7551f1ece24fb8c522ebd4413f71",
  "Build/outils_assistant.ps1": "4bddbf7fc4899d16eca967694269d1db9f14a935177d35310ed0f2faa9839582",
  "Build/outils_construction.ps1": "2fe847faa5a4be2af45ff73ecd31e984b5c405dafcd478fc9fc6ee22d57650e1",
  "Build/outils_installation.ps1": "daa32259aca3c5266f2f6114eb6be008efcf02a31bcd5ee95b862fee0178e793",
  "Build/outils_telechargement.ps1": "a20c18c69d5bbd823a72d036d047cd20505985440529fe03dbd9323ecf74d705",
  "Build/restaurer_poste.ps1": "055022fe8b814f2d2f2cf0950d952608a91b97621833d0e464b3ee16527e0127",
  "Build/ruban_unifie.xml": "23440d0ff03843e9f40f7e02d62a4addbdbeed85a558ff85674e8bf40e429978",
  "Build/schemas.json": "9b6c70a98f9129de98bf9aa1f1f6714b714e1610352801e0d806ab7254f3d0d4",
  "Build/sqlite.lock.json": "fb253313df37c7dace2cda37ef2bdb59111fe19e39720155a64647e0b1759326",
  "Build/valider_preparation.ps1": "923deb54a89d041f86cf3a39a492c2c6fde738ebff55fe7eb6fcfe372caa617a",
  "DonneesInitiales/Base/base_travail_correspondants_v1.xlsx": "233baf23b5094a42901b9ca06e83b6904354f5b24a19c7d0b89b339253ff2922",
  "DonneesInitiales/Config/DDE/declencheurs_demandes.txt": "430fb6e84a5e6137bb4707f6336a2453db47792460d400f58ecd8c76245a2f2d",
  "DonneesInitiales/Config/DDE/examens_complementaires.txt": "b5b753eec87a7aac522701922da5bb4d358d4da51c49d6cc8aa8e710562bbfe0",
  "DonneesInitiales/Config/DDE/exclusions_demandes.txt": "0881891d72a5834d2a5aa54ee0c1ac2441f5b0cad5b69a4478c5dc9fb558c07b",
  "DonneesInitiales/Config/Gras_Expressions.xlsx": "748fc3736765a8b6a23aead5154a3fec3365aacfe2f6cc14369ab9dd24f423ab",
  "DonneesInitiales/Config/Gras_Medicaments.xlsx": "c7934222c5e02b613f73dffd9a509521fbfa03735ec2ce51d46f1ee8f698d07a",
  "DonneesInitiales/Config/Nomenclature.xlsx": "dc6ac8c5478f193e7a927c1a7aea17d9bbc4407144068d539005fd4908d7ad43",
  "DonneesInitiales/Config/cerfa_positions.txt": "f09aece2ca125f32bb303bbbd16b9dbdc20e006807d0977d5fa95bdeb95a9876",
  "DonneesInitiales/Config/substitutions.txt": "80f82f0cf411b94c10368c3458703a6ad44e779778d3741d2008717f5692078a",
  "DonneesInitiales/Modeles/LETTRE TYPE.dot": "56c0ff3cb8167cb6746fda4ca772dee83d1e49a9480ead4f94017681c5dd55ed",
  "INSTALLATION_MULTI_POSTES.md": "1d8855ec78ccda2d315cf29cb5813d10c959b55c193d7f28655684d0242f309d",
  "INTEGRATION_UNIFIEE.md": "b94a76904d32c45d663428074997b3855794df76a16b389e008beabd03ffb4f0",
  "Installer.cmd": "8074b6ad7a2734a6c962b3641aeb2199cb611e1c9325c84aa3e8635576830efb",
  "Installer.ps1": "eb6dd013e1d1c624116613e19e3ff7a787890f16a0b07e3ff42e25fd8e121b8b",
  "ModelesSource/Cabinet(1).dotm": "493ed179d97ec4f15cafe39e44f9a6d2c0ce56a42927cf0fd4acafd860dabdb5",
  "ModelesSource/Cabinet.xlsm": "714e9458066bd28b8b6e6f6e9f65b2d673d3285c1915b57216e780d5e75f04ff",
  "ModelesSource/ModeleCourrierChatGPT_PROD(6).dotm": "a8ccb702e8defc9fce19b9afb968c4f7cd5b25d59da562dba172bc6b51b85524",
  "README.md": "5ef1e9623367195b334188990f801eff3374edacc815ac9fd631c61627cc4338",
  "RECETTE_WINDOWS.md": "8749183e7c9d0e57f019a3a075b239181feb64dc74b4ce4e10ce62ecba81e362",
  "Serveur/.dockerignore": "891c19e4a4cc9f2b65c507511effbb989345d165389833708b19671f3fc4b58f",
  "Serveur/.env.example": "54b6a40fe98c2a1c40d5a97159ef4ca819a7761128e1c46c7efa480ee166809b",
  "Serveur/Dockerfile": "529096dcb342f690289782313bc1a1d21e44d22edb9f69b0dd014d7bc1ba009d",
  "Serveur/INSTALLATION_NAS.md": "6014fee7d6f9bc59a52e9052daf27a7c44185060decec797df10d10f98627102",
  "Serveur/cabinet/__init__.py": "eb5a55e6e2b94787dba75aa3fbad0c5da564e583ab4bab302dbb5a04e8a4491a",
  "Serveur/cabinet/admin.py": "1b86b34db19d9b440f06e3d0b45163c639c3ca1ac6b38b16d8347c9f3c3c60b1",
  "Serveur/cabinet/api.py": "c1cfb55c87ad9dbd78df743fa0c458a90fda75ac9deb5231114587faaebb94a1",
  "Serveur/cabinet/domain.py": "0301a4015932b0af1bd1b04246cef41c118c90fafacc387bdd5203ec63cda2f5",
  "Serveur/cabinet/files.py": "3e4f0328c40d5d434664ad1f6d7d566a5146c522403217ec8005e54d237fee2c",
  "Serveur/cabinet/migration.py": "8ac8921d2a06f797de380dab684eda5ca361a0e3543b0f35ba3470e575f03119",
  "Serveur/cabinet/schema.sql": "7b45a91d184717a0da6288d0bbf80701f76e7f40b3796b13f1686e17eb47f737",
  "Serveur/cabinet/service.py": "fe30a579d35ab5787c5438fac489ab6f4187e2a1c666752642236d2247147f89",
  "Serveur/compose.yaml": "01f71ab3ad3995f18ec54cbe1f93b11437c2a578e5c4571f66fc696bfd8fc50d",
  "Serveur/init-db.sh": "4c6c9072a4931befa287d66fe063f8656631789bab7b89467cbb7fe4bf737288",
  "Serveur/preparer_secrets.py": "1276da746a1aaf7cdde53067bee4719d0e77ec72b8cfb39b5bbb3727dc9c1ab2",
  "Serveur/requirements.txt": "d07637652438bddf5e52ab1593c9a3f98d5c378b80d5c748c60efa8020413a8b",
  "Serveur/restaurer_nas_vide.sh": "e618a8e85f9c3925f7c434ec5407fb63c7947249d2ad927b367674aaeca0ddba",
  "Serveur/sauvegarder.sh": "914425e1661cfa4947c44a712e235a5c59702f1865bbb7ce5351b74bec9564e6",
  "Serveur/tests/conftest.py": "9128ca4d57dfb31ab63e0b8f1f27d9ebe9f0e6908f1a63030d63f67c736a93be",
  "Serveur/tests/test_domain.py": "a63c51b9214139daa60f5b2d5248fb85fae4b2cf130ec9c2ef2ad8c1f77393bb",
  "Serveur/tests/test_migration.py": "42535463e221ba0049a37241e93488e917e8240cf4b15114bfa7118ca0899064",
  "Serveur/tests/test_service.py": "a0a2a9f0ec1e9a990a9b70db9644f05d6725b46ec2993fe913e0df03752691c0",
  "Serveur/verifier_sauvegarde.sh": "cd26d9a182d4a3469730aaa307dd77897ba9d1eba1411817137040380f0f42fa",
  "Src/Commun/modConfig.bas": "73e8313d9142637a94b629a2d92d04d68fb77eef8d43f963a0f0127403057d19",
  "Src/Commun/modFichiers.bas": "7874f118e55e445942b67acbdd678f257657a9d8071a5226e980e84abebb2ca9",
  "Src/Commun/modGdt.bas": "24c4c7e3251343482e1e961b94008fdbc09d1d7730fbd307a7e252328e049e34",
  "Src/Commun/modJson.bas": "95a7c2ceb1c7d4041ffdc6b92e144d1a0066523d89a3836bd0c383bdf0317636",
  "Src/Commun/modLog.bas": "71df3a5a8df115a7757f32f4b7f979263e744597dc429500115711db53afb961",
  "Src/Commun/modServiceNas.bas": "0d1ac8803018e92ab26c8d1a31e0e7f5b6891743280c1c6d00541745df898121",
  "Src/Commun/modTexte.bas": "505d95a246a545e914b13e121f907d61b23d635a8fa95c24b308c1d4f43e3726",
  "Src/ConfigDefaut/config.ini": "690c45fd11a297c48cbf9ed992842000ef0794d9986b6e132ab2ad37e4904358",
  "Src/Excel/Feuil1.cls": "8b6d1fc4f872634510d7a8323acaeca8b89024c115aac0c3e7c5dcfd25331905",
  "Src/Excel/ThisWorkbook.cls": "8a39e77a2c2842d1eda5907a62a7f0cebfa380b1d35abdb0593bbda426bbadb1",
  "Src/Excel/modActes.bas": "f7e03c810c5cd63997497fee15ca10d98b45fcc7ad2d10321deae4fef94e70e0",
  "Src/Excel/modAgenda.bas": "ad6db082342afaa3becad334e94da6d3b414d607e9bd056ca63e7cf450cb2926",
  "Src/Excel/modAgendaVue.bas": "a10adeb8fb77e217b84b85a8e66495911ab234ea2d1763bd6b2285262a4e6caf",
  "Src/Excel/modBaseIO.bas": "e18573dbfdec4e7daf8e89da66175f12eec7ec82f8e410a06af044e0deeaeeea",
  "Src/Excel/modCerfaPrint.bas": "60e259b194cec30892333bf0fe0b5d93a54df352306bd59d9dd6263bc983c82c",
  "Src/Excel/modEchange.bas": "6ba940c4cda9506b98caab5abb4c73adb2838e2cc566c0c4d238f37c42383388",
  "Src/Excel/modJournal.bas": "f6a070602dbc4ae3eed137c330e51309ea3c505543db2c663ca0e1c4e257dd9d",
  "Src/Excel/modUI.bas": "196d2616e76db8923e6597f06f96020bebaff45a79fc5a9c5611d21cee90c08d",
  "Src/Excel/ufChoixActe.vba": "e674782549d343de5994462d07b2f0cbaee5c916b4b7124264bdcb7129a1df55",
  "Src/Excel/ufCorrespEdit.vba": "376b40235220f678a9fcad44318a393e846413fc1e96275f0d5f4c300c0ab170",
  "Src/Excel/ufListe.vba": "1d40b76607705d85decc2a7d25da79d224c6dba6171fdc0701396f5812c4ebd7",
  "Src/Excel/ufPatientEdit.vba": "8781bc6a97c74b9b422e2af04e529258b6916b5d8a7ce0d4404faf5a9df307e6",
  "Src/Excel/ufRdvEdit.vba": "70248dd9e8786a109786d0884e2d829638751522b8c5d998cf76dc10c347457f",
  "Src/Integration/modAttenteLocale.bas": "d151c3434d09bcb7c8b49109070ddfbd34165e27c6b8dd0ffea6554d50c6188f",
  "Src/Integration/modAuditTests.bas": "a87d4d6082bdaa5d231eb59f114d9b9da7971b9b1951733b8021f9c536bcc75a",
  "Src/Integration/modControleCourrier.bas": "64286eed3730b9fd98f1406b014298433eba102d762f28f9f8bda46d75ea87cf",
  "Src/Integration/modIntegrationUnifie.bas": "8f0834b6cca2a4e47742e8e480f94a78f8b79f44dfad155a36703e2d67db685b",
  "Src/Integration/modPowerMicUnifie.bas": "daf004b3d3d542c4f59e6c0aae93135297f77a3747be59ff08625574bcd9bf1d",
  "Src/Prod6/ThisDocument.cls": "6d7b7fd22869872b04bcc106514b36a55fb7bf8d9fa8eceb59a213b2770db4ba",
  "Src/Prod6/frmCTDestination.vba": "1d2f26456daa9f424139c153f5534779b30d212629e138f322f6b5275673019d",
  "Src/Prod6/modAnonymisation.bas": "2c0edc11aaa0bdd4b90f63f1807deb52521040f5af2a9d6a0c6a1f0855fcbe88",
  "Src/Prod6/modBaseCorrespondants.bas": "65770068f1af13f50007cdc36df8c9622996cbbfae3927bb582e61e5612e0ab8",
  "Src/Prod6/modCabinetTestDemandes.bas": "7b62a9a4add74ff13689067c702f3718ff1fae1ef8d8aa2c13802247e614a762",
  "Src/Prod6/modCabinetTestDestinations.bas": "3ac7897f2ddbc37cad35f611a91a976dc7d42f2ed780a16f7c1ca5608e48710b",
  "Src/Prod6/modCommandesDDERAC.bas": "a7672487958bf29026454d78f4e24a98399347828fbf63d6ffb2b95a7fcbddbb",
  "Src/Prod6/modDemandesAnnexes.bas": "769992c4d53166a3b98825054e1d03ce58ff2d0c74fe33c6ac67c24908d63eb8",
  "Src/Prod6/modDestinations.bas": "66f5181006ba957fa02b7226b40545889c0963c2b3bf02ccde9a9576bbb3aa9c",
  "Src/Prod6/modDetectionDemandesDico.bas": "a137a9aaad79fdbae45726d34355cc08648e38c3a7f710aeb2dd3084985bc82f",
  "Src/Prod6/modEnregistrementCorrespondants.bas": "535e57d7991a670f6a53bb3f2e96311abc46e71100d8a5ae219a069c64d1710b",
  "Src/Prod6/modGenerationDemandes.bas": "d426719dd9d36a3948e4b96e353aa7c89d66c10720dc58a64213e963b0778c49",
  "Src/Prod6/modGlobals.bas": "8c65a454418cf93fe126f67ba8c0758c2058ed7c6fc9ca818581d223102c8770",
  "Src/Prod6/modInsertion.bas": "e058f9cef6cb76222792679a80b23791cc97a072a84363f68016fde5087aa896",
  "Src/Prod6/modLettresComplementairesModele.bas": "91de776840af3b328d294ff697b5392644ecd1285de1944f48cc5b4edee2c259",
  "Src/Prod6/modMiseEnPageFinale.bas": "2d5993d3dd62a03f25035d46b6b70ef2ec3dcdb2f1a72ea1df6ece34c1f8fb0c",
  "Src/Prod6/modNormalisationGrasLocal.bas": "177713a5d9821032ed2718470da73e3caba51f393e365e253b69c642ef7b370f",
  "Src/Prod6/modOpenAI_v22_corrige.bas": "d82edd591923382e2e5d8ca90c38c4111f14d0ee2235dc9a792f6112d6c750f9",
  "Src/Prod6/modProdRapide.bas": "dc273ce4bedd936d465364d66ce6debb8bf77a99ba0cef78805690b174d640b3",
  "Src/Prod6/modPrompts.bas": "8bafef03425ab26a0eb6e46ef9eea3dff4cf529bda4ee3654d860d9712355288",
  "Src/Prod6/modSaisieCorrespondants.bas": "f1c644ee8f0b51b6fb12837974e3e5fab7c39d93163871879436a7ad9159372a",
  "Src/Prod6/modSaisieStructures.bas": "51aa1784c5066b39b646b48f831a3c069170a544ec8367b36d9ae3f8035fc0a7",
  "Src/Prod6/modSortieDragon.bas": "eeccbc81af7c56fba610eb7c86b658dbabc2e1f7aba7fe9b1310e69783a3aae1",
  "Src/Prod6/modUtilitaires.bas": "dc574f7a1fc7621fc1af61f7348f5a0c8ecd606c34c47d6d23acaf13102eac13",
  "Src/Word/modAnonymise.bas": "124207b9d0ab9d658c6fb6ea4067067f7ad938472a813eb3faa09035040f1a7e",
  "Src/Word/modApiConfiguration.bas": "4b9ff64a318592959b55703823116faf022d4dbb5e47b92738f0be5b93b96b9a",
  "Src/Word/modBase.bas": "c04ac90da7a8e5936ebdcbc62ac140aa441f0898e63cd4566eef7668b05f3184",
  "Src/Word/modClaude.bas": "ee5b2c971ad47e60d98e2741d50ec1ab052a18d1dab662b8910cc5c7154fc361",
  "Src/Word/modCourrier.bas": "912a1cd399d635cc5b382254ab8e89c74b883e16a5b11c6b4e58a419adc48a72",
  "Src/Word/modDemandes.bas": "c3ac084dabdd22670a00afce8e9018989dc501816f5144eed27148b3798c77ef",
  "Src/Word/modDerivees.bas": "0961570d433349b9b518afd6c063bf462b79b646a3fb381c03ec9c6fdb64cf0c",
  "Src/Word/modEcg.bas": "bda7e2dcdc3025cb98bfc55b6da2c87c75b3abcf85f710a113735053c9b29811",
  "Src/Word/modGras.bas": "90bbc5b2416ee212cb5d8f9fd6ed5a09548ab521995ce845fe4ac01c9014c6a9",
  "Src/Word/modPatient.bas": "01a8553d774e7df3cbbfd9d39b9ba106006b9d6ce0b5e041e8608bb0daf2d1b6",
  "Src/Word/modPowerMic.bas": "4f22aa1ec058f8663fdba691405ab887f9ff6c77666a59f5c61873fa5e572da3",
  "Src/Word/modRaccourcis.bas": "1e918155d53beee84eb657ec7131a2e3aeb9d6b95a61c9dbe7179e10e0ec64ad",
  "Src/Word/modRuban.bas": "39e62d77f674ef5e58b590b2244b42abd365c4e4f82733f4e0512d711de84e48",
  "Src/Word/modSchemas.bas": "45fb971cbf846a2af4993756e471fa1ef40fc7cd6f5f9ddc55a042f104938ed6",
  "Src/Word/modSubstitutions.bas": "a24ae57953c43e2c2bafd09e539d13badb46f5d8d417b37a162694b9fc94c302",
  "Src/Word/modValidation.bas": "f345d12214367bd59666cd960b782662182085333dbffb94414cff5a94dd748d",
  "Src/Word/ufListe.vba": "fd55bbfcc0c18450f1d5ff2739aaa5f266c2b52d40b7fbcee5358c520c683113",
  "Src/Word/ufProgression.vba": "49b383108d2e7c85f27b4f4aadd438c0251042e412cc36c06a60adb7a611f6b2",
  "Tests/audit_statique.py": "71989526985a724ad4f6c070a1497145c4bd0a6c49856411e8653ac9e9e6d17b",
  "Tests/inventaire_sources.json": "9f38d3321368eea89974991631095f006a8ef7921d7bf8148939a0d606418e4b",
  "Tests/nettoyage_sources.json": "f3b54768aca28e41878d9bd36a147d7f2f7da5815609e85576d2a03986c5e07d",
  "Tests/schema_reponse_api.json": "fbf5f9805faa23b77cdffe368bdd737e320665c58b75a528a268f69ac06665f4",
  "Tests/test_assistant.ps1": "11cc4bb1293777569a72537d23a203dbc70da60f191ae37fc9650b9b728bc2f6",
  "Tests/test_choix_nas.ps1": "54dd2afe5dde485758abdb0dc8064f00bf5e6a7b87689fe41e8deebd65201c0c",
  "Tests/test_construction.ps1": "4718fae561570b4d477c7a990cd24aec3cdd5535cc288d406cc2d545cd1b145e",
  "Tests/test_installation.ps1": "9b56f991c5949918038f04535259b800f8f293866f309fa9dbc37fcf12b44617",
  "Tests/test_sqlite.py": "b72e003154f16caf168fbd660a0e68c676196ec52eca3c496cefac603c11368d",
  "Tests/verification_livraison.json": "824c011534ea20fa09f6cf299ab28b7b7498fc67ce1fc158270e849992196486"
}
'@ | ConvertFrom-Json
$bootLock=$null
try {
    if ($env:OS -ne 'Windows_NT') { throw 'Ce lanceur necessite Windows.' }
    Write-Host 'Cabinet Cardio - installation guidee'
    Write-Host 'Utilisez votre session Windows habituelle, sans Executer en tant qu administrateur.'
    $cache=Join-Path $env:LOCALAPPDATA 'CabinetCardio\Installation\Sources'
    [void][IO.Directory]::CreateDirectory($cache)
    $bootLock=[IO.File]::Open((Join-Path $cache 'telechargement.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $package=Join-Path $cache $Commit
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
    & (Join-Path $package 'Build\assistant_installation.ps1')
    exit 0
} catch {
    Write-Host ''
    Write-Host ('Installation interrompue : '+$_.Exception.Message) -ForegroundColor Red
    Write-Host 'Relancez ce meme fichier apres correction. Le dossier prepare est conserve.'
    exit 1
} finally { if ($null -ne $bootLock) { $bootLock.Dispose() } }
