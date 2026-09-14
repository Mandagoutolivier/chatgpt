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

function Isoler-PaquetInvalide([string]$Dossier,$Empreintes) {
    if (-not (Test-Path -LiteralPath $Dossier)) { return '' }
    try { Tester-PaquetCabinet $Dossier $Empreintes;return '' }
    catch {
        $cause=$_.Exception.Message
        $archive=$Dossier+'.a-verifier-'+[guid]::NewGuid().ToString('N')
        [IO.Directory]::Move($Dossier,$archive)
        Write-Host ('Paquet incomplet ou modifie conserve dans : '+$archive)
        Write-Host ('Diagnostic : '+$cause)
        return $archive
    }
}

# Ce modele est assemble avec les outils et les empreintes par generer_lanceur.py.
# Aucun mot de passe ni jeton GitHub n est demande ou conserve.
$Commit='bb239763288c7a2024a2548e4cc69949a0fdbf9b'
$Empreintes=@'
{
  ".dockerignore": "30eae15fc1fa5b14ff10f03e58ed06cb18f519802455a73df93d0de4bf6d807f",
  "AUDIT.md": "d28f2a32618dea4ab44bf848879a848ab55e0788cc723346bc2570a2e35517a9",
  "Build/Inventorier_U0.ps1": "cc884b657b41c101987ec1749478d55d916d0da3557a4bdbc5cb22c73f5a7b2f",
  "Build/Tester_Droits_U0.ps1": "e8241e52f3a7b77b64dff584803e58dbac0c0a4c73247a34c1145b97bdbaef8d",
  "Build/Tester_U1_Office.ps1": "6ef0ece2a75ef6c475b1f4d60c80d67a3c876d531ff8224f05cd3de113d2f6d3",
  "Build/assistant_installation.ps1": "1949b59d2a460079bbf9758ae15f7ea8cdb8d2d47099a186736cd31d0bfd06ef",
  "Build/construire_cabinet_secretariat.ps1": "f5725a5192700ccf501cae854c0a1e3e7ea028301c8bdd78a57aa52401b9bcab",
  "Build/construire_modele_unifie.ps1": "48bf847188fccbd61ec81cca4bc168648f6cdac3775a9fb2729bfbee1827c238",
  "Build/donnees_initiales.json": "2c6a672f018d9db815c4f3fcdc91892db9b9a43bb3c7477bc83f49cb4b23b790",
  "Build/initialiser_nas.ps1": "4a3c3bec32126dfa6cf4b34f018fb275d370f08cd581f3f7d616d6a0123862e7",
  "Build/installer_multi_postes.ps1": "7aa9e4fcfbbd17ca8253256134c1f32534a0d6f87c917184bf3c8bc336af31e5",
  "Build/manifest.json": "e7e3084d3c04f3d7dd07125cc9ade54184aa293ed496adb4f05f89e536f21d9e",
  "Build/outils_assistant.ps1": "eb0ccaf8a1639d70c255beaeb15881e07aa1b3832bd795e1a6a4b302beb81f91",
  "Build/outils_construction.ps1": "f91c4159e5ff82d4599102773e8df380f2fa1870030ce893204f4bb9c90bcb37",
  "Build/outils_installation.ps1": "9c5c7204755db1d5cf554fe50ba0c6e55269f4daab0f26522a3c6412b8826d64",
  "Build/outils_recette_u1.ps1": "390f95f05efc59fed3a42183513653a81152c4da491e0dc7ab145a014c4eee6f",
  "Build/outils_telechargement.ps1": "2e6c3ce03391f6116276ec61a67625a3bae0c8a3f695acc065bc0a6f714ae3af",
  "Build/release.json": "1fc870de93069e67d603fe1dd682a68ac9c1b3329413a0eb93658e3cebaf3bda",
  "Build/restaurer_poste.ps1": "7fc5d7008e385a1ef669ae868c71b41498ac9172b417a6e007c91ef3302ef672",
  "Build/ruban_unifie.xml": "e3adc12dc7dbd3bb6041c5c76ceceeb7defdf86c6faf155ae4083957c548b858",
  "Build/schemas.json": "9b6c70a98f9129de98bf9aa1f1f6714b714e1610352801e0d806ab7254f3d0d4",
  "Build/valider_preparation.ps1": "a1171e4487193323675e2dca771a919d2c0ce6ae26edd9bd6e3112744bf30cb4",
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
  "INSTALLATION_MULTI_POSTES.md": "2bc373d79cf3096f296e5bc4cb79a4919b3658d2fdf7f0479e61491f52c0efa2",
  "INTEGRATION_UNIFIEE.md": "57198c83114597fc15ff24ca4857edc680dc6ff750a8ac8c7350d99287e9a64d",
  "Installer.cmd": "8074b6ad7a2734a6c962b3641aeb2199cb611e1c9325c84aa3e8635576830efb",
  "Installer.ps1": "a095debedfe575ad21b21b8786021cae70f9887d4ebbd4dd66e73cca7f948926",
  "ModelesSource/Cabinet(1).dotm": "493ed179d97ec4f15cafe39e44f9a6d2c0ce56a42927cf0fd4acafd860dabdb5",
  "ModelesSource/Cabinet.xlsm": "714e9458066bd28b8b6e6f6e9f65b2d673d3285c1915b57216e780d5e75f04ff",
  "ModelesSource/ModeleCourrierChatGPT_PROD(6).dotm": "a8ccb702e8defc9fce19b9afb968c4f7cd5b25d59da562dba172bc6b51b85524",
  "README.md": "fd8ac89f4ba1390dff85fe9d19c6c3c5b3c88e74f1a64ce24dc42f9305b37f91",
  "RECETTE_WINDOWS.md": "8749183e7c9d0e57f019a3a075b239181feb64dc74b4ce4e10ce62ecba81e362",
  "Serveur/.dockerignore": "891c19e4a4cc9f2b65c507511effbb989345d165389833708b19671f3fc4b58f",
  "Serveur/.env.example": "04908d383a89fcbf08243d22a3bacc7f730ba5900171e9cf47fe0e28e91513de",
  "Serveur/Dockerfile": "1a8851fe7f1f529f5cbde88afc0009b96cc96c80d9703849c531449a5923ddd2",
  "Serveur/Dockerfile.maintenance": "21a7fa12e209588c2e3175da8c3a5ed388a83ce322f78e64e17289d59c5f8b75",
  "Serveur/INSTALLATION_NAS.md": "cab4710b11cf2addedec71718e52079bfbb5357611817cb331b763edde93a693",
  "Serveur/cabinet/__init__.py": "eb5a55e6e2b94787dba75aa3fbad0c5da564e583ab4bab302dbb5a04e8a4491a",
  "Serveur/cabinet/actes.py": "312c08e7ad7f9143ff32a98bb04258aee23f794f6f6bb2853c375693981d6979",
  "Serveur/cabinet/admin.py": "1b86b34db19d9b440f06e3d0b45163c639c3ca1ac6b38b16d8347c9f3c3c60b1",
  "Serveur/cabinet/api.py": "6506b501c326d4cbdcc26221c8b65c9484a3e1028f97f236e7d914fb8f34f773",
  "Serveur/cabinet/contract.py": "8bbd4dbcea3920323e70754776e909cc3f7567409876f769a69dd7df058913b7",
  "Serveur/cabinet/domain.py": "336e77f0c0d25c7203cab332ec5945c8998846ad7e61fe23c343fd334afff1e5",
  "Serveur/cabinet/files.py": "bb29eee213d2fd3f9f430ea76a81805072fe17c0ec8786b6a221e1768f9ad639",
  "Serveur/cabinet/migration.py": "509e10d3e1abb1c58568a1b0478893d403c2693ce5c2963a4bda2cb1f05ccbd1",
  "Serveur/cabinet/migration_u1.py": "9611fa58400a722800c945ed97a75a4fb354eba36b7282ee02acbcca486fa810",
  "Serveur/cabinet/recovery.py": "e9f6d4d41ba183dfe79ae7a56c24776936fb78dc9088073cfd50852fc2f8684e",
  "Serveur/cabinet/schema.sql": "5b2fb144318ca3bedb38a39ec9dd000c105474a9a295e4547755d2fed7eeadab",
  "Serveur/cabinet/service.py": "809744c3a8cd238621275b258a62d987354d88850fc50b14a46eeb830cfdc834",
  "Serveur/compose.recette.yaml": "209d2ae6de3741fc156d4eee53daea08d0a2049b1ca9eb66286614b36e0bf261",
  "Serveur/compose.restauration.yaml": "a0486945db508cd7b9504fd75f4fd3af5314e0bb930090e820a88a9cb609362b",
  "Serveur/compose.verification.yaml": "88c1503ec917a203e3daf1d9239e6b6bfed2fa7c43cb6768ac4fe2d67f49f20e",
  "Serveur/compose.yaml": "5cd7f45eee1acd8fc5d7b95fc3c29f9aa48adf779031ad7eb452e05542a4e28e",
  "Serveur/init-db.sh": "4c6c9072a4931befa287d66fe063f8656631789bab7b89467cbb7fe4bf737288",
  "Serveur/preparer_secrets.py": "1276da746a1aaf7cdde53067bee4719d0e77ec72b8cfb39b5bbb3727dc9c1ab2",
  "Serveur/requirements.txt": "d07637652438bddf5e52ab1593c9a3f98d5c378b80d5c748c60efa8020413a8b",
  "Serveur/restaurer_nas_vide.sh": "11acfef9dfb71400deba9c9f8616e47ea4d0100f15f2005ce8f0b7370e0926a3",
  "Serveur/sauvegarder.sh": "45e3d0d407b23f2ab37b2469567f169144b85db42f58c710b1f4b82fa206e86a",
  "Serveur/tests/conftest.py": "9128ca4d57dfb31ab63e0b8f1f27d9ebe9f0e6908f1a63030d63f67c736a93be",
  "Serveur/tests/recovery_smoke.py": "08624bbbc70b05efe57832ecf8536b904703499d3898c852f61f139564288ea9",
  "Serveur/tests/test_domain.py": "3b12f0003c4ace100971d7b4a4caf983a545b0d6f6ee761a66b79c9000b8dedd",
  "Serveur/tests/test_migration.py": "42535463e221ba0049a37241e93488e917e8240cf4b15114bfa7118ca0899064",
  "Serveur/tests/test_service.py": "59f951e5c2789030affeb0a56713e111184b0cacb02f9e34d5a690a0e85757b0",
  "Serveur/tests/test_u0.py": "4c83d22f81fab09aefdf5ec28cb0d11b1e3d140c5129dcf7dfb449822ab43e07",
  "Serveur/tests/test_u1.py": "f4b795f54ba3ef959f9eaecbe08a6ff9580f07297af982f0306fe70a8b92ebe7",
  "Serveur/verifier_sauvegarde.sh": "65d1add476f28a25be3a1ff9e30073c49c999eaeadbe18ed68e84493815a4506",
  "Src/Commun/modConfig.bas": "2b4417a29f6f77c2a8e350141e7d15be1919e5829406ddf906c5c3b8c26524c7",
  "Src/Commun/modFichiers.bas": "7874f118e55e445942b67acbdd678f257657a9d8071a5226e980e84abebb2ca9",
  "Src/Commun/modGdt.bas": "24c4c7e3251343482e1e961b94008fdbc09d1d7730fbd307a7e252328e049e34",
  "Src/Commun/modJson.bas": "95a7c2ceb1c7d4041ffdc6b92e144d1a0066523d89a3836bd0c383bdf0317636",
  "Src/Commun/modLog.bas": "487da3c9f9fa8266aef430a63f92422787a28f0bc1d12ef7cfd0c9cb94e2faf8",
  "Src/Commun/modServiceNas.bas": "ae6d2ab9dcd7fa01151e8e72ccd3c9e78ee327798e5692bc13c023f3c17cd0d8",
  "Src/Commun/modTexte.bas": "cf753d80d6ee8d5b009ac35a05e42a6c9fb1b6cae24358249a7d4061d97f9fc7",
  "Src/ConfigDefaut/config.ini": "690c45fd11a297c48cbf9ed992842000ef0794d9986b6e132ab2ad37e4904358",
  "Src/Excel/Feuil1.cls": "8b6d1fc4f872634510d7a8323acaeca8b89024c115aac0c3e7c5dcfd25331905",
  "Src/Excel/FeuilAgenda.cls": "52f6239f0b04aea0e6b65e6cb30459592791d31b99ee490bb75ffe067d10fd63",
  "Src/Excel/ThisWorkbook.cls": "9f3a8a8c43e263590484651c2dea8863fd61e7032cc27a100e3a80cb14612717",
  "Src/Excel/modActes.bas": "4eb2de602cb1b48fa86f9999033e676712f16561c8b6d1e5d3a74a5db5c100f3",
  "Src/Excel/modAgenda.bas": "c9a4208489534fb7339b2986fa067f22d863e0569f1ebd7fca828488be48d8bf",
  "Src/Excel/modAgendaVue.bas": "c2517971cd4726c75c47725ff3f50aefc7f0bbbc631ce69f6d2b6ac7567197da",
  "Src/Excel/modBaseIO.bas": "bcfb87f0d09b5e62127385030a16fd627d1ac49d9221f32fc73e35aefd5c172a",
  "Src/Excel/modCerfaPrint.bas": "f490a8b96223751ec601d8e2191af2445ccdbe2b79369b66cf35ba4f5c2eb291",
  "Src/Excel/modEchange.bas": "fbcfa55845fdaacf87330f1c6e07b8108e6b75b2d4f9f00462b4f3932a2cd318",
  "Src/Excel/modJournal.bas": "8c9f0edb62936cef2aff6ddeb2edbbf465f49b0154ae7dec7f768bbcace32129",
  "Src/Excel/modRecetteU1Excel.bas": "0ee18a0ce94f9d93379129c377679084d6698a5bf7cf67e4f3c026978921d8d6",
  "Src/Excel/modUI.bas": "34cbddaffe84d129c70a3963e73fb3d2b0c34b17f51cdd4e5ed16a2009357906",
  "Src/Excel/ufChoixActe.vba": "78034d6f1452d6e090b6bfff06bfb1ac28e485c618cb719864eb7e526fd86114",
  "Src/Excel/ufCorrespEdit.vba": "376b40235220f678a9fcad44318a393e846413fc1e96275f0d5f4c300c0ab170",
  "Src/Excel/ufListe.vba": "1d40b76607705d85decc2a7d25da79d224c6dba6171fdc0701396f5812c4ebd7",
  "Src/Excel/ufPatientEdit.vba": "8781bc6a97c74b9b422e2af04e529258b6916b5d8a7ce0d4404faf5a9df307e6",
  "Src/Excel/ufRdvEdit.vba": "b77490776a98c3caf72ef8eacd20d2cb94e2799a4b4b45c5e8eef88fed79655d",
  "Src/Integration/modAttenteLocale.bas": "1069e93b6e2b01993d5981674e09d6a16b7ff26a8271a9e4b372b2a056446254",
  "Src/Integration/modAuditTests.bas": "a87d4d6082bdaa5d231eb59f114d9b9da7971b9b1951733b8021f9c536bcc75a",
  "Src/Integration/modControleCourrier.bas": "7c317c4c792bed9085b68237cdec7a9fc8a014973c4cd64a96da9816b79e2bd5",
  "Src/Integration/modEtatCourrier.bas": "1f50c4dc21dd2a4d868462edb7dc01e41b4232984e6423b8c3d30993b2794005",
  "Src/Integration/modIntegrationUnifie.bas": "55c401062d1a85e79a57932178773b69394f814080704ae9bbabd6129d678d98",
  "Src/Integration/modPowerMicUnifie.bas": "3948b45445aaccc4896508a1f50e99dcd04b6da9e30f3ee95b029a52b61c486d",
  "Src/Integration/modRecetteU0.bas": "75fe44170262584fc150b0417451e0d87d2b4d9b4cfa7a98f4b738a57f56f1a5",
  "Src/Integration/modRecetteU1.bas": "b95d2117183781dd201e060d05f19d66a959d5133d35cfab6679d0c6d7d10922",
  "Src/Integration/modRechercheWord.bas": "63137fbff9b20f04070e95829f8ce032be5c14190c2b4eea0418c7e55b3ef4fa",
  "Src/Prod6/ThisDocument.cls": "6d7b7fd22869872b04bcc106514b36a55fb7bf8d9fa8eceb59a213b2770db4ba",
  "Src/Prod6/frmCTDestination.vba": "1d2f26456daa9f424139c153f5534779b30d212629e138f322f6b5275673019d",
  "Src/Prod6/modAnonymisation.bas": "2c0edc11aaa0bdd4b90f63f1807deb52521040f5af2a9d6a0c6a1f0855fcbe88",
  "Src/Prod6/modBaseCorrespondants.bas": "ae5cfa3585806742364fd5d1b0174e16cf114cdab8b701163c2bf146be0fe2f0",
  "Src/Prod6/modCabinetTestDemandes.bas": "7b62a9a4add74ff13689067c702f3718ff1fae1ef8d8aa2c13802247e614a762",
  "Src/Prod6/modCabinetTestDestinations.bas": "3ac7897f2ddbc37cad35f611a91a976dc7d42f2ed780a16f7c1ca5608e48710b",
  "Src/Prod6/modCommandesDDERAC.bas": "8fdc6842aabcfcbc7fa3a7065642a9c59ce7eaab5640a2e29184e56a9e61db68",
  "Src/Prod6/modDemandesAnnexes.bas": "b4b73917a6ac8a41d7c193d7cf5a802878c52cff5f3a96a4ed7d5cc5c57df4f8",
  "Src/Prod6/modDestinations.bas": "386ff0c13833096e3304ed1fc43eb50447d638783168696a4aea470b00fc6b00",
  "Src/Prod6/modDetectionDemandesDico.bas": "a137a9aaad79fdbae45726d34355cc08648e38c3a7f710aeb2dd3084985bc82f",
  "Src/Prod6/modEnregistrementCorrespondants.bas": "535e57d7991a670f6a53bb3f2e96311abc46e71100d8a5ae219a069c64d1710b",
  "Src/Prod6/modGenerationDemandes.bas": "d9d6141c130f175e1b21d9e10ae001cef25f88d3856d014208e6e1b99f2113a4",
  "Src/Prod6/modGlobals.bas": "8c65a454418cf93fe126f67ba8c0758c2058ed7c6fc9ca818581d223102c8770",
  "Src/Prod6/modInsertion.bas": "4873abdf727ed1059413a8129aeacea01231a413a809cbe78f8c227164156d66",
  "Src/Prod6/modLettresComplementairesModele.bas": "57186d7b8da7eecfda55763e0bac2a5b323f56031129184f0ab7aad594c13a88",
  "Src/Prod6/modMiseEnPageFinale.bas": "2d5993d3dd62a03f25035d46b6b70ef2ec3dcdb2f1a72ea1df6ece34c1f8fb0c",
  "Src/Prod6/modNormalisationGrasLocal.bas": "177713a5d9821032ed2718470da73e3caba51f393e365e253b69c642ef7b370f",
  "Src/Prod6/modOpenAI_v22_corrige.bas": "338a1956c3429ae8cc0581894f062bbb8bf38713823446b7c0a3f31a8837153a",
  "Src/Prod6/modProdRapide.bas": "27ea39acd0e34863fd90f12ad99573672499ea034b59c8384b78cb66b26a13fe",
  "Src/Prod6/modPrompts.bas": "01ca059b62885151338a8a73dd7f4f1ad744a9cf0664d7d6e4654ad729f5791c",
  "Src/Prod6/modSaisieCorrespondants.bas": "f1c644ee8f0b51b6fb12837974e3e5fab7c39d93163871879436a7ad9159372a",
  "Src/Prod6/modSaisieStructures.bas": "51aa1784c5066b39b646b48f831a3c069170a544ec8367b36d9ae3f8035fc0a7",
  "Src/Prod6/modSortieDragon.bas": "48c2770cf627204c7dbcefc64376413855201dc2e818a86ed3d3bd7e6fec92a3",
  "Src/Prod6/modUtilitaires.bas": "dc574f7a1fc7621fc1af61f7348f5a0c8ecd606c34c47d6d23acaf13102eac13",
  "Src/Word/modAnonymise.bas": "b40022cf9f3f3752a7277c67119863dce3f954ec16a7d8131c886326e594567d",
  "Src/Word/modApiConfiguration.bas": "4b9ff64a318592959b55703823116faf022d4dbb5e47b92738f0be5b93b96b9a",
  "Src/Word/modBase.bas": "c04ac90da7a8e5936ebdcbc62ac140aa441f0898e63cd4566eef7668b05f3184",
  "Src/Word/modClaude.bas": "ee5b2c971ad47e60d98e2741d50ec1ab052a18d1dab662b8910cc5c7154fc361",
  "Src/Word/modCourrier.bas": "35f26a15b105e9595d5d9ad56a651491ef062a6573cced8a811d0d73f4ae40ca",
  "Src/Word/modDemandes.bas": "c3ac084dabdd22670a00afce8e9018989dc501816f5144eed27148b3798c77ef",
  "Src/Word/modDerivees.bas": "0961570d433349b9b518afd6c063bf462b79b646a3fb381c03ec9c6fdb64cf0c",
  "Src/Word/modEcg.bas": "bda7e2dcdc3025cb98bfc55b6da2c87c75b3abcf85f710a113735053c9b29811",
  "Src/Word/modGras.bas": "18f266fa40ef6491292ea4a36d7212c2fc580323e6bc8c8da434c6412232d351",
  "Src/Word/modPatient.bas": "01a8553d774e7df3cbbfd9d39b9ba106006b9d6ce0b5e041e8608bb0daf2d1b6",
  "Src/Word/modPowerMic.bas": "4f22aa1ec058f8663fdba691405ab887f9ff6c77666a59f5c61873fa5e572da3",
  "Src/Word/modRaccourcis.bas": "7a375fdcaed4e0351254c8a5bc8b423d3fc89e33d6090fb63975c3f26e23dc1a",
  "Src/Word/modRuban.bas": "d1ff7bf10a4d6c0a7cb2c9fa56f4e828763f475ca0a69c135e940aa072d8abb4",
  "Src/Word/modSchemas.bas": "45fb971cbf846a2af4993756e471fa1ef40fc7cd6f5f9ddc55a042f104938ed6",
  "Src/Word/modSubstitutions.bas": "a24ae57953c43e2c2bafd09e539d13badb46f5d8d417b37a162694b9fc94c302",
  "Src/Word/modValidation.bas": "c3cf2412edc314995d706e61c5a36f94bc955535730e30558a93ae8373aa9d4a",
  "Src/Word/ufListe.vba": "fd55bbfcc0c18450f1d5ff2739aaa5f266c2b52d40b7fbcee5358c520c683113",
  "Src/Word/ufProgression.vba": "49b383108d2e7c85f27b4f4aadd438c0251042e412cc36c06a60adb7a611f6b2",
  "Tests/audit_statique.py": "71989526985a724ad4f6c070a1497145c4bd0a6c49856411e8653ac9e9e6d17b",
  "Tests/inventaire_sources.json": "137a8eb19984d1ac0a1f4113bfffaf2e5f983d07c5c7153af9f95919d64ec539",
  "Tests/nettoyage_sources.json": "f3b54768aca28e41878d9bd36a147d7f2f7da5815609e85576d2a03986c5e07d",
  "Tests/schema_reponse_api.json": "fbf5f9805faa23b77cdffe368bdd737e320665c58b75a528a268f69ac06665f4",
  "Tests/test_assistant.ps1": "11cc4bb1293777569a72537d23a203dbc70da60f191ae37fc9650b9b728bc2f6",
  "Tests/test_attente_office.ps1": "1859e958edd9d74ce5b3d27d8d8faa976bf5576e5138d9a3fa6ac608f53b7677",
  "Tests/test_choix_nas.ps1": "54dd2afe5dde485758abdb0dc8064f00bf5e6a7b87689fe41e8deebd65201c0c",
  "Tests/test_compilation_office.ps1": "dfb927de3bf82ba3d31a304b17f944fdb0ae3d73b589ab2f7002aadcf0b2de9b",
  "Tests/test_construction.ps1": "ea6f3384499cb1c2ec48c2bf0f66145ef736c95ec0e65ef84057f2c1b835e46c",
  "Tests/test_feuilles_excel.ps1": "4d552b64e4a44913daa035af1ee346e2c7d9ac3f1b1d8c9218f1eef82a21c48a",
  "Tests/test_image_permissions.sh": "ae34c465ef05c08667703e9bc2609f05112e862c2115859edd03ca80ed8b5c2d",
  "Tests/test_installation.ps1": "923bff34cbe62f22af7bd188fff5123d227fc932d09a886ab2865824b3f2b682",
  "Tests/test_lancement_direct.ps1": "44dd1fc6b61d15d73c779b298764470954a77f51a6f158ec9282c77bbdf71cc5",
  "Tests/test_recovery_docker.sh": "7250d30f5db6d05b15dabe859bcdc21d93197e10d0184ffea5ad29c7d8c99a37",
  "Tests/test_u0_installation.ps1": "5f87ef07bd91f8f9b5aa5698ede582059938590438338001f8a900074281a8c1",
  "Tests/test_u1_installation.ps1": "238228e934846fccd9a63a61020c488fbd292ce386920cc40da7242ccca6023f",
  "Tests/verification_livraison.json": "e1733c2ce5b5523783ef9d8385971ba768499d6acb7e411b98179576cf9cbb50",
  "U0_RECETTE.md": "30208817c337df7b89349d71ecbf80bfe77b9f456ebde5e184aef27ef1531987",
  "U1_RECETTE.md": "faec11f24bc95db3e35a2d690483ba76b9d1cd8c18c30f136ad837962a86d959"
}
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
