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
$Commit='3fc341094a7895beb6fac0f6dfcc420b6a9b8755'
$Empreintes=@'
{
  ".dockerignore": "30eae15fc1fa5b14ff10f03e58ed06cb18f519802455a73df93d0de4bf6d807f",
  "AUDIT.md": "d28f2a32618dea4ab44bf848879a848ab55e0788cc723346bc2570a2e35517a9",
  "Archives/README.md": "08b636eeba18ee9bfb6b21dda06a5f4c13a3a5d1fa353fbba7af3d30e147a8d5",
  "Archives/Vba/modCabinetTestDemandes.bas": "7b62a9a4add74ff13689067c702f3718ff1fae1ef8d8aa2c13802247e614a762",
  "Archives/Vba/modCabinetTestDestinations.bas": "3ac7897f2ddbc37cad35f611a91a976dc7d42f2ed780a16f7c1ca5608e48710b",
  "Archives/Vba/modClaude.bas": "ee5b2c971ad47e60d98e2741d50ec1ab052a18d1dab662b8910cc5c7154fc361",
  "Build/Inventorier_U0.ps1": "cc884b657b41c101987ec1749478d55d916d0da3557a4bdbc5cb22c73f5a7b2f",
  "Build/QUALIFICATION_OFFICE_ISOLEE.md": "4ab4b0709c02c54c99e1a940c6229595a7604f0a5f340c7ec4203f909bad6551",
  "Build/Tester_Droits_U0.ps1": "e8241e52f3a7b77b64dff584803e58dbac0c0a4c73247a34c1145b97bdbaef8d",
  "Build/Tester_U1_Office.ps1": "ce797b61045bd6ef6b71096b71bc0b02d1f5928b749df0fae2f1e11552ee87dd",
  "Build/Tester_U2_Office.ps1": "9996d8d46851af5f7dc2dbf3c287f6890abda6c00aa12338a8af6e5955524eac",
  "Build/Tester_U2_Office_Isole.ps1": "5aa03945436451dc43b676ea064030e79017adb9473c35d25aade78f9138f694",
  "Build/assistant_installation.ps1": "feceb2df423106defe1ad0f4658bd291e63d6a4d21fc409af35654ee94fc7933",
  "Build/construire_cabinet_secretariat.ps1": "fc0040304f510ba664a26bc8493b8f875d601a9b4d7ab2977fcadf8b6c58e811",
  "Build/construire_modele_unifie.ps1": "e3fc00b5fd4671979dbfbb6c29f3d6b1648b5ca9b5399e80f25026181e88483f",
  "Build/donnees_initiales.json": "2c6a672f018d9db815c4f3fcdc91892db9b9a43bb3c7477bc83f49cb4b23b790",
  "Build/initialiser_nas.ps1": "fcc628697d4b3881e2a7a13cb67556e8479221af4b935c1fb1576c6c9a4e1ce9",
  "Build/installer_multi_postes.ps1": "9d0ebf35fbfceaa9ce0f06260d45e95192ecc024dfe716f174128f3849a887d6",
  "Build/manifest.json": "6727bae24c0829e835f28021f2658e3e796c17195908603a8fbe182f56bee55b",
  "Build/outils_assistant.ps1": "eb0ccaf8a1639d70c255beaeb15881e07aa1b3832bd795e1a6a4b302beb81f91",
  "Build/outils_configuration.ps1": "d7a6e3bbbc967d51130b16834fa3ea26fa4b30f4ade6da9437b509f89e09176f",
  "Build/outils_construction.ps1": "cf5c18869b03c5caf21e9ffa01b17e6c60889efd71f9b9c14ebef41fe423fc17",
  "Build/outils_installation.ps1": "51d362da933bb1407ae77a7684bc72e5f8e667e29e40af241b27fe34d4d7c38a",
  "Build/outils_recette_isolee.ps1": "b45ee83307be8794594ebb4ab410dc7b87abc95d2c74a6e0c86c428115c1e5f4",
  "Build/outils_recette_u1.ps1": "903df29f64a71efe8d51eb950b38b2924c09a4d693323f5bc99698ebabdd2959",
  "Build/outils_telechargement.ps1": "2e6c3ce03391f6116276ec61a67625a3bae0c8a3f695acc065bc0a6f714ae3af",
  "Build/release.json": "b08dbeb43cac54a84e05716616768b8460b2139d2f6ef795f132ffb59a6acd10",
  "Build/restaurer_poste.ps1": "5cf5a5bbc0b734fbb5dda346f56957e23df2d624f40f88ecb0432ff033d3c874",
  "Build/ruban_unifie.xml": "e3adc12dc7dbd3bb6041c5c76ceceeb7defdf86c6faf155ae4083957c548b858",
  "Build/schemas.json": "9b6c70a98f9129de98bf9aa1f1f6714b714e1610352801e0d806ab7254f3d0d4",
  "Build/valider_preparation.ps1": "5fc10bf3559c354c8ad373183879af4a52d40f3caf7b00b3e629453cfd64c2af",
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
  "INSTALLATION_MULTI_POSTES.md": "2295885a45e99f42c28420456f8959c8409b98c8e39f8f6f58ae80d2f73acf65",
  "INTEGRATION_UNIFIEE.md": "57198c83114597fc15ff24ca4857edc680dc6ff750a8ac8c7350d99287e9a64d",
  "Installer.cmd": "8074b6ad7a2734a6c962b3641aeb2199cb611e1c9325c84aa3e8635576830efb",
  "Installer.ps1": "4c283264e30752b0093ab59e75bde31b721a8e2829df00f1b1ebea47f57c8c8a",
  "ModelesSource/Cabinet(1).dotm": "493ed179d97ec4f15cafe39e44f9a6d2c0ce56a42927cf0fd4acafd860dabdb5",
  "ModelesSource/Cabinet.xlsm": "714e9458066bd28b8b6e6f6e9f65b2d673d3285c1915b57216e780d5e75f04ff",
  "ModelesSource/ModeleCourrierChatGPT_PROD(6).dotm": "a8ccb702e8defc9fce19b9afb968c4f7cd5b25d59da562dba172bc6b51b85524",
  "README.md": "93c7cf59eb7ade6c691a1cba9fbdb93449057d92864f26e6891d822536057f6e",
  "RECETTE_WINDOWS.md": "a157287ea0028c4d85089430ce3d984560f9ecca2f82a644725b9a40aab05016",
  "Serveur/.dockerignore": "891c19e4a4cc9f2b65c507511effbb989345d165389833708b19671f3fc4b58f",
  "Serveur/.env.example": "15e5c5794c9e58cd13075352a65d4d147c89328efb839612aefc63e6766cd0df",
  "Serveur/.env.u2-test.example": "3ab3b24b3b84b3499f8d6cf3dba8dbd0dc249da01df3e88ec607574a55871a1e",
  "Serveur/Dockerfile": "446879ab581b6fb73de6b84228d8c7f81e4367515f0dca39e2bac464917bba3a",
  "Serveur/Dockerfile.maintenance": "21a7fa12e209588c2e3175da8c3a5ed388a83ce322f78e64e17289d59c5f8b75",
  "Serveur/INSTALLATION_NAS.md": "3e028e07ea94950c8285538a868349dce6c86c94692e7c44ed10acca596d69fd",
  "Serveur/MAINTENANCE_U2.md": "52faf2c13d2d898f36a8bfd4bf22ea3f1a437535b6b894c2a6e03b84dfb368e8",
  "Serveur/cabinet/__init__.py": "8b8a39b7cdfce56a5338ba7d0ddf259764cc30e37b3f169129ac008b8373ae6a",
  "Serveur/cabinet/actes.py": "312c08e7ad7f9143ff32a98bb04258aee23f794f6f6bb2853c375693981d6979",
  "Serveur/cabinet/admin.py": "7241026b515e629d552610285a6c1d2943ac5405830a06b6a58cb1de86a4dce2",
  "Serveur/cabinet/api.py": "9b7e1c7f966d67d7f827807e90921fcfd206910cba1c918997b828da8da5e94b",
  "Serveur/cabinet/contract.py": "6b07e2277cd85819ea0a0ee6e943cc97cc5cb9191acfc99c73c218917a9151d5",
  "Serveur/cabinet/correspondants.py": "d4bfbe42a3f77e9f3d14896ef5643e28f10531f645079bc87a45b0168460d77f",
  "Serveur/cabinet/domain.py": "336e77f0c0d25c7203cab332ec5945c8998846ad7e61fe23c343fd334afff1e5",
  "Serveur/cabinet/editing.py": "d62462f7ec776ebfdb05dfaaf40dc4d42f2df8f23049e11f1f6f87e2299965c2",
  "Serveur/cabinet/files.py": "efb53105d755fbe1e057a5acb95531a705df9d09b36e167271c174850f9e1dd6",
  "Serveur/cabinet/maintenance.py": "a80b9013c235ec1a4699834b9011c81c8a5dfac632f1e8c3511ad7198705fc2d",
  "Serveur/cabinet/migration.py": "aeb873232da1d531c3c08ad37d03de502465f4b92d47d9c7927841c1264e32c9",
  "Serveur/cabinet/migration_u1.py": "9611fa58400a722800c945ed97a75a4fb354eba36b7282ee02acbcca486fa810",
  "Serveur/cabinet/recovery.py": "d9b65b5bd4015cf6ce4b06b3968260249598d4a9fd695291d9959f4d81c185c6",
  "Serveur/cabinet/schema.sql": "5b2fb144318ca3bedb38a39ec9dd000c105474a9a295e4547755d2fed7eeadab",
  "Serveur/cabinet/service.py": "38ac38c5e13f0c1e876caaf12596758155c9dd8b3e5de120f8fffc32977ce726",
  "Serveur/compose.recette.yaml": "6cd3f3475549f972b00beec9f1d49118859b25941752339920193f3b7168bc4e",
  "Serveur/compose.restauration.yaml": "a0486945db508cd7b9504fd75f4fd3af5314e0bb930090e820a88a9cb609362b",
  "Serveur/compose.verification.yaml": "803afca9617b54a74c3f37ef275645ae486f8543859c9477cab63778932ddb07",
  "Serveur/compose.yaml": "189c22156fe94b2d9818d8d5b351f0801aa3461b42136d802ea66c03d24d99f8",
  "Serveur/init-db.sh": "4c6c9072a4931befa287d66fe063f8656631789bab7b89467cbb7fe4bf737288",
  "Serveur/preparer_secrets.py": "1276da746a1aaf7cdde53067bee4719d0e77ec72b8cfb39b5bbb3727dc9c1ab2",
  "Serveur/requirements-runtime.txt": "c1d43f6c2602e33b8d747125bb0a8e1b111d4d7fbe0a8d9e41358cfc2851d8b6",
  "Serveur/requirements-test.txt": "415b95fb11523391a6556e17aefb37a27cefe6d91d8def1dd4d4fb34e795f338",
  "Serveur/requirements.txt": "ec641d77d9fba9baa341266e8379a125c23a47bf3cd828dc39d3e0b6201dc84c",
  "Serveur/restaurer_nas_vide.sh": "11acfef9dfb71400deba9c9f8616e47ea4d0100f15f2005ce8f0b7370e0926a3",
  "Serveur/sauvegarder.sh": "45e3d0d407b23f2ab37b2469567f169144b85db42f58c710b1f4b82fa206e86a",
  "Serveur/tests/conftest.py": "9128ca4d57dfb31ab63e0b8f1f27d9ebe9f0e6908f1a63030d63f67c736a93be",
  "Serveur/tests/recovery_smoke.py": "2590f63cc6087cab2956d44cd5875bfe99ca97d389e2f3bd423b518f49bdbbe0",
  "Serveur/tests/test_audits_u2.py": "08aa002ae11efc6c9b2cd0b6c9c2ac51c94be1a463c359a84e90b835b4d2521f",
  "Serveur/tests/test_correspondants_statuts.py": "41b9dbece17aaa2695ce172156cbdbf76f5a0facc8e214ea7ca6fda9fa205567",
  "Serveur/tests/test_destinataire_avant_ia.py": "dcaf0701422fef99f0212617bf506ae0807ba169661e118ce92118b8e3fe3a7c",
  "Serveur/tests/test_domain.py": "3b12f0003c4ace100971d7b4a4caf983a545b0d6f6ee761a66b79c9000b8dedd",
  "Serveur/tests/test_migration.py": "c8f49b032f0c311a52b669dd4d5f12137539dd3936a0d509f7e3d7f63dcc77e7",
  "Serveur/tests/test_service.py": "ef323f66a79d53d51eed95e60992a91c8c3e90d98372727bddc7c0d463aaaafa",
  "Serveur/tests/test_u0.py": "4c83d22f81fab09aefdf5ec28cb0d11b1e3d140c5129dcf7dfb449822ab43e07",
  "Serveur/tests/test_u1.py": "eeec49827b59619e7cc0938067265b6bda37154d458e4cf4d7a49ec07f5a11df",
  "Serveur/tests/test_u2.py": "8cf447c3da5049ad3ccdcb4a528f0550f3e817029a06fd1d8b4bf1886924c948",
  "Serveur/verifier_sauvegarde.sh": "65d1add476f28a25be3a1ff9e30073c49c999eaeadbe18ed68e84493815a4506",
  "Src/Commun/modCommandesLocales.bas": "14b16ee820dd6bfe339bfeb96710e262c5f6a8d935123e34b68859e17e0cc009",
  "Src/Commun/modConfig.bas": "1b4185d53f368160c731a978f636a029aa4d00e679fe48e78cf06868218477c3",
  "Src/Commun/modDonneesTransport.bas": "a2c730b1976a2f37b779f2a945f85a1e15bc11601656423f8d9079654dc556e8",
  "Src/Commun/modFichiers.bas": "7874f118e55e445942b67acbdd678f257657a9d8071a5226e980e84abebb2ca9",
  "Src/Commun/modGdt.bas": "4c575e42fa465925e3a46112d0bd5f99940adff521affdfd278fb76682590271",
  "Src/Commun/modIndexAnnexes.bas": "6cf4a3a89edf23c90f9c7f3af982da7e33e80d2ed6862c83e0c7c353a10022b5",
  "Src/Commun/modJson.bas": "95a7c2ceb1c7d4041ffdc6b92e144d1a0066523d89a3836bd0c383bdf0317636",
  "Src/Commun/modLog.bas": "487da3c9f9fa8266aef430a63f92422787a28f0bc1d12ef7cfd0c9cb94e2faf8",
  "Src/Commun/modServiceNas.bas": "2894660337f3c746849f5160b48d903418e3b137fc9da66c76cf9dd0f79f6e08",
  "Src/Commun/modTexte.bas": "cf753d80d6ee8d5b009ac35a05e42a6c9fb1b6cae24358249a7d4061d97f9fc7",
  "Src/ConfigDefaut/config.ini": "f4325a72562e7432a5eb083dc2df8c0607825305c80bc5f6469f9e774a598f09",
  "Src/Excel/Feuil1.cls": "8b6d1fc4f872634510d7a8323acaeca8b89024c115aac0c3e7c5dcfd25331905",
  "Src/Excel/FeuilAgenda.cls": "52f6239f0b04aea0e6b65e6cb30459592791d31b99ee490bb75ffe067d10fd63",
  "Src/Excel/ThisWorkbook.cls": "9f3a8a8c43e263590484651c2dea8863fd61e7032cc27a100e3a80cb14612717",
  "Src/Excel/modActes.bas": "7ca1b47cfc92a74115d8779af826fafd2212ce40321a39118060ebccb89dace2",
  "Src/Excel/modAgenda.bas": "c9a4208489534fb7339b2986fa067f22d863e0569f1ebd7fca828488be48d8bf",
  "Src/Excel/modAgendaVue.bas": "b0fd16b61264420789b0aab2842b69a97402b4e38959da99d3ffea12fa5442a0",
  "Src/Excel/modBaseIO.bas": "bcfb87f0d09b5e62127385030a16fd627d1ac49d9221f32fc73e35aefd5c172a",
  "Src/Excel/modCerfaPrint.bas": "dfd8b669367e0513b57f6997d0d204ef07ecafdf287a892f0c1b958583180324",
  "Src/Excel/modEchange.bas": "fbcfa55845fdaacf87330f1c6e07b8108e6b75b2d4f9f00462b4f3932a2cd318",
  "Src/Excel/modEditionSecretariat.bas": "fc6eb35079f6d125aeee4b0f7a83ac75248e82915d1282ac363f5e290cffbd0f",
  "Src/Excel/modJournal.bas": "866f32223e48f069c8d9b09cd082412f17f8a66021e5f7377ff4772425925218",
  "Src/Excel/modUI.bas": "271986da309ee11754c3e5c5529ee0c7700b6cde2d3a13391f5fe70ec38fadf1",
  "Src/Excel/ufChoixActe.vba": "09aa766e715379c97afec5068b605fa8b042cbd21b65c02d2679d7e7d9bb60f0",
  "Src/Excel/ufCorrespEdit.vba": "df69c57a58f3c79f239eea81da64c38d7d1802d921396076159a87b67a63005c",
  "Src/Excel/ufListe.vba": "1d40b76607705d85decc2a7d25da79d224c6dba6171fdc0701396f5812c4ebd7",
  "Src/Excel/ufPatientEdit.vba": "8781bc6a97c74b9b422e2af04e529258b6916b5d8a7ce0d4404faf5a9df307e6",
  "Src/Excel/ufRdvEdit.vba": "b77490776a98c3caf72ef8eacd20d2cb94e2799a4b4b45c5e8eef88fed79655d",
  "Src/Integration/modAttenteLocale.bas": "e41813ce204b60d60ab15667ba006bc2b205c491917dcd6dcde8589c7794897f",
  "Src/Integration/modControleCourrier.bas": "b38de4a5e54893f6aa9f99b64d0a009a1b674f089f794435ae0a603cd3c560dd",
  "Src/Integration/modCycleCourrier.bas": "d15600d1a7634c93db5f80355fe98cdf856e43aeefc4068c72a611a48cbb9c91",
  "Src/Integration/modEtatCourrier.bas": "1f50c4dc21dd2a4d868462edb7dc01e41b4232984e6423b8c3d30993b2794005",
  "Src/Integration/modFileArrivees.bas": "5990f478c63f1ff31f15cff7b8ecd47309b92b5b37a8b5c138ed83357a636df3",
  "Src/Integration/modIntegrationUnifie.bas": "55c401062d1a85e79a57932178773b69394f814080704ae9bbabd6129d678d98",
  "Src/Integration/modPowerMicUnifie.bas": "f3b8c25e4165a2b6e9800602d8b21c96a38bc2bdee87462c118677fb2a945a65",
  "Src/Integration/modRechercheWord.bas": "63137fbff9b20f04070e95829f8ce032be5c14190c2b4eea0418c7e55b3ef4fa",
  "Src/Prod6/ThisDocument.cls": "6d7b7fd22869872b04bcc106514b36a55fb7bf8d9fa8eceb59a213b2770db4ba",
  "Src/Prod6/frmCTDestination.vba": "1d2f26456daa9f424139c153f5534779b30d212629e138f322f6b5275673019d",
  "Src/Prod6/modAnonymisation.bas": "2c0edc11aaa0bdd4b90f63f1807deb52521040f5af2a9d6a0c6a1f0855fcbe88",
  "Src/Prod6/modBaseCorrespondants.bas": "ae5cfa3585806742364fd5d1b0174e16cf114cdab8b701163c2bf146be0fe2f0",
  "Src/Prod6/modCommandesDDERAC.bas": "8fdc6842aabcfcbc7fa3a7065642a9c59ce7eaab5640a2e29184e56a9e61db68",
  "Src/Prod6/modDemandesAnnexes.bas": "b4b73917a6ac8a41d7c193d7cf5a802878c52cff5f3a96a4ed7d5cc5c57df4f8",
  "Src/Prod6/modDestinations.bas": "fafbc8f3de5c9c43c19ae3c77a7e594bd1ba39166619d4fb5611b93aa1db499b",
  "Src/Prod6/modDetectionDemandesDico.bas": "a137a9aaad79fdbae45726d34355cc08648e38c3a7f710aeb2dd3084985bc82f",
  "Src/Prod6/modEnregistrementCorrespondants.bas": "535e57d7991a670f6a53bb3f2e96311abc46e71100d8a5ae219a069c64d1710b",
  "Src/Prod6/modGenerationDemandes.bas": "d9d6141c130f175e1b21d9e10ae001cef25f88d3856d014208e6e1b99f2113a4",
  "Src/Prod6/modGlobals.bas": "e7dfd0affcc2382b061f63e07b284b72bb513b465d42472a628f71bc80b022f9",
  "Src/Prod6/modInsertion.bas": "4873abdf727ed1059413a8129aeacea01231a413a809cbe78f8c227164156d66",
  "Src/Prod6/modLettresComplementairesModele.bas": "910b3a9ffe1bfce04291a1457f857879d82fade2fd76de59e1053c1972594567",
  "Src/Prod6/modMiseEnPageFinale.bas": "b85a06b8fd0d783d5c8112095fef1ccd23e754394f8acbca16a74072d5abeeaf",
  "Src/Prod6/modNormalisationGrasLocal.bas": "177713a5d9821032ed2718470da73e3caba51f393e365e253b69c642ef7b370f",
  "Src/Prod6/modOpenAI_v22_corrige.bas": "a74a543adfe6cbde3b906da5820b77cde40ec5cba7147b59b00058e7bd346e3b",
  "Src/Prod6/modProdRapide.bas": "1e890098202bf7f1daabfb0f7b659a883688682c00cb4f9e0eef2a971353c339",
  "Src/Prod6/modPrompts.bas": "01ca059b62885151338a8a73dd7f4f1ad744a9cf0664d7d6e4654ad729f5791c",
  "Src/Prod6/modSaisieCorrespondants.bas": "f1c644ee8f0b51b6fb12837974e3e5fab7c39d93163871879436a7ad9159372a",
  "Src/Prod6/modSaisieStructures.bas": "51aa1784c5066b39b646b48f831a3c069170a544ec8367b36d9ae3f8035fc0a7",
  "Src/Prod6/modSortieDragon.bas": "0bd0cf550bec8944cdb23156b76595e2f2e9b3843f95fb9063b91708e9b606f2",
  "Src/Prod6/modUtilitaires.bas": "dc574f7a1fc7621fc1af61f7348f5a0c8ecd606c34c47d6d23acaf13102eac13",
  "Src/Word/modAnonymise.bas": "b40022cf9f3f3752a7277c67119863dce3f954ec16a7d8131c886326e594567d",
  "Src/Word/modApiConfiguration.bas": "c87a6a730407061c67416b8b735a6541a7b5c92514861fe548cce7b5f1ac91cd",
  "Src/Word/modBase.bas": "c04ac90da7a8e5936ebdcbc62ac140aa441f0898e63cd4566eef7668b05f3184",
  "Src/Word/modCourrier.bas": "574eab221674a2257d8c44b789971d2196fbbcdbb9e716ace9e5ede22802a500",
  "Src/Word/modDemandes.bas": "c3ac084dabdd22670a00afce8e9018989dc501816f5144eed27148b3798c77ef",
  "Src/Word/modDerivees.bas": "b169a58e21e0bd874e80927aceeaa1fdde03578f770f7e2049cf9dc7a5be1edc",
  "Src/Word/modEcg.bas": "bda7e2dcdc3025cb98bfc55b6da2c87c75b3abcf85f710a113735053c9b29811",
  "Src/Word/modGras.bas": "18f266fa40ef6491292ea4a36d7212c2fc580323e6bc8c8da434c6412232d351",
  "Src/Word/modPatient.bas": "01a8553d774e7df3cbbfd9d39b9ba106006b9d6ce0b5e041e8608bb0daf2d1b6",
  "Src/Word/modPowerMic.bas": "4f22aa1ec058f8663fdba691405ab887f9ff6c77666a59f5c61873fa5e572da3",
  "Src/Word/modRaccourcis.bas": "ce1d8b35755be82e5539bebb358d1b12b8571db73d2d72c3469f3788cb06e417",
  "Src/Word/modRuban.bas": "d1ff7bf10a4d6c0a7cb2c9fa56f4e828763f475ca0a69c135e940aa072d8abb4",
  "Src/Word/modSchemas.bas": "45fb971cbf846a2af4993756e471fa1ef40fc7cd6f5f9ddc55a042f104938ed6",
  "Src/Word/modSubstitutions.bas": "a24ae57953c43e2c2bafd09e539d13badb46f5d8d417b37a162694b9fc94c302",
  "Src/Word/modValidation.bas": "23c6bed607959550e8bdc451633aebeba00c3f2d7fb7eb5b65bd454e162a1be6",
  "Src/Word/ufArriveesMedecin.vba": "e933be7228273a37e7ace7953bf0fde186cb04b2d80cc5e996133395ba7f61d3",
  "Src/Word/ufListe.vba": "fd55bbfcc0c18450f1d5ff2739aaa5f266c2b52d40b7fbcee5358c520c683113",
  "Src/Word/ufProgression.vba": "49b383108d2e7c85f27b4f4aadd438c0251042e412cc36c06a60adb7a611f6b2",
  "Tests/Vba/excel/modRecetteU1Excel.bas": "491c10b6f2c40284a9a3286c11922d37647d15cc1ac97d3dbc17f220e2b9476b",
  "Tests/Vba/word/modAuditTests.bas": "290f74cff7025931df0106d1ecc7d8046e888010f5dc0ccb9523bcceb9a2eae2",
  "Tests/Vba/word/modRecetteModeleCourrier.bas": "21b0b49a516097d2787544c51bc1093d593282fe82d729bc871fb8c94a9c0797",
  "Tests/Vba/word/modRecettePresentationAnnexe.bas": "134ef91fd9e1b50ec84ad3b9fe770f370d416f33ce03de118b2c7c504bf27f66",
  "Tests/Vba/word/modRecetteU0.bas": "75fe44170262584fc150b0417451e0d87d2b4d9b4cfa7a98f4b738a57f56f1a5",
  "Tests/Vba/word/modRecetteU1.bas": "d992b9d3fdd312e1b65e9621d616edd20973a3757a17585fa640266c13379992",
  "Tests/Vba/word/modRecetteU2.bas": "798ace1db06d0b8d04a98cf1460bdfc9c494b3c0e8a347ca28db54290aaf7fb8",
  "Tests/audit_statique.py": "11e10b5757c30b3cc5718841d0d6ebfe235a7bcdbc13d77fae7356da01dba671",
  "Tests/inventaire_architecture.py": "1667db4e7a5f1e5fa0bdedd293a179188a9e8dff2b8bde4116273e4641d5b384",
  "Tests/inventaire_sources.json": "52c6fdcbbd448a900c3344c1d780b1fdc349973a2500dc509b382540c67a4f33",
  "Tests/inventaire_u2.json": "943f8f29c0eb0fd00d3878ab8a82902dd98f646d88fcc36e68d75ae0c636a058",
  "Tests/nettoyage_sources.json": "f3b54768aca28e41878d9bd36a147d7f2f7da5815609e85576d2a03986c5e07d",
  "Tests/schema_reponse_api.json": "fbf5f9805faa23b77cdffe368bdd737e320665c58b75a528a268f69ac06665f4",
  "Tests/test_assistant.ps1": "11cc4bb1293777569a72537d23a203dbc70da60f191ae37fc9650b9b728bc2f6",
  "Tests/test_attente_office.ps1": "1859e958edd9d74ce5b3d27d8d8faa976bf5576e5138d9a3fa6ac608f53b7677",
  "Tests/test_choix_nas.ps1": "54dd2afe5dde485758abdb0dc8064f00bf5e6a7b87689fe41e8deebd65201c0c",
  "Tests/test_compilation_office.ps1": "dfb927de3bf82ba3d31a304b17f944fdb0ae3d73b589ab2f7002aadcf0b2de9b",
  "Tests/test_connexion_service.ps1": "275cae8a974bd47df1ad615308d723f034f0a599d23d3dc6993abcad9b84426f",
  "Tests/test_construction.ps1": "ea6f3384499cb1c2ec48c2bf0f66145ef736c95ec0e65ef84057f2c1b835e46c",
  "Tests/test_edition_documents.py": "34f2bb24f2c4764eb4f72a84284e88e73183cb121ad79ff153ab70e733bbea8f",
  "Tests/test_feuilles_excel.ps1": "4d552b64e4a44913daa035af1ee346e2c7d9ac3f1b1d8c9218f1eef82a21c48a",
  "Tests/test_image_permissions.sh": "ea684664d1111acf78305173ece4fac556511f1d6c7e7041a3c9d47ff3f97305",
  "Tests/test_installation.ps1": "923bff34cbe62f22af7bd188fff5123d227fc932d09a886ab2865824b3f2b682",
  "Tests/test_lancement_direct.ps1": "44dd1fc6b61d15d73c779b298764470954a77f51a6f158ec9282c77bbdf71cc5",
  "Tests/test_recette_20260920.py": "84221b1a0256f2a6a112272da8b1d9f67fa75ec99dc8eaf9541021439e6bb2d7",
  "Tests/test_recette_isolee.ps1": "2027086609d48bbef79407065e04f6b52f5acd27f7571b2bb4df63b880776672",
  "Tests/test_recovery_docker.sh": "3fac41479b6ffe419633b8287102f336e4db84861c0fc879637dc49d608b4fec",
  "Tests/test_restauration_poste.ps1": "46672e21e6e44c79609af2b04a862bc31327f4c5222e2962840bd6434ce4d352",
  "Tests/test_u0_installation.ps1": "5f87ef07bd91f8f9b5aa5698ede582059938590438338001f8a900074281a8c1",
  "Tests/test_u1_installation.ps1": "238228e934846fccd9a63a61020c488fbd292ce386920cc40da7242ccca6023f",
  "Tests/test_u2_architecture.py": "faa2b2f55fc717c469b4cf1aaf442c9f468b222b3e841f5647f771189eea0919",
  "Tests/test_u2_manifeste.ps1": "b73a657cdf3434fcce70addca01cff68b812c0361882418bda9dfd7b7e4792bd",
  "Tests/verification_livraison.json": "e1733c2ce5b5523783ef9d8385971ba768499d6acb7e411b98179576cf9cbb50",
  "U0_RECETTE.md": "30208817c337df7b89349d71ecbf80bfe77b9f456ebde5e184aef27ef1531987",
  "U1_RECETTE.md": "faec11f24bc95db3e35a2d690483ba76b9d1cd8c18c30f136ad837962a86d959",
  "U2_CORRECTIFS_AUDITS.md": "1d435d0ecf702209c34d9c7639cf74fa541daa6d2f62e652f24147f673608c5b",
  "U2_RECETTE.md": "26c95631d51b1ef4f3ecd1c07fb78a4e5280d7fb4f87207279d9f402b72ec856"
}
'@ | ConvertFrom-Json
$bootLock=$null
try {
    if ($env:OS -ne 'Windows_NT') { throw 'Ce lanceur necessite Windows.' }
    Write-Host ('Cabinet Cardio - installation guidee - version '+$Commit.Substring(0,7))
    Write-Host 'Recette U2b uniquement : utilisez le compte Windows dedie aux tests, jamais le profil clinique U0.'
    Write-Host 'Lancez sans Executer en tant qu administrateur.'
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
