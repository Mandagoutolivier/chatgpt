# Instrumentation des copies de qualification uniquement ; jamais une livraison utilisable.
function Remplacer-ProcedureRecette([string]$Code,[string]$Nom,[string]$Corps) {
    $pattern='(?ims)^(?<debut>(?:Public|Private|Friend)?\s*(?<type>Function|Sub)\s+'+[regex]::Escape($Nom)+'\([^\r\n]*\)[^\r\n]*\r?\n).*?^End\s+\k<type>\s*$'
    $matches=[regex]::Matches($Code,$pattern)
    if ($matches.Count -ne 1) { throw ('Procedure a isoler non unique : '+$Nom) }
    return [regex]::Replace($Code,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Groups['debut'].Value+$Corps+"`r`nEnd "+$m.Groups['type'].Value})
}

function Preparer-CopieRecetteIsolee($Projet,[string]$Sortie,[ValidateSet('Word','Excel')][string]$Hote) {
    $bloquer='    Err.Raise vbObjectError + 1999, "Qualification isolee", "Acces externe interdit pendant la qualification."'
    $base=Join-Path $Sortie 'Environnement'
    $racine=Join-Path $base 'DonneesFictives'
    $replacements=@{APPDATA=(Join-Path $base 'AppData');LOCALAPPDATA=(Join-Path $base 'LocalAppData');TEMP=(Join-Path $base 'Temp');TMP=(Join-Path $base 'Temp')}
    foreach ($path in @($replacements.Values)+@($racine,(Join-Path $racine 'Config'),(Join-Path $racine 'Logs'))) { [void][IO.Directory]::CreateDirectory($path) }
    $patches=[Collections.Generic.List[object]]::new()
    foreach ($component in @($Projet.VBComponents)) {
        $module=$component.CodeModule
        $count=[int]$module.CountOfLines
        if ($count -eq 0) { continue }
        $before=[string]$module.Lines(1,$count)
        $code=$before
        switch ([string]$component.Name) {
            'modServiceNas' { $code=Remplacer-ProcedureRecette $code 'Appeler' $bloquer }
            'modOpenAI_v22_corrige' { $code=Remplacer-ProcedureRecette $code 'AppelerOpenAIRaw' $bloquer }
            'modApiConfiguration' {
                $code=Remplacer-ProcedureRecette $code 'LireCleOpenAICabinet' $bloquer
                $code=Remplacer-ProcedureRecette $code 'LireCleApi' $bloquer
            }
            'modConfig' {
                $pattern='(?im)^(Public Function racine\(\) As String\s*)$'
                if ([regex]::Matches($code,$pattern).Count -ne 1) { throw 'Fonction racine de la copie inattendue.' }
                $injection='    If Len(mRacine) = 0 Then mRacine = "'+$racine.Replace('"','""')+'"'
                $code=[regex]::Replace($code,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value+"`r`n"+$injection})
            }
        }
        # Coupe toute procedure de demarrage / fermeture des seules copies construites.
        $autos=[regex]::Matches($code,'(?im)^(?:Public|Private|Friend)?\s*Sub\s+(AutoExec|AutoOpen|AutoNew|AutoClose|AutoExit|Document_\w+|Workbook_\w+|Worksheet_\w+)\(')
        foreach ($entry in $autos) { $code=Remplacer-ProcedureRecette $code $entry.Groups[1].Value "    ' Evenement neutralise dans la copie de qualification." }
        foreach ($name in $replacements.Keys) {
            $pathLiteral='"'+$replacements[$name].Replace('"','""')+'"'
            $code=[regex]::Replace($code,'(?i)Environ\$?\("'+$name+'"\)',[Text.RegularExpressions.MatchEvaluator]{param($m) $pathLiteral})
        }
        # Aucune impression ni boite d apercu meme si une regression atteint ce code.
        $code=[regex]::Replace($code,'(?im)^\s*\w+\.Print(?:Out|Preview)\b[^\r\n]*\r?$', $bloquer)
        if ($code -match '(?i)CreateObject\("(?:WinHttp\.|MSXML2?\.(?:Server)?XMLHTTP|Microsoft\.XMLHTTP)' -or
            $code -match '(?i)Environ\$?\("(?:APPDATA|LOCALAPPDATA|TEMP|TMP|OPENAI_API_KEY)"\)' -or
            $code -match '(?im)^\s*\w+\.Print(?:Out|Preview)\b') { throw ('Entree externe non neutralisee : '+$component.Name) }
        if ($code -cne $before) {
            $module.DeleteLines(1,$count)
            $module.AddFromString($code)
            $patches.Add([pscustomobject]@{Module=[string]$component.Name;CopieUniquement=$true})
        }
    }
    if ('modServiceNas' -notin @($patches|ForEach-Object {$_.Module})) { throw 'Blocage RPC absent de la copie.' }
    if ($Hote -eq 'Word' -and 'modOpenAI_v22_corrige' -notin @($patches|ForEach-Object {$_.Module})) { throw 'Blocage OpenAI absent.' }
    [IO.File]::WriteAllText((Join-Path $Sortie ('instrumentation-'+$Hote+'.json')),($patches|ConvertTo-Json -Depth 4),[Text.UTF8Encoding]::new($false))
}
