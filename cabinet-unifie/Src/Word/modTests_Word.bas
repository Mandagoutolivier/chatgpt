Attribute VB_Name = "modTests_Word"
Option Explicit
Option Private Module
' =====================================================================
' modTests_Word - Suites de tests lancees par run_tests.ps1 (COM).
' Chaque Test_Jn ecrit PASS/FAIL dans <Racine>\Logs\tests.log.
' Donnees FICTIVES uniquement (SandboxData).
' =====================================================================

Public Sub Test_J0(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "J0 socle"
    On Error GoTo Echec

    ' --- configuration ---
    modLog.Verifier "config.ini lu", modConfig.Config("api", "modele", "") <> "", modConfig.Config("api", "modele", "")
    modLog.Verifier "racine existe", Len(Dir$(modConfig.racine(), vbDirectory)) > 0

    ' --- UTF-8 aller-retour ---
    Dim p As String, t As String
    p = modConfig.chemin("Logs") & "\_test_utf8.txt"
    modFichiers.EcrireTexteUTF8 p, "éàçœ€ ligne1" & vbCrLf & "ligne2"
    t = modFichiers.LireTexteUTF8(p)
    modLog.Verifier "utf8 aller-retour", InStr(t, "éàçœ€") = 1 And InStr(t, "ligne2") > 0
    Kill p

    ' --- drapeaux ---
    Dim d As Object, d2 As Object, cheminDrapeau As String
    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = "P00001"
    d("Type") = "consultation"
    d("Accent") = "échéance"
    cheminDrapeau = modFichiers.EcrireDrapeau(modConfig.chemin("Echange") & "\AEnvoyer", "_test_" & modFichiers.IdUnique(), d)
    Set d2 = modFichiers.LireDrapeau(cheminDrapeau)
    modLog.Verifier "drapeau aller-retour", d2("PatientID") = "P00001" And d2("Accent") = "échéance"
    Kill cheminDrapeau

    ' --- verrous cooperatifs ---
    modLog.Verifier "verrou acquis", modFichiers.AcquerirVerrou("testlock", 2000)
    modLog.Verifier "verrou refuse en double", Not modFichiers.AcquerirVerrou("testlock", 1200)
    modFichiers.RelacherVerrou "testlock"
    modLog.Verifier "verrou reacquis apres liberation", modFichiers.AcquerirVerrou("testlock", 2000)
    modFichiers.RelacherVerrou "testlock"

    ' --- copie locale de la base ---
    Dim cheminLocal As String
    cheminLocal = modFichiers.CopieLocale(modConfig.FichierPatients())
    modLog.Verifier "copie locale patients", Len(Dir$(cheminLocal)) > 0, cheminLocal

    ' --- JSON ---
    Dim r As Object, txt As String
    Set r = modJson.JsonParse("{""content"":[{""type"":""text"",""text"":""Bonjour\nété""}],""stop_reason"":""end_turn"",""usage"":{""input_tokens"":12}}")
    txt = modJson.JsonTexteReponse(r)
    modLog.Verifier "json texte reponse", txt = "Bonjour" & vbLf & "été", txt
    modLog.Verifier "json nombre", r("usage")("input_tokens") = 12
    Dim rErr As Object
    Set rErr = modJson.JsonParse("{""type"":""error"",""error"":{""type"":""invalid_request_error"",""message"":""bad key""}}")
    modLog.Verifier "json message erreur", modJson.JsonErreurMessage(rErr) = "bad key"
    modLog.Verifier "json echappement", modJson.JsonEchapper("a""b\c" & vbCrLf & "d") = "a\""b\\c\nd"

    ' --- sauvegarde horodatee ---
    modFichiers.SauvegardeHorodatee modConfig.FichierPatients()
    modLog.Verifier "sauvegarde horodatee", Len(Dir$(modConfig.chemin("Sauvegardes") & "\Patients_*.xlsx")) > 0

    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

' --- sondes de diagnostic (bissection des blocages) -------------------
Public Sub Test_SONDE(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "sondes"
    On Error GoTo Echec
    Dim chemin As String
    chemin = modFichiers.CopieLocale(modConfig.FichierPatients())
    modLog.TestResultat "sonde copie locale", True, chemin
    Dim n As Long, col As Collection
    Set col = modBase.LireTable(modConfig.FichierPatients(), "PATIENTS")
    n = col.Count
    modLog.TestResultat "sonde ADO LireTable", n = 20, "lignes=" & n
    Dim p As Object
    Set p = modBase.PatientParID("P00001")
    modLog.TestResultat "sonde PatientParID", Not p Is Nothing
    Dim f As ufListe
    Set f = New ufListe
    modLog.TestResultat "sonde instanciation ufListe", True
    Unload f
    Dim doc As Document
    Set doc = Documents.Add(Template:=modConfig.chemin("Modeles") & "\LETTRE TYPE.dot")
    modLog.TestResultat "sonde Documents.Add modele", True, doc.Bookmarks.Count & " signets"
    doc.Close 0
    Exit Sub
Echec:
    modLog.TestResultat "sonde exception " & Err.Number, False, Err.Description
End Sub

Public Sub Test_J3(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "J3 substitutions + anonymisation (hors ligne)"
    On Error GoTo Echec
    Dim pat As Object, cor As Object, ctx As Object
    Dim corps As String, anonyme As String, pb As String, final As String

    Set pat = modBase.PatientParID("P00001")     ' FABREGUE Jean, 01/01/1935
    Set cor = modBase.CorrespondantParID("C0001") ' Dr GARRIGUE Paul

    ' --- substitutions locales ---
    corps = "La tension est a 130/80 mm de mercure, la frequence a 70 par minute."
    corps = modSubstitutions.AppliquerSubstitutions(corps)
    modLog.Verifier "substitution mmHg", InStr(corps, "130/80 mmHg") > 0
    modLog.Verifier "substitution /min", InStr(corps, "70 /min") > 0
    corps = modSubstitutions.AppliquerSubstitutions("Il prend du kardégic 75 et de l'eliquis.")
    modLog.Verifier "medicaments en majuscules", InStr(corps, "KARDÉGIC") > 0 And InStr(corps, "ELIQUIS") > 0, corps

    ' --- anonymisation ---
    corps = "J'ai recu en consultation Monsieur Jean FABREGUE, ne le 01/01/1935, " & _
            "adresse par le Docteur Garrigue. Ce patient (M. Fabregue, 1er janvier 1935) " & _
            "vit au 1 rue des Oliviers a Montpellier. Tel : 04 67 11 22 00."
    Set ctx = modAnonymise.Construire(pat, cor)
    anonyme = modAnonymise.Anonymiser(corps, ctx)
    modLog.Verifier "nom absent", InStr(1, anonyme, "FABREGUE", vbTextCompare) = 0, anonyme
    modLog.Verifier "prenom absent", InStr(anonyme, "Jean") = 0
    modLog.Verifier "ddn numerique absente", InStr(anonyme, "01/01/1935") = 0
    modLog.Verifier "ddn en lettres absente", InStr(anonyme, "1er janvier 1935") = 0
    modLog.Verifier "nom correspondant absent", InStr(1, anonyme, "Garrigue", vbTextCompare) = 0
    modLog.Verifier "adresse absente", InStr(anonyme, "rue des Oliviers") = 0
    modLog.Verifier "telephone absent", InStr(anonyme, "04 67 11 22 00") = 0
    modLog.Verifier "balises presentes", InStr(anonyme, "{{PAT_NOM}}") > 0 And InStr(anonyme, "{{PAT_DDN}}") > 0

    ' --- scan residuel ---
    pb = modAnonymise.ScanResiduel(anonyme, ctx)
    modLog.Verifier "scan residuel propre", Len(pb) = 0, pb
    pb = modAnonymise.ScanResiduel(anonyme & " NIR 1 35 01 34 000 001 42", ctx)
    modLog.Verifier "scan bloque un NIR", InStr(pb, "NIR") > 0, pb
    pb = modAnonymise.ScanResiduel("Rappeler le patient au 06 11 22 33 44.", ctx)
    modLog.Verifier "scan bloque un telephone", InStr(pb, "telephone") > 0, pb
    pb = modAnonymise.ScanResiduel("Le patient Fabregue est revenu.", ctx)
    modLog.Verifier "scan bloque une identite restante", InStr(pb, "identite") > 0, pb

    ' --- verification des balises au retour ---
    pb = modAnonymise.VerifierBalisesRetour(anonyme, ctx)
    modLog.Verifier "balises retour valides", Len(pb) = 0, pb
    pb = modAnonymise.VerifierBalisesRetour("Texte avec {{PAT_INVENTEE}}.", ctx)
    modLog.Verifier "balise alteree detectee", Len(pb) > 0, pb

    ' --- reinjection ---
    final = modAnonymise.Reinjecter(anonyme, ctx)
    modLog.Verifier "reinjection nom", InStr(final, "FABREGUE") > 0
    modLog.Verifier "reinjection ddn", InStr(final, "01/01/1935") > 0
    modLog.Verifier "plus aucune balise", InStr(final, "{{") = 0, final

    ' --- nettoyage de la reponse : lignes vides supprimees ---
    Dim net As String
    net = modClaude.NettoyerReponse("Un." & vbLf & vbLf & "Deux." & vbCrLf & vbCrLf & vbCrLf & " Trois. " & vbLf)
    modLog.Verifier "reponse sans lignes vides", net = "Un." & vbCr & "Deux." & vbCr & "Trois.", Replace(net, vbCr, "|")

    ' --- prompt et payload debug (sans appel reseau) ---
    Dim prompt As String
    prompt = modClaude.ChargerPrompt("correction.txt")
    modLog.Verifier "prompt correction charge", InStr(prompt, "balises") > 0
    modLog.Verifier "prompt sans marqueur reference", InStr(prompt, "[COURRIERS_DE_REFERENCE]") = 0
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

' Appel API reel (necessite %APPDATA%\CabinetCardio\api.key) - donnees fictives
Public Sub Test_J3API(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "J3 appel API reel"
    On Error GoTo Echec
    Dim pat As Object, cor As Object, ctx As Object
    Dim corps As String, anonyme As String, reponse As String, final As String

    Set pat = modBase.PatientParID("P00001")
    Set cor = modBase.CorrespondantParID("C0001")
    corps = "j'ai vu ce jour monsieur Jean FABREGUE ne le 01/01/1935 pour bilan de palpitation " & _
            "la tension et a 130/80 mm de mercure la frequence cardiaque a 70 par minute " & _
            "l'auscultation est sans particularite il prend du kardégic 75"
    corps = modSubstitutions.AppliquerSubstitutions(corps)
    Set ctx = modAnonymise.Construire(pat, cor)
    anonyme = modAnonymise.Anonymiser(corps, ctx)
    modLog.Verifier "pre-envoi : scan propre", Len(modAnonymise.ScanResiduel(anonyme, ctx)) = 0

    reponse = modClaude.AppelerClaude(modClaude.ChargerPrompt("correction.txt"), anonyme)
    modLog.Verifier "reponse recue", Len(reponse) > 20, Left$(reponse, 80)
    modLog.Verifier "balises retour intactes", Len(modAnonymise.VerifierBalisesRetour(reponse, ctx)) = 0
    ' verification cruciale : payload_debug ne contient AUCUNE identite
    Dim payload As String
    payload = modFichiers.LireTexteUTF8(modConfig.chemin("Logs") & "\payload_debug.json")
    modLog.Verifier "payload sans nom", InStr(1, payload, "FABREGUE", vbTextCompare) = 0
    modLog.Verifier "payload sans ddn", InStr(payload, "01/01/1935") = 0
    final = modAnonymise.Reinjecter(reponse, ctx)
    modLog.Verifier "final reinjecte", InStr(final, "FABREGUE") > 0, Left$(final, 120)
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

Public Sub Test_J5W(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "J5 validation cote medecin"
    On Error GoTo Echec
    Dim pat As Object, cor As Object, doc As Document
    Dim fichiers As String, drapeaux As String

    Set pat = modBase.PatientParID("P00002")
    Set cor = modBase.CorrespondantParID("C0002")
    Set doc = modCourrier.CreerCourrierPour(pat, cor)
    modCourrier.RemplacerCorps doc, "Corps de test pour la validation."
    ' simulation de ValiderCourrier sans MsgBox : on reprend sa logique
    Dim dossier As String, base As String, d As Object, chemin As String
    dossier = modPatient.DossierPatient(pat)
    base = Format$(Now, "yyyymmdd-hhnnss") & " test"
    doc.SaveAs2 dossier & "\" & base & ".docx", 12
    doc.ExportAsFixedFormat dossier & "\" & base & ".pdf", 17
    modLog.Verifier "docx enregistre", Len(Dir$(dossier & "\" & base & ".docx")) > 0
    modLog.Verifier "pdf exporte", Len(Dir$(dossier & "\" & base & ".pdf")) > 0
    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = pat("ID")
    d("Nom") = pat("Nom")
    d("TypeCourrier") = "consultation"
    d("CheminDocx") = dossier & "\" & base & ".docx"
    d("CheminPdf") = dossier & "\" & base & ".pdf"
    chemin = modFichiers.EcrireDrapeau(modConfig.chemin("Echange") & "\AEnvoyer", _
                                       modFichiers.IdUnique() & "_" & pat("ID"), d)
    modLog.Verifier "drapeau depose", Len(Dir$(chemin)) > 0, chemin
    doc.Close 0
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

Public Sub Test_GDT(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "GDT envoi identite vers ECG"
    On Error GoTo Echec
    Dim pat As Object, dossier As String, chemin As String, contenu As String
    Dim lignes() As String, declare8100 As Long

    Set pat = modBase.PatientParID("P00001")
    dossier = modConfig.chemin("Logs")
    chemin = modGdt.EcrireGdtPatient(pat, dossier)
    modLog.Verifier "fichier IMPORT.GDT ecrit", Len(Dir$(chemin)) > 0, chemin

    ' relecture en ANSI (le GDT doit etre cp1252, pas UTF-8)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(chemin, 1, False, 0)
    contenu = ts.ReadAll
    ts.Close
    lignes = Split(contenu, vbCrLf)
    modLog.Verifier "satz 6302", lignes(0) = "01380006302", lignes(0)
    modLog.Verifier "version GDT 02.00", lignes(3) = "014921802.00", lignes(3)
    modLog.Verifier "nom patient", InStr(contenu, "0173101FABREGUE") > 0
    modLog.Verifier "prenom patient", InStr(contenu, "0133102Jean") > 0
    modLog.Verifier "ddn JJMMAAAA", InStr(contenu, "017310301011935") > 0
    modLog.Verifier "code examen", InStr(contenu, "8402EKG01") > 0
    declare8100 = Val(Mid$(lignes(1), 8, 5))
    modLog.Verifier "longueur 8100 = taille fichier", declare8100 = FileLen(chemin), _
                    declare8100 & " / " & FileLen(chemin)
    Kill chemin
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

' Sonde d'environnement : ce que Word voit du dossier %APPDATA%\CabinetCardio
Public Sub Test_ENV(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "ENV vue de Word sur %APPDATA%"
    On Error Resume Next
    Dim fso As Object, d As String, f As Object, liste As String, t As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    d = Environ$("APPDATA") & "\CabinetCardio"
    modLog.TestResultat "APPDATA=" & Environ$("APPDATA") & " USER=" & Environ$("USERNAME"), True
    modLog.TestResultat "dossier existe (FSO)", fso.FolderExists(d), d
    For Each f In fso.GetFolder(d).Files
        liste = liste & f.Name & "(" & f.Size & ") "
    Next f
    modLog.TestResultat "listing FSO", True, liste
    modLog.TestResultat "chemin.txt FSO", fso.FileExists(d & "\chemin.txt")
    modLog.TestResultat "chemin.txt Dir$", Len(Dir$(d & "\chemin.txt")) > 0
    modLog.TestResultat "api.key FSO", fso.FileExists(d & "\api.key")
    Err.Clear
    t = modFichiers.LireTexteUTF8(d & "\chemin.txt")
    modLog.TestResultat "lecture chemin.txt", Err.Number = 0, IIf(Err.Number = 0, Left$(t, 80), Err.Description)
    Err.Clear
    Dim num As Integer, ligne As String
    num = FreeFile
    Open d & "\chemin.txt" For Input As #num
    If Err.Number = 0 Then Line Input #num, ligne: Close #num
    modLog.TestResultat "Open For Input", Err.Number = 0, IIf(Err.Number = 0, ligne, Err.Description)
End Sub

' Reproduit le flux reel : creation, frappe (dictee) dans le corps, correction
Public Sub Test_RETRAITS(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "RETRAITS : frappe puis correction"
    On Error GoTo Echec
    Dim pat As Object, cor As Object, doc As Document, i As Long, s As String
    Set pat = modBase.PatientParID("P00001")
    Set cor = modBase.CorrespondantParID("C0001")
    Set doc = modCourrier.CreerCourrierPour(pat, cor)
    ' formats du modele autour du corps
    s = ""
    For i = 1 To doc.Paragraphs.Count
        s = s & i & ":" & Round(doc.Paragraphs(i).LeftIndent / 28.35, 1) & "/" & Round(doc.Paragraphs(i).FirstLineIndent / 28.35, 1) & " "
    Next i
    modLog.TestResultat "formats initiaux (gauche/1re ligne cm)", True, s
    modLog.Verifier "en-tete conserve (retrait 1er paragraphe)", Abs(doc.Paragraphs(1).LeftIndent - doc.Paragraphs(1).LeftIndent) < 0.5
    ' frappe comme en dictee : curseur place par PlacerCurseurCorps, Entree entre paragraphes
    doc.Activate
    modCourrier.PlacerCurseurCorps doc
    selection.TypeText "Premier paragraphe dicte."
    selection.TypeParagraph
    selection.TypeText "Deuxieme paragraphe dicte."
    selection.TypeParagraph
    selection.TypeText "Troisieme paragraphe dicte."
    s = Formats(doc.Bookmarks("CORPS").Range)
    Dim k As Long, st As String
    For k = 1 To doc.Bookmarks("CORPS").Range.Paragraphs.Count
        st = st & "[" & doc.Bookmarks("CORPS").Range.Paragraphs(k).Style & "] "
    Next k
    modLog.Verifier "apres frappe : retraits uniformes (dictee)", uniforme(doc.Bookmarks("CORPS").Range), s & " styles : " & st
    modLog.TestResultat "corps relu apres frappe", InStr(modCourrier.RecupererCorps(doc), "Troisieme") > 0, Replace(modCourrier.RecupererCorps(doc), vbCr, " | ")
    Dim espF As String, espFOk As Boolean, qf As Paragraph
    espFOk = True
    For Each qf In doc.Bookmarks("CORPS").Range.Paragraphs
        espF = espF & qf.SpaceBefore & "/" & qf.SpaceAfter & IIf(qf.SpaceBeforeAuto Or qf.SpaceAfterAuto, "auto", "") & " "
        If qf.SpaceBeforeAuto Or qf.SpaceAfterAuto Or Abs(qf.SpaceBefore - 12) > 0.5 Or qf.SpaceAfter > 0.5 Then espFOk = False
    Next qf
    modLog.Verifier "espacement apres frappe (avant/apres pt)", espFOk, espF
    ' correction simulee : 4 paragraphes
    modCourrier.RemplacerCorps doc, "Un." & vbCr & "Deux." & vbCr & "Trois." & vbCr & "Quatre."
    s = Formats(doc.Bookmarks("CORPS").Range)
    modLog.Verifier "apres correction : retraits uniformes", uniforme(doc.Bookmarks("CORPS").Range), s
    ' espacement : 12 pt avant, 0 apres, pas d'espacement automatique
    Dim espOk As Boolean, esp As String, q As Paragraph
    espOk = True
    For Each q In doc.Bookmarks("CORPS").Range.Paragraphs
        esp = esp & q.SpaceBefore & "/" & q.SpaceAfter & IIf(q.SpaceBeforeAuto Or q.SpaceAfterAuto, "auto", "") & " "
        If q.SpaceBeforeAuto Or q.SpaceAfterAuto Or Abs(q.SpaceBefore - 12) > 0.5 Or q.SpaceAfter > 0.5 Then espOk = False
    Next q
    modLog.Verifier "espacement des paragraphes (avant/apres pt)", espOk, esp
    s = ""
    For i = 1 To doc.Paragraphs.Count
        s = s & i & ":" & Round(doc.Paragraphs(i).LeftIndent / 28.35, 1) & "/" & Round(doc.Paragraphs(i).FirstLineIndent / 28.35, 1) & "[" & Left$(doc.Paragraphs(i).Range.Text, 12) & "] "
    Next i
    modLog.TestResultat "document complet apres correction", True, s
    doc.Close 0
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

Private Function Formats(ByVal rng As Range) As String
    Dim k As Long, s As String
    For k = 1 To rng.Paragraphs.Count
        s = s & Round(rng.Paragraphs(k).LeftIndent / 28.35, 1) & "/" & Round(rng.Paragraphs(k).FirstLineIndent / 28.35, 1) & " "
    Next k
    Formats = s
End Function

Private Function uniforme(ByVal rng As Range) As Boolean
    Dim k As Long
    uniforme = True
    For k = 2 To rng.Paragraphs.Count
        If Abs(rng.Paragraphs(k).LeftIndent - rng.Paragraphs(1).LeftIndent) > 0.5 Or _
           Abs(rng.Paragraphs(k).FirstLineIndent - rng.Paragraphs(1).FirstLineIndent) > 0.5 Then uniforme = False
    Next k
End Function

Public Sub Test_J2(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "J2 courrier sans ressaisie"
    On Error GoTo Echec
    Dim pat As Object, cor As Object, doc As Document, t As String

    Set pat = modBase.PatientParID("P00001")
    If Not modLog.Verifier("patient P00001 lu", Not pat Is Nothing) Then Exit Sub
    modLog.Verifier "identite patient", pat("Nom") = "FABREGUE" And pat("Prenom") = "Jean", pat("Nom") & " " & pat("Prenom")
    modLog.Verifier "medecin traitant renseigne", Len(pat("MedTraitantID")) > 0, pat("MedTraitantID")

    Set cor = modBase.CorrespondantParID("C0001")
    If Not modLog.Verifier("correspondant C0001 lu", Not cor Is Nothing) Then Exit Sub

    Set doc = modCourrier.CreerCourrierPour(pat, cor)
    t = doc.Content.Text
    modLog.Verifier "destinataire dans courrier", InStr(t, cor("Nom")) > 0
    modLog.Verifier "formule d'appel", InStr(t, cor("FormuleAppel")) > 0
    modLog.Verifier "jetons du modele tous remplaces", _
        InStr(t, "MTetMS") = 0 And InStr(t, "_LETTRE") = 0 And InStr(t, "<") = 0, t
    If doc.Bookmarks.Exists("POLITESSE") Then
        modLog.Verifier "formule de politesse", InStr(t, cor("FormulePolitesse")) > 0
    End If
    If doc.Bookmarks.Exists("CONCERNE") Then
        modLog.Verifier "ddn dans courrier (Concerne)", InStr(t, pat("DDN")) > 0, pat("DDN")
    End If
    If doc.Bookmarks.Exists("EXPEDITEUR") Then
        modLog.Verifier "expediteur (config)", InStr(t, modConfig.Config("MEDECIN", "Nom")) > 0
    End If
    modLog.Verifier "signet CORPS present", doc.Bookmarks.Exists("CORPS")
    modLog.Verifier "variable PatientID", doc.Variables("PatientID") = "P00001"
    modLog.Verifier "variable CorrespondantID", doc.Variables("CorrespondantID") = "C0001"
    ' identite patient sans dictee (Ctrl+Alt+P)
    Dim identite As String
    identite = modCourrier.TexteIdentitePatient(pat)
    modLog.Verifier "identite patient calculee", _
        InStr(identite, "Monsieur Jean FABREGUE, ") = 1 And InStr(identite, " ans") > 0, identite

    modCourrier.RemplacerCorps doc, "Ligne un du corps." & vbCr & "Ligne deux du corps." & vbCr & _
                                    "Ligne trois du corps." & vbCr & "Ligne quatre du corps."
    t = modCourrier.RecupererCorps(doc)
    modLog.Verifier "corps ecrit puis relu", InStr(t, "Ligne un du corps.") > 0 And InStr(t, "Ligne quatre du corps.") > 0
    modLog.Verifier "signet CORPS conserve apres ecriture", doc.Bookmarks.Exists("CORPS")
    ' mise en forme uniforme des paragraphes du corps (retraits identiques au 1er)
    Dim rc As Range, k As Long, uniforme As Boolean, detail As String
    Set rc = doc.Bookmarks("CORPS").Range
    uniforme = True
    For k = 1 To rc.Paragraphs.Count
        detail = detail & Round(rc.Paragraphs(k).LeftIndent / 28.35, 1) & "/" & Round(rc.Paragraphs(k).FirstLineIndent / 28.35, 1) & " "
        If Abs(rc.Paragraphs(k).LeftIndent - rc.Paragraphs(1).LeftIndent) > 0.5 Or _
           Abs(rc.Paragraphs(k).FirstLineIndent - rc.Paragraphs(1).FirstLineIndent) > 0.5 Then uniforme = False
    Next k
    modLog.Verifier "retraits uniformes dans le corps (cm gauche/1re ligne)", uniforme And rc.Paragraphs.Count >= 4, detail

    doc.Close 0
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

Public Sub Test_GRAS(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "GRAS : medicaments (majuscules) et expressions"
    On Error GoTo Echec
    Dim pat As Object, cor As Object, doc As Document, n As Long, n2 As Long, rng As Range, t As String
    modLog.Verifier "dictionnaire medicaments present", modFichiers.FichierExiste(modGras.FichierMedicaments()), modGras.FichierMedicaments()
    modLog.Verifier "dictionnaire expressions present", modFichiers.FichierExiste(modGras.FichierExpressions()), modGras.FichierExpressions()
    modGras.InvaliderCache
    modLog.Verifier "medicaments charges (>100)", modGras.NbMedicaments() > 100, modGras.NbMedicaments()
    modLog.Verifier "expressions chargees (>40)", modGras.NbExpressions() > 40, modGras.NbExpressions()
    Set pat = modBase.PatientParID("P00001")
    Set cor = modBase.CorrespondantParID("C0001")
    Set doc = modCourrier.CreerCourrierPour(pat, cor)
    modCourrier.RemplacerCorps doc, "Le patient prend de l'aspirine et de l'amiodarone depuis un an." & vbCr & _
        "Il presente une fibrillation auriculaire ancienne ; l'ECG du jour est sinusal, sans hta connue (le mot hypertension seul reste normal)."
    n = modGras.AppliquerGras(doc)
    modLog.Verifier "occurrences marquees (>=4)", n >= 4, n & " occurrence(s)"
    t = doc.Bookmarks("CORPS").Range.Text
    modLog.Verifier "medicaments en MAJUSCULES", InStr(t, "ASPIRINE") > 0 And InStr(t, "AMIODARONE") > 0, Left$(t, 80)
    Set rng = doc.Bookmarks("CORPS").Range
    rng.Find.Text = "ASPIRINE": rng.Find.MatchCase = False
    modLog.Verifier "ASPIRINE en gras", rng.Find.Execute And rng.Font.Bold = True
    Set rng = doc.Bookmarks("CORPS").Range
    rng.Find.Text = "fibrillation auriculaire"
    modLog.Verifier "expression en gras, casse conservee", rng.Find.Execute And rng.Font.Bold = True And rng.Text = "fibrillation auriculaire"
    Set rng = doc.Bookmarks("CORPS").Range
    rng.Find.Text = "hta"
    modLog.Verifier "HTA (mot entier) en gras", rng.Find.Execute And rng.Font.Bold = True
    Set rng = doc.Bookmarks("CORPS").Range
    rng.Find.Text = "patient"
    modLog.Verifier "mot hors dictionnaire non gras", rng.Find.Execute And rng.Font.Bold = False
    Set rng = doc.Bookmarks("CORPS").Range
    rng.Find.Text = "hypertension seul"
    modLog.Verifier "expression partielle non gras", rng.Find.Execute And rng.Font.Bold = False
    modLog.Verifier "signet CORPS conserve", doc.Bookmarks.Exists("CORPS") And doc.Bookmarks("CORPS").Range.Paragraphs.Count = 2
    n2 = modGras.AppliquerGras(doc)
    modLog.Verifier "second passage stable", n2 = n, n2
    doc.Close False
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

' Appel reel de l'API OpenAI (gpt-4.1) sur un texte FICTIF anonymise :
' meme circuit que Claude, seul le transport change.
Public Sub Test_OPENAI(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "OPENAI : appel reel via Responses API (texte anonymise)"
    On Error GoTo Echec
    Dim pat As Object, cor As Object, ctx As Object, anonyme As String, reponse As String, final As String
    modLog.Verifier "fournisseur par defaut = claude", modClaude.FournisseurApi() = "claude", modClaude.FournisseurApi()
    modLog.Verifier "cle OpenAI disponible", Len(modClaude.LireCleOpenAICabinet()) > 20
    Set pat = modBase.PatientParID("P00001")
    Set cor = modBase.CorrespondantParID("C0001")
    Set ctx = modAnonymise.Construire(pat, cor)
    anonyme = modAnonymise.Anonymiser("J'ai recu Monsieur Jean FABREGUE, ne le 01/01/1935, adresse par le Docteur GARRIGUE." & vbCr & _
        "Il prend de l'aspirine. La tension est a 140 sur 80 millimetres de mercure, le rythme est sinusal a 70 par minute." & vbCr & _
        "Au total, examen rassurant.", ctx)
    modLog.Verifier "pre-envoi : scan propre", Len(modAnonymise.ScanResiduel(anonyme, ctx)) = 0
    reponse = modClaude.AppelerClaude(modClaude.ChargerPrompt("correction.txt"), anonyme, "openai")
    modLog.Verifier "reponse OpenAI recue", Len(reponse) > 20, Left$(reponse, 100)
    modLog.Verifier "balises retour intactes", Len(modAnonymise.VerifierBalisesRetour(reponse, ctx)) = 0, Left$(reponse, 200)
    Dim payload As String
    payload = modFichiers.LireTexteUTF8(modConfig.chemin("Logs") & "\payload_debug.json")
    modLog.Verifier "payload OpenAI (model gpt-4.1)", InStr(payload, """model"":""gpt-4.1""") > 0
    modLog.Verifier "payload sans nom", InStr(1, payload, "FABREGUE", vbTextCompare) = 0
    modLog.Verifier "payload sans ddn", InStr(payload, "01/01/1935") = 0
    final = modAnonymise.Reinjecter(reponse, ctx)
    modLog.Verifier "final reinjecte", InStr(final, "FABREGUE") > 0, Left$(final, 160)
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

' Lettres derivees (hors ligne, sans appel API) : detection des demandes,
' profils, prompt, registre et destinataires pre-remplis
Public Sub Test_DERIVEES(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "DERIVEES demandes d'examen (hors ligne)"
    On Error GoTo Echec
    Dim prompt As String, cor As Object, liste As Collection, d As Object, p As Object, dem As Collection

    ' --- detection : declencheur + examen dans la meme phrase, sans exclusion ---
    Set dem = modDemandes.DetecterDemandes("Il est asymptomatique. Je prescris un test d'effort pour compléter ce bilan. " & _
        "Il avait réalisé une scintigraphie en 2019. Je l'adresse au Docteur X pour un avis pneumologique.")
    modLog.Verifier "2 demandes reperees", dem.Count = 2, CStr(dem.Count)
    If dem.Count >= 1 Then modLog.Verifier "test d'effort -> TEST_EFFORT", dem(1)("Code") = "TEST_EFFORT", dem(1)("Code")
    If dem.Count >= 2 Then modLog.Verifier "avis pneumologique -> profil", dem(2)("Code") = "AVIS_PNEUMOLOGIQUE", dem(2)("Code")
    If dem.Count >= 1 Then modLog.Verifier "phrase de prescription conservee", InStr(dem(1)("Phrase"), "Je prescris") > 0
    Set dem = modDemandes.DetecterDemandes("Nous poursuivons les investigations en réalisant une IRM de stress et un coroscanner.")
    modLog.Verifier "irm de stress prime sur irm", dem.Count = 2 And dem(1)("Code") = "IRM_DE_STRESS", CStr(dem.Count)
    modLog.Verifier "coroscanner et non scanner", dem.Count = 2 And dem(2)("Code") = "COROSCANNER"
    Set dem = modDemandes.DetecterDemandes("Le scanner thoracique de 2020 était normal. Un test d'effort pourrait être réalisé.")
    modLog.Verifier "sans declencheur ni avec exclusion : rien", dem.Count = 0, CStr(dem.Count)
    modLog.Verifier "motif avec *", modDemandes.MotifCorrespond("je complete le bilan cardiologique par une irm", "je complete le bilan*par*")

    ' --- profils ---
    Set p = modDemandes.ChargerProfil("TEST_EFFORT")
    modLog.Verifier "profil TEST_EFFORT charge", modDemandes.ValeurProfil(p, "IDENTITE", "LIBELLE") = "test d'effort"
    modLog.Verifier "profil : ordre des rubriques", InStr(modDemandes.ValeurProfil(p, "STRUCTURE", "ORDRE"), "ECG") > 0
    Set p = modDemandes.ChargerProfil("SCINTIGRAPHIE_MYOCARDIQUE")
    modLog.Verifier "profil scinti : 5 modalites", modDemandes.ModalitesProfil(p).Count = 5
    modLog.Verifier "texte de modalite", InStr(modDemandes.TexteExamen(p, "sous RAPISCAN"), "RAPISCAN") > 0
    Set p = modDemandes.ChargerProfil("CODE_INCONNU_XYZ")
    modLog.Verifier "profil inconnu -> AUTRE_EXAMEN", modDemandes.ValeurProfil(p, "IDENTITE", "CODE") = "AUTRE_EXAMEN"
    modLog.Verifier "liste des profils", modDemandes.ListerProfils().Count >= 35

    ' --- prompt : demande en tete, registre, pas de resume ---
    Set p = modDemandes.ChargerProfil("TEST_EFFORT")
    prompt = modDerivees.ConstruirePrompt(p, "", True, "Madame {{PAT_PRENOM}} {{PAT_NOM}}, 75 ans", "", "Je prescris un test d'effort.")
    modLog.Verifier "prompt : premiere phrase = demande", InStr(prompt, "Merci de réaliser un test d'effort à Madame {{PAT_PRENOM}} {{PAT_NOM}}, 75 ans, {MOTIF}") > 0 And InStr(prompt, "ne jamais inventer un motif") > 0
    modLog.Verifier "prompt : tutoiement", InStr(prompt, "Je te serais reconnaissant") > 0
    modLog.Verifier "prompt : rubrique ECG avec prefixe", InStr(prompt, "Électrocardiogramme") > 0 And InStr(prompt, "Sur le tracé en") > 0
    modLog.Verifier "prompt : derniere phrase = objectif du profil", InStr(prompt, "Merci de confirmer l'absence de coronaropathie.") > 0 And InStr(prompt, "Je te serais reconnaissant") = 0
    modLog.Verifier "prompt : phrase de prescription", InStr(prompt, "Je prescris un test d'effort.") > 0
    modLog.Verifier "prompt : interdit 'Je revois'", InStr(prompt, "Je revois") > 0 And InStr(prompt, "Au total") > 0
    modLog.Verifier "prompt : sans courriers de reference", InStr(prompt, "Courrier de référence") = 0 And InStr(prompt, "[COURRIERS_DE_REFERENCE]") = 0
    modLog.Verifier "prompt : sans marqueur restant", InStr(prompt, "{{CONSIGNES_TYPE}}") = 0 And InStr(prompt, "{{MOTIF}}") = 0 And InStr(prompt, "{ARTICLE_EXAMEN}") = 0
    Set p = modDemandes.ChargerProfil("HOSPITALISATION_CCN")
    prompt = modDerivees.ConstruirePrompt(p, "rapidement", False, "Monsieur {{PAT_PRENOM}} {{PAT_NOM}}, 68 ans", "", "")
    modLog.Verifier "prompt hosp : degre d'urgence insere", InStr(prompt, "Merci de prendre en charge rapidement Monsieur {{PAT_PRENOM}} {{PAT_NOM}}, 68 ans") > 0
    modLog.Verifier "prompt hosp : demande finale", InStr(prompt, "prendre en charge") > 0 And InStr(prompt, "{DEGRE_URGENCE}") = 0
    Set p = modDemandes.ChargerProfil("SCORE_CALCIQUE")
    prompt = modDerivees.ConstruirePrompt(p, "Contrôle", True, "Madame {{PAT_PRENOM}} {{PAT_NOM}}, 65 ans", "", "")
    modLog.Verifier "modalite = phrase complete", InStr(prompt, "Merci de réévaluer le score calcique de Madame") > 0

    ' --- registre et formules par defaut ---
    Set cor = CreateObject("Scripting.Dictionary"): cor.CompareMode = 1
    cor("FormuleAppel") = "": cor("Tutoiement") = "tu"
    modLog.Verifier "tutoiement explicite", modCourrier.EstTutoye(cor)
    modLog.Verifier "appel proche", modCourrier.AppelParDefaut(True) = "Cher Ami,"
    modLog.Verifier "politesse proche", modCourrier.PolitesseParDefaut(True) = "Bien cordialement."
    modLog.Verifier "politesse confrere", modCourrier.PolitesseParDefaut(False) = "Bien confraternellement."
    cor("Tutoiement") = "": cor("FormuleAppel") = "Mon Cher Jacques,"
    modLog.Verifier "tutoiement deduit de l'appel", modCourrier.EstTutoye(cor)
    cor("FormuleAppel") = "Cher Confrère,"
    modLog.Verifier "vouvoiement par defaut", Not modCourrier.EstTutoye(cor)

    ' --- destinataires pre-remplis (classeur des specialistes) ---
    Set liste = modDerivees.CorrespondantsPourTypes("TEST_EFFORT")
    modLog.Verifier "specialistes epreuve d'effort trouves", liste.Count > 0, "config [DERIVEES] FichierSpecialistes"
    If liste.Count > 0 Then
        Set d = liste(1)
        modLog.Verifier "priorite 1 en tete", d("Priorite") <= liste(liste.Count)("Priorite")
        modLog.Verifier "bloc destinataire present", Len(d("BlocDestinataire")) > 0, d("NomDestinataire")
        modLog.Verifier "formules renseignees", Len(d("FormuleAppel")) > 0 And Len(d("FormulePolitesse")) > 0
        modLog.Verifier "champs modCourrier presents", d.Exists("CP") And d.Exists("Ville") And d.Exists("Specialite")
    End If
    Set p = modDemandes.ChargerProfil("TEST_EFFORT")
    Set d = modDerivees.DestinataireAutomatique(p)
    modLog.Verifier "destinataire automatique", Not d Is Nothing And Len(d("BlocDestinataire")) > 0, d("NomDestinataire")
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub
