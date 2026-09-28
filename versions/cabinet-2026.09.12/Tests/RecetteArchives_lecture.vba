
' Injecte uniquement pendant la recette ; retire avant livraison.
Private Sub U2AssertArchive(ByVal ok As Boolean, ByVal nom As String, ByRef n As Long)
    If Not ok Then Err.Raise vbObjectError + 1910, , "ECHEC " & nom
    n = n + 1
End Sub
Private Function U2RefusArchive(ByVal d As Object) As Boolean
    Dim p As String
    On Error GoTo Refus
    p = CopieLectureArchive(d)
    Exit Function
Refus:
    U2RefusArchive = (Err.Number = vbObjectError + 1116)
End Function
Public Function U2TesterArchives(ByVal racineEssai As String, ByVal shaDocx As String, ByVal shaPdf As String) As String
    Dim d As Object, p As String, cache As String, fso As Object
    Dim w As Object, doc As Object, n As Long, securite As Long, liens As Boolean, nb As Long
    Dim numero As Long, description As String, ff As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Echec
    modConfig.DefinirRacine racineEssai
    Set d = CreateObject("Scripting.Dictionary")
    d("CheminDocx") = racineEssai & "\Documents\" & shaDocx & ".docx"
    d("sha_docx") = shaDocx
    d("CheminPdf") = racineEssai & "\Documents\" & shaPdf & ".pdf"
    d("sha_pdf") = shaPdf
    p = CopieLectureArchive(d): cache = p
    U2AssertArchive InStr(1, p, Environ$("LOCALAPPDATA") & "\CabinetCardio\LecturesArchives\", vbTextCompare) = 1, "copie locale", n
    U2AssertArchive modDonneesTransport.EmpreinteFichierSHA256(p) = shaDocx, "copie integre", n
    U2AssertArchive (GetAttr(p) And vbReadOnly) <> 0, "attribut lecture seule", n
    Set w = GetObject(, "Word.Application")
    securite = w.AutomationSecurity: liens = w.Options.UpdateLinksAtOpen
    nb = w.Documents.Count
    OuvrirCourrier d
    Set doc = w.ActiveDocument
    U2AssertArchive doc.ReadOnly And StrComp(doc.FullName, p, vbTextCompare) = 0, "Word local lecture seule", n
    U2AssertArchive w.AutomationSecurity = securite And w.Options.UpdateLinksAtOpen = liens, "options Word restaurees", n
    OuvrirCourrier d
    U2AssertArchive w.Documents.Count = nb + 1, "reouverture sans doublon", n
    doc.Close 0: Set doc = Nothing
    d("sha_docx") = "invalide"
    U2AssertArchive U2RefusArchive(d), "empreinte invalide refusee", n
    d.Remove "sha_docx"
    U2AssertArchive U2RefusArchive(d), "empreinte absente refusee", n
    d("sha_docx") = shaDocx
    d("CheminDocx") = racineEssai & "\Documents\..\Documents\" & shaDocx & ".docx"
    U2AssertArchive U2RefusArchive(d), "traversee de chemin refusee", n
    d("CheminDocx") = racineEssai & "\Documents\" & shaDocx & ".docx"
    SetAttr cache, vbNormal
    ff = FreeFile: Open cache For Append As #ff: Print #ff, "ALTERATION RECETTE": Close #ff
    U2AssertArchive U2RefusArchive(d), "copie locale alteree refusee", n
    fso.DeleteFile cache, True
    p = CopieLectureArchive(d)
    ff = FreeFile: Open CStr(d("CheminDocx")) For Append As #ff: Print #ff, "ALTERATION RECETTE": Close #ff
    U2AssertArchive U2RefusArchive(d), "archive source alteree refusee", n
    d.Remove "CheminDocx"
    p = CopieLectureArchive(d)
    U2AssertArchive LCase$(Right$(p, 4)) = ".pdf" And modDonneesTransport.EmpreinteFichierSHA256(p) = shaPdf And (GetAttr(p) And vbReadOnly) <> 0, "repli PDF local controle", n
    modConfig.DefinirRacine ""
    U2TesterArchives = CStr(n) & "/12 PASS"
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    Close #ff
    If Not doc Is Nothing Then doc.Close 0
    modConfig.DefinirRacine ""
    On Error GoTo 0
    U2TesterArchives = "FAIL " & CStr(numero) & " " & description
End Function
