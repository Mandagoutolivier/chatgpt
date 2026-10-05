Attribute VB_Name = "modIndexAnnexes"
Option Explicit
' Reperes nommes durables, independants de la pagination Word.
Public Sub FixerMeta(ByVal doc As Object, ByVal nom As String, ByVal valeur As String)
    Dim v As Object
    For Each v In doc.Variables
        If StrComp(v.Name, nom, vbTextCompare) = 0 Then
            v.Value = valeur
            Exit Sub
        End If
    Next v
    doc.Variables.Add nom, valeur
End Sub

Public Function LireMeta(ByVal doc As Object, ByVal nom As String) As String
    On Error Resume Next
    LireMeta = CStr(doc.Variables(nom).Value)
End Function

Public Sub IndexerAnnexe(ByVal doc As Object, ByVal source As Object, ByVal debut As Long, ByVal numero As Long, ByVal destinataireID As String)
    Dim r As Object, suffixe As String, offset As Long, longueur As Long
    If numero < 1 Or numero > 999 Or Len(Trim$(destinataireID)) = 0 Then Err.Raise vbObjectError + 1200, , "Index d annexe invalide."
    If Not source.Bookmarks.Exists("U2_DEST_SOURCE") Then Err.Raise vbObjectError + 1200, , "Repere destinataire de l annexe absent."
    Set r = source.Bookmarks("U2_DEST_SOURCE").Range
    offset = r.Start - source.Content.Start
    longueur = r.End - r.Start
    suffixe = Format$(numero, "000")
    If debut + offset + longueur > doc.Content.End - 1 Then Err.Raise vbObjectError + 1200, , "Destinataire hors de l annexe."
    doc.Bookmarks.Add "U2ANN_DEST_" & suffixe, doc.Range(debut + offset, debut + offset + longueur)
    FixerMeta doc, "AnnexeDestinataire_" & suffixe, destinataireID
    If doc.Bookmarks.Exists("U2_DEST_SOURCE") Then doc.Bookmarks("U2_DEST_SOURCE").Delete
End Sub

Public Function ListerAnnexes(ByVal doc As Object) As Collection
    Dim result As New Collection, b As Object, it As Object, suffixe As String
    For Each b In doc.Bookmarks
        If b.Name Like "U2ANN_DEST_###" Then
            suffixe = Right$(b.Name, 3)
            Set it = CreateObject("Scripting.Dictionary")
            it("ID") = suffixe
            it("Annexe") = "Annexe " & CStr(CLng(suffixe))
            it("Destinataire") = Replace(Replace(b.Range.Text, Chr$(11), " "), vbCr, " ")
            it("DestinataireID") = LireMeta(doc, "AnnexeDestinataire_" & suffixe)
            result.Add it
        End If
    Next b
    Set ListerAnnexes = result
End Function

Public Sub RemplacerDestinataireAnnexe(ByVal doc As Object, ByVal numero As String, ByVal correspondant As Object)
    Dim nom As String, r As Object, police As Object, para As Object, texte As String
    If Not numero Like "###" Then Err.Raise vbObjectError + 1201, , "Numero d annexe invalide."
    nom = "U2ANN_DEST_" & numero
    If Not doc.Bookmarks.Exists(nom) Then Err.Raise vbObjectError + 1201, , "Repere de l annexe introuvable."
    If Len(Trim$(CStr(correspondant("ID")))) = 0 Or CStr(correspondant("Actif")) <> "1" Or CStr(correspondant("AValider")) <> "0" Then Err.Raise vbObjectError + 1201, , "Correspondant non valide."
    texte = CStr(correspondant("BlocDestinataire"))
    If Len(Trim$(texte)) = 0 Then Err.Raise vbObjectError + 1201, , "Adresse vide."
    texte = Replace(Replace(Replace(texte, vbCrLf, Chr$(11)), vbCr, Chr$(11)), vbLf, Chr$(11))
    Set r = doc.Bookmarks(nom).Range
    Set police = r.Font.Duplicate
    Set para = r.ParagraphFormat.Duplicate
    r.Text = texte
    r.Font = police
    r.ParagraphFormat = para
    doc.Bookmarks.Add nom, r
    FixerMeta doc, "AnnexeDestinataire_" & numero, CStr(correspondant("ID"))
    r.Select
End Sub
