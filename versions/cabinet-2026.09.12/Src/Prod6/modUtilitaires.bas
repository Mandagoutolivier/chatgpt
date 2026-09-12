Attribute VB_Name = "modUtilitaires"
Option Explicit

Public Function NettoyerTexteParagraphe(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(11), "")
    s = Replace(s, Chr(12), "")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, " ")

    s = Trim(s)

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NettoyerTexteParagraphe = s

End Function

Public Function CommencePar(ByVal texte As String, ByVal debut As String) As Boolean

    texte = LCase(NettoyerTexteParagraphe(texte))
    debut = LCase(NettoyerTexteParagraphe(debut))

    CommencePar = (Left$(texte, Len(debut)) = debut)

End Function

Public Sub CopierTexteSimpleDansPressePapiers(ByVal texte As String)

    Dim docTemp As Document

    Set docTemp = Documents.Add(Visible:=False)
    docTemp.Range.Text = texte
    docTemp.Range.Copy
    docTemp.Close SaveChanges:=wdDoNotSaveChanges

End Sub

Public Function LireTexteDepuisPressePapiersWord() As String

    Dim docTemp As Document
    Dim rngTemp As Range
    Dim texte As String

    LireTexteDepuisPressePapiersWord = ""

    On Error GoTo ErreurLecture

    Set docTemp = Documents.Add(Visible:=False)

    Set rngTemp = docTemp.Range
    rngTemp.PasteSpecial DataType:=wdPasteText

    texte = docTemp.Range.Text

    'Retirer le dernier marqueur de paragraphe ajouté par Word.
    If Len(texte) > 0 Then
        If Right$(texte, 1) = vbCr Or Right$(texte, 1) = vbLf Then
            texte = Left$(texte, Len(texte) - 1)
        End If
    End If

    docTemp.Close SaveChanges:=wdDoNotSaveChanges

    LireTexteDepuisPressePapiersWord = texte

    Exit Function

ErreurLecture:

    On Error Resume Next

    If Not docTemp Is Nothing Then
        docTemp.Close SaveChanges:=wdDoNotSaveChanges
    End If

    LireTexteDepuisPressePapiersWord = ""

End Function
