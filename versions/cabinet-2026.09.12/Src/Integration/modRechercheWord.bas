Attribute VB_Name = "modRechercheWord"
Option Explicit

' Ne pas effacer les criteres de police ou de remplacement de l utilisateur.
Public Function Trouver(ByVal zone As Range, ByVal texte As String, Optional ByVal motEntier As Boolean = False, Optional ByVal versAvant As Boolean = True) As Boolean
    Dim recherche As Find, etat As Object, cle As Variant, noms As Variant
    Dim numero As Long, description As String, erreurRestauration As Long
    On Error GoTo Echec
    Set recherche = zone.Find
    Set etat = CreateObject("Scripting.Dictionary")
    noms = Array("Text", "Forward", "Wrap", "MatchCase", "MatchWholeWord", "MatchWildcards", "MatchSoundsLike", "MatchAllWordForms", "MatchPrefix", "MatchSuffix", "IgnoreSpace", "IgnorePunct", "Format")
    For Each cle In noms: etat(CStr(cle)) = CallByName(recherche, CStr(cle), VbGet): Next cle
    recherche.MatchPrefix = False: recherche.MatchSuffix = False
    recherche.IgnoreSpace = False: recherche.IgnorePunct = False
    Trouver = recherche.Execute(FindText:=texte, MatchCase:=False, MatchWholeWord:=motEntier, MatchWildcards:=False, MatchSoundsLike:=False, MatchAllWordForms:=False, Forward:=versAvant, Wrap:=wdFindStop, Format:=False)
Sortie:
    On Error Resume Next
    If Not etat Is Nothing Then
        For Each cle In etat.Keys
            Err.Clear
            CallByName recherche, CStr(cle), VbLet, etat(CStr(cle))
            If Err.Number <> 0 Then erreurRestauration = Err.Number
        Next cle
    End If
    On Error GoTo 0
    If numero <> 0 Then Err.Raise numero, "Recherche Word", description
    If erreurRestauration <> 0 Then Err.Raise vbObjectError + 1171, , "Options de recherche Word non restaurees."
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    Resume Sortie
End Function
