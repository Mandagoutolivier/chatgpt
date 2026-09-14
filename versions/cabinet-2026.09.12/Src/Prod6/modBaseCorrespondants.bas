Attribute VB_Name = "modBaseCorrespondants"
Option Explicit

Public Function PROD_FichierExiste(ByVal chemin As String) As Boolean
    PROD_FichierExiste = modFichiers.FichierExiste(chemin)
End Function

Public Function RecordCorrespondant(ByVal cor As Object, Optional ByVal examen As String = "") As String
    Const SEP As String = "§§"
    RecordCorrespondant = CStr(cor("ID")) & SEP & examen & SEP & CStr(cor("Nom")) & " " & CStr(cor("Prenom")) & SEP & _
        CStr(cor("StructureID")) & SEP & CStr(cor("BlocDestinataire")) & SEP & CStr(cor("FormuleAppel")) & SEP & _
        CStr(cor("Tutoiement")) & SEP & "1" & SEP & CStr(cor("CP")) & SEP & CStr(cor("Ville")) & SEP & _
        CStr(cor("ID")) & SEP & CStr(cor("FormulePolitesse"))
End Function

Public Function DestinationParDefautPourTypeExamen(ByVal typeExamen As String) As String
    Dim p As Object, r As Object
    On Error GoTo Echec
    Set p = modServiceNas.Parametres(): p("examen") = typeExamen
    Set r = modServiceNas.Appeler("correspondent.resolve", p)
    DestinationParDefautPourTypeExamen = RecordCorrespondant(r, typeExamen)
    Exit Function
Echec:
    If Err.Number = vbObjectError + 1141 Then Exit Function
    Dim numero As Long, description As String
    numero = Err.Number: description = Err.Description
    On Error GoTo 0
    Err.Raise numero, "Destinations", description
End Function

Public Function RechercherDestinationsParTypeExamen(ByVal typeExamen As String) As String
    Dim cor As Object, texte As String
    For Each cor In modBase.Correspondants()
        If InStr(1, ";" & CStr(cor("TypesExamen")) & ";", ";" & typeExamen & ";", vbTextCompare) > 0 Then
            texte = texte & CStr(cor("ID")) & " : " & CStr(cor("BlocDestinataire")) & vbCrLf
        End If
    Next cor
    RechercherDestinationsParTypeExamen = texte
End Function

Public Function DecrireDestinationParDefaut(ByVal recordDestination As String) As String
    DecrireDestinationParDefaut = ChampDestination(recordDestination, 2) & " - " & ChampDestination(recordDestination, 9)
End Function

Public Function ChampDestination(ByVal recordDestination As String, _
                                 ByVal indexChamp As Long) As String

    Dim champs() As String

    Const SEP As String = "§§"

    champs = Split(recordDestination, SEP)

    If indexChamp < LBound(champs) Or indexChamp > UBound(champs) Then
        ChampDestination = ""
    Else
        ChampDestination = champs(indexChamp)
    End If

End Function
