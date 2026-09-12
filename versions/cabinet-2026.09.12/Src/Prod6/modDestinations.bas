Attribute VB_Name = "modDestinations"
Option Explicit

Public Function CT_DestinationParCle(ByVal cleDestination As String) As String
    Dim p As Object, r As Object
    Set p = modServiceNas.Parametres(): p("cle") = cleDestination
    Set r = modServiceNas.Appeler("correspondent.resolve", p)
    CT_DestinationParCle = modBaseCorrespondants.RecordCorrespondant(r, CStr(r("TypesExamen")))
End Function

Public Function CT_ListeDestinationsPourSelection() As Collection
    Dim col As New Collection, cor As Object, d As Object
    For Each cor In modBase.Correspondants()
        If CStr(cor("Actif")) <> "0" And CStr(cor("AValider")) <> "1" Then
            Set d = modServiceNas.Parametres()
            d("Cle") = CStr(cor("ID")): d("Nom") = CStr(cor("Nom")) & " " & CStr(cor("Prenom"))
            d("Type") = CStr(cor("TypesExamen")): d("Structure") = CStr(cor("StructureID")): d("Ville") = CStr(cor("Ville"))
            col.Add d
        End If
    Next cor
    Set CT_ListeDestinationsPourSelection = col
End Function
