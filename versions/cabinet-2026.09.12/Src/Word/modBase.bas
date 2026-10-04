Attribute VB_Name = "modBase"
Option Explicit
Private mCorrespondants As Collection

' Lectures ciblees par HTTPS. Pas de copie locale de Patients.xlsx.
Public Function patients(Optional ByVal forcer As Boolean = False) As Collection
    Set patients = modServiceNas.LireTable("PATIENTS")
End Function

Public Function Correspondants(Optional ByVal forcer As Boolean = False) As Collection
    If forcer Or mCorrespondants Is Nothing Then
        Set mCorrespondants = modServiceNas.LireTable("CORRESPONDANTS")
    End If
    Set Correspondants = mCorrespondants
End Function

Public Function PatientParID(ByVal id As String, Optional ByVal forcer As Boolean = False) As Object
    Set PatientParID = modServiceNas.LireID("PATIENTS", id)
End Function

Public Function CorrespondantParID(ByVal id As String) As Object
    Set CorrespondantParID = modServiceNas.LireID("CORRESPONDANTS", id)
End Function

Public Sub ChargerBases(Optional ByVal forcer As Boolean = False)
    Dim chemin As String, liste As Object, item As Variant
    chemin = Environ$("LOCALAPPDATA") & "\CabinetCardio\correspondants-session.json"
    If Not modFichiers.FichierExiste(chemin) Then Err.Raise vbObjectError + 1168, , "Cache de correspondants de session absent."
    Set liste = modJson.JsonParse(modFichiers.LireTexteUTF8(chemin))
    If TypeName(liste) <> "Collection" Then Err.Raise vbObjectError + 1168, , "Cache de correspondants de session invalide."
    For Each item In liste
        If TypeName(item) <> "Dictionary" Then Err.Raise vbObjectError + 1168, , "Correspondant de session invalide."
        If Not item.Exists("ID") Or Not item.Exists("BlocDestinataire") Then Err.Raise vbObjectError + 1168, , "Correspondant de session incomplet."
    Next item
    Set mCorrespondants = liste
End Sub
