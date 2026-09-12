Attribute VB_Name = "modBase"
Option Explicit
' Lectures ciblees par HTTPS. Pas de copie locale de Patients.xlsx.
Public Function patients(Optional ByVal forcer As Boolean = False) As Collection
    Set patients = modServiceNas.LireTable("PATIENTS")
End Function
Public Function Correspondants(Optional ByVal forcer As Boolean = False) As Collection
    Set Correspondants = modServiceNas.LireTable("CORRESPONDANTS")
End Function
Public Function PatientParID(ByVal id As String, Optional ByVal forcer As Boolean = False) As Object
    Set PatientParID = modServiceNas.LireID("PATIENTS", id)
End Function
Public Function CorrespondantParID(ByVal id As String) As Object
    Set CorrespondantParID = modServiceNas.LireID("CORRESPONDANTS", id)
End Function
Public Sub ChargerBases(Optional ByVal forcer As Boolean = False)
    Dim r As Object
    Set r = modServiceNas.Appeler("whoami", modServiceNas.Parametres())
End Sub
