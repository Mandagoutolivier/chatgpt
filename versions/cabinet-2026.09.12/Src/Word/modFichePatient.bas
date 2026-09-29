Attribute VB_Name = "modFichePatient"
Option Explicit

Public Sub Unifie_FichePatient()
    Dim doc As Document, pat As Object, f As ufPatientMedecin, id As String
    If modPowerMicUnifie.Unifie_OperationEnCours() Then Exit Sub
    On Error GoTo Echec
    If Documents.Count = 0 Then Err.Raise vbObjectError + 1291, , "Ouvrez le courrier du patient depuis les patients arrives ou les brouillons."
    Set doc = ActiveDocument
    id = Trim$(modIntegrationUnifie.VariableDoc(doc, "PatientID"))
    If Len(id) = 0 Or Len(Trim$(modIntegrationUnifie.VariableDoc(doc, "ConsultationID"))) = 0 Then Err.Raise vbObjectError + 1291, , "Ce document n est pas rattache a une consultation."
    ' Lecture directe : la fiche doit rester accessible pour corriger une
    ' identite incomplete ou devenue differente de celle du brouillon.
    Set pat = modBase.PatientParID(id, True)
    Set f = New ufPatientMedecin
    f.Charger doc, pat
    f.Show vbModal
Sortie:
    On Error Resume Next
    If Not f Is Nothing Then Unload f
    If Not doc Is Nothing Then doc.Activate
    Exit Sub
Echec:
    MsgBox "Fiche patient : " & Err.Description, vbExclamation, "Cabinet"
    Resume Sortie
End Sub

Public Function IdentiteCourrierDifferente(ByVal doc As Document, ByVal pat As Object) As Boolean
    Dim k As Variant
    For Each k In Array("Nom", "Prenom", "DDN", "Sexe")
        If StrComp(Trim$(modIntegrationUnifie.VariableDoc(doc, "Patient_" & k)), CStr(pat(k)), vbBinaryCompare) <> 0 Then
            IdentiteCourrierDifferente = True
            Exit Function
        End If
    Next k
End Function

' Appel uniquement apres confirmation explicite dans la fiche. Aucun
' remplacement aveugle de noms, d ages ou de pronoms dans le texte clinique.
Public Sub ConfirmerIdentiteCourrier(ByVal doc As Document, ByVal pat As Object)
    Dim actuel As Object, k As Variant
    If doc.ReadOnly Then Err.Raise vbObjectError + 1291, , "Le courrier est en lecture seule."
    If Trim$(modIntegrationUnifie.VariableDoc(doc, "PatientID")) <> CStr(pat("ID")) Then Err.Raise vbObjectError + 1291, , "Patient du courrier different."
    Set actuel = modBase.PatientParID(CStr(pat("ID")), True)
    If CStr(actuel("_revision")) <> CStr(pat("_revision")) Then Err.Raise vbObjectError + 1291, , "Fiche modifiee sur un autre poste : fermez puis rechargez la fiche."
    If MsgBox("Confirmez-vous avoir verifie et, si necessaire, corrige l identite, l age et les accords dans le courrier ET ses annexes ?" & vbCrLf & vbCrLf & _
        CStr(actuel("Nom")) & " " & CStr(actuel("Prenom")) & vbCrLf & _
        "Naissance : " & CStr(actuel("DDN")) & " - Sexe : " & CStr(actuel("Sexe")) & vbCrLf & _
        "Une nouvelle correction et relecture seront requises avant transmission.", vbYesNo + vbQuestion + vbDefaultButton2, "Identite du courrier") <> vbYes Then Exit Sub
    modControleCourrier.InvaliderRelecture doc
    modIntegrationUnifie.FixerVariable doc, "CycleU1", ""
    modIntegrationUnifie.FixerVariable doc, "PublicationRevision", ""
    For Each k In Array("Nom", "Prenom", "DDN", "Sexe")
        modIntegrationUnifie.FixerVariable doc, "Patient_" & k, CStr(actuel(k))
    Next k
    modIntegrationUnifie.InitialiserPatientProd doc
    doc.Save
End Sub
