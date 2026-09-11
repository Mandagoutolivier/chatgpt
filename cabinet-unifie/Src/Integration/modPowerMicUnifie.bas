Attribute VB_Name = "modPowerMicUnifie"
Option Explicit

' Points d'entree uniques a affecter aux quatre touches du PowerMic.
Public Sub Unifie_A_NouvelleLettre()
    On Error GoTo Erreur
    Dim attente As Object, pat As Object, cor As Object, doc As Document
    Set attente = modAttenteLocale.ChoisirAttente()
    If attente Is Nothing Then
        Cabinet_A_NouvelleLettre
        Exit Sub
    End If
    Set pat = modBase.PatientParID(attente("PatientID"))
    If pat Is Nothing Then Err.Raise vbObjectError + 980, , "Patient absent de la base NAS."
    If Len(pat("MedTraitantID")) > 0 Then Set cor = modBase.CorrespondantParID(pat("MedTraitantID"))
    If cor Is Nothing Then
        Set doc = modCourrier.CreerCourrierRapidePour(pat)
    Else
        Set doc = modCourrier.CreerCourrierPour(pat, cor)
    End If
    doc.Variables("RdvID") = attente("RdvID")
    modGdt.EcrireGdtPatient pat
    modAttenteLocale.ConsommerAttente attente
    doc.Activate
    modCourrier.AllerDestinataire
    Exit Sub
Erreur:
    MsgBox "Ouverture du courrier impossible : " & Err.Description, vbExclamation, "Cabinet unifie"
End Sub

Public Sub Unifie_B_FormuleAppel()
    Cabinet_B_FormuleAppel
End Sub

Public Sub Unifie_C_InsererPatient()
    Cabinet_P_Patient
End Sub

Public Sub Unifie_D_Finaliser()
    PR_CorrigerToutEnUnClic
End Sub
