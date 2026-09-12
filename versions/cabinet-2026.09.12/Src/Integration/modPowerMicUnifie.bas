Attribute VB_Name = "modPowerMicUnifie"
Option Explicit
Private mOccupe As Boolean

Public Sub Unifie_A_NouvelleLettre()
    If mOccupe Or modProdRapide.PR_EnCours() Then Exit Sub
    mOccupe = True
    On Error GoTo Erreur
    Dim attente As Object, pat As Object, doc As Document, sauvegarde As Boolean
    Set attente = modAttenteLocale.ChoisirAttente()
    If attente Is Nothing Then GoTo Sortie
    Set pat = modBase.PatientParID(CStr(attente("PatientID")), True)
    If pat Is Nothing Then Err.Raise vbObjectError + 980, "modPowerMicUnifie", "Patient absent de la base NAS."
    modAttenteLocale.ConsommerAttente attente
    ' Le destinataire est dicte avec le raccourci Dragon, comme demande.
    Set doc = modCourrier.CreerCourrierRapidePour(pat)
    modIntegrationUnifie.FixerVariable doc, "RdvID", CStr(attente("RdvID"))
    modIntegrationUnifie.FixerVariable doc, "ConsultationID", CStr(attente("ConsultationID"))
    modIntegrationUnifie.FixerVariable doc, "AnneeAgenda", CStr(Year(modTexte.DateFr(CStr(attente("DateRdv")))))
    modIntegrationUnifie.FixerVariable doc, "ReservationNas", CStr(attente("ReservationNas"))
    modIntegrationUnifie.InitialiserPatientProd doc
    modIntegrationUnifie.SauvegarderBrouillon doc
    sauvegarde = True
    modAttenteLocale.EnregistrerBrouillon attente, doc.FullName
    modGdt.EcrireGdtPatient pat
    doc.Activate
    modCourrier.AllerDestinataire
Sortie:
    mOccupe = False
    Exit Sub
Erreur:
    Dim description As String
    description = Err.Description
    On Error Resume Next
    If Not sauvegarde Then
        If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
        modAttenteLocale.LibererReservation attente
    End If
    On Error GoTo 0
    mOccupe = False
    MsgBox "Ouverture interrompue : " & description & IIf(sauvegarde, vbCrLf & "Le brouillon reste ouvert et enregistre sur le NAS.", ""), vbExclamation, "Cabinet"
End Sub

Public Sub Unifie_B_FormuleAppel()
    modCourrier.AllerAppel
End Sub

Public Sub Unifie_C_InsererPatient()
    On Error GoTo Erreur
    Dim pat As Object
    Set pat = modIntegrationUnifie.PatientVerifie(ActiveDocument)
    Selection.TypeText modCourrier.TexteIdentitePatient(pat)
    Exit Sub
Erreur:
    MsgBox "Insertion du patient impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub
Public Sub Unifie_D_Finaliser()
    If mOccupe Or modProdRapide.PR_EnCours() Then Exit Sub
    mOccupe = True
    On Error GoTo Erreur
    If Trim$(modIntegrationUnifie.VariableDoc(ActiveDocument, "RelectureEnAttente")) = "1" Then
        modControleCourrier.ValiderEtTransmettre ActiveDocument
    Else
        modProdRapide.PR_CorrigerToutEnUnClic
    End If
Sortie:
    mOccupe = False
    Exit Sub
Erreur:
    MsgBox "Finalisation interrompue : " & Err.Description, vbExclamation, "Cabinet"
    Resume Sortie
End Sub
