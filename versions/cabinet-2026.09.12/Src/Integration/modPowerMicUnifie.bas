Attribute VB_Name = "modPowerMicUnifie"
Option Explicit
Private mOccupe As Boolean

Public Sub Unifie_A_NouvelleLettre()
    modFileArrivees.Unifie_AfficherFileArrivees
End Sub

' Noms historiques appeles par les commandes Dragon PowerMic A et B.
Public Sub PowerMic_A_NouvelleLettre()
    Unifie_A_NouvelleLettre
End Sub

Public Sub PowerMic_B_FormuleAppel()
    Unifie_B_FormuleAppel
End Sub

Public Function Unifie_OperationEnCours() As Boolean
    Unifie_OperationEnCours = mOccupe Or modProdRapide.PR_EnCours() Or modCycleCourrier.CycleEnCours()
End Function

Public Function Unifie_DemarrerConsultation(ByVal attente As Object) As Boolean
    If mOccupe Or modProdRapide.PR_EnCours() Then Exit Function
    mOccupe = True
    On Error GoTo Erreur
    Dim pat As Object, doc As Document, sauvegarde As Boolean
    If attente Is Nothing Then GoTo Sortie
    Set pat = modBase.PatientParID(CStr(attente("PatientID")), True)
    If pat Is Nothing Then Err.Raise vbObjectError + 980, "modPowerMicUnifie", "Patient absent de la base NAS."
    modAttenteLocale.ConsommerAttente attente
    ' Relire apres la reservation atomique ; le serveur controle l identite affichee.
    Set pat = modBase.PatientParID(CStr(attente("PatientID")), True)
    If CStr(pat("Nom")) <> CStr(attente("Nom")) Or CStr(pat("Prenom")) <> CStr(attente("Prenom")) Or CStr(pat("DDN")) <> CStr(attente("DDN")) Or CStr(pat("Sexe")) <> CStr(attente("Sexe")) Then Err.Raise vbObjectError + 981, , "Identite modifiee : actualisez la liste."
    ' Le destinataire est dicte avec le raccourci Dragon, comme demande.
    Set doc = modCourrier.CreerCourrierRapidePour(pat)
    modIntegrationUnifie.FixerVariable doc, "RdvID", CStr(attente("RdvID"))
    modIntegrationUnifie.FixerVariable doc, "ConsultationID", CStr(attente("ConsultationID"))
    modIntegrationUnifie.FixerVariable doc, "AnneeAgenda", CStr(Year(modTexte.DateFr(CStr(attente("DateRdv")))))
    modIntegrationUnifie.FixerVariable doc, "DateActe", CStr(attente("DateRdv"))
    modIntegrationUnifie.FixerVariable doc, "ReservationNas", CStr(attente("ReservationNas"))
    modIntegrationUnifie.InitialiserPatientProd doc
    modIntegrationUnifie.SauvegarderBrouillon doc
    sauvegarde = True
    modGdt.EcrireGdtPatient pat
    doc.Activate
    modCourrier.AllerDestinataire
    Unifie_DemarrerConsultation = True
Sortie:
    mOccupe = False
    Exit Function
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
End Function

Public Sub Unifie_B_FormuleAppel()
    On Error GoTo Erreur
    ' Le bloc AutoTexte Dragon est conserve et lie ici au correspondant stable.
    Call modCourrier.AssocierDestinataireDicte(ActiveDocument)
    modCourrier.AllerAppel
    Exit Sub
Erreur:
    MsgBox "Destinataire non reconnu : " & Err.Description, vbExclamation, "Cabinet"
End Sub

Public Sub Unifie_C_InsererPatient()
    On Error GoTo Erreur
    Dim pat As Object, precedent As String
    If Documents.Count = 0 Then Err.Raise vbObjectError + 958, , "Aucun brouillon ouvert."
    If Selection.Start <> Selection.End Then Err.Raise vbObjectError + 958, , "Placez le curseur sans selection avant d inserer le patient."
    If Selection.Start > 0 Then precedent = ActiveDocument.Range(Selection.Start - 1, Selection.Start).Text
    Set pat = modIntegrationUnifie.PatientMemorise(ActiveDocument)
    Selection.TypeText modCourrier.PrefixeIdentitePatient(precedent) & modCourrier.TexteIdentitePatient(pat)
    Exit Sub
Erreur:
    MsgBox "Insertion du patient impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub

' Couper Dragon des l'entree dans D. Sans cette coupure, un bruit reconnu
' pendant la correction peut encore modifier le corps deja relu.
Private Function CouperMicrophoneDragon() As Boolean
    On Error GoTo Sortie
    Dim micro As Object
    Set micro = CreateObject("Dragon.MicBtn")
    micro.Register Empty
    micro.MicState = 0
    CouperMicrophoneDragon = (micro.MicState = 0)
Sortie:
    On Error Resume Next
    If Not micro Is Nothing Then micro.UnRegister
    Set micro = Nothing
End Function

Public Sub Unifie_D_Finaliser()
    If Not CouperMicrophoneDragon() Then
        MsgBox "Le microphone Dragon n a pas pu etre coupe. Aucun envoi ; coupez-le et reprenez D.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    If mOccupe Or modProdRapide.PR_EnCours() Then Exit Sub
    mOccupe = True
    On Error GoTo Erreur
    modIntegrationUnifie.InitialiserPatientProd ActiveDocument
    ' Le medecin relit avant D. Le cycle corrige et publie seulement si
    ' tous les controles bloquants sont conformes.
    modProdRapide.PR_CorrigerToutEnUnClic
Sortie:
    mOccupe = False
    Exit Sub
Erreur:
    MsgBox "Finalisation interrompue : " & Err.Description, vbExclamation, "Cabinet"
    Resume Sortie
End Sub
