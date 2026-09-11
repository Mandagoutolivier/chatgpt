Attribute VB_Name = "modUI"
Option Explicit
' =====================================================================
' modUI - Actions des boutons de la feuille Accueil (poste secretaire).
' =====================================================================

Public Sub UI_NouveauPatient()
    On Error GoTo Erreur
    Dim f As ufPatientEdit
    Set f = New ufPatientEdit
    f.ChargerNouveau
    f.Show vbModal
    Unload f
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_ModifierPatient()
    On Error GoTo Erreur
    Dim pat As Object, f As ufPatientEdit
    Set pat = ChoisirPatientX()
    If pat Is Nothing Then Exit Sub
    Set f = New ufPatientEdit
    f.ChargerExistant pat
    f.Show vbModal
    Unload f
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_NouveauRdv()
    On Error GoTo Erreur
    Dim f As ufRdvEdit
    Set f = New ufRdvEdit
    f.Show vbModal
    Unload f
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_ArriveePatient()
    On Error GoTo Erreur
    Dim rdvs As Collection, f As ufListe, r As Object
    Set rdvs = modAgenda.RdvDuJour()
    If rdvs.Count = 0 Then
        MsgBox "Aucun rendez-vous aujourd'hui." & vbCrLf & _
               "Utilisez 'Prise de rendez-vous' pour en creer un.", vbInformation, "Cabinet"
        Exit Sub
    End If
    Set f = New ufListe
    f.Configurer "Arrivee d'un patient - RDV du jour", rdvs, _
                 Array("Heure", "Nom", "Prenom", "TypeActe", "Statut"), "40 pt;110 pt;90 pt;70 pt;60 pt"
    f.Show vbModal
    If f.Annule Then Unload f: Exit Sub
    Set r = f.Resultat
    Unload f
    modAgenda.MarquerStatut r("ID"), "Arrive"
    ' publication NAS ; le poste medecin alimente ensuite son cache SQLite et le GDT
    Dim noteEcg As String, p As Object, patientTrouve As Boolean
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        If p("ID") = r("PatientID") Then
            modEchange.PublierArrivee r, p
            noteEcg = vbCrLf & "Identite publiee pour le poste medecin et l'ECG."
            patientTrouve = True
            Exit For
        End If
    Next p
    If Not patientTrouve Then Err.Raise vbObjectError + 920, "modUI", "Patient du rendez-vous introuvable."
    MsgBox r("Prenom") & " " & r("Nom") & " marque ARRIVE a " & Format$(Now, "hh:nn") & "." & noteEcg, _
           vbInformation, "Cabinet"
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_Correspondants()
    On Error GoTo Erreur
    Dim f As ufListe, fc As ufCorrespEdit
    Set f = New ufListe
    f.Configurer "Correspondants (Nouveau... pour en creer un)", _
                 modBaseIO.LireTableX(modConfig.FichierPatients(), "CORRESPONDANTS"), _
                 Array("Titre", "Nom", "Prenom", "Specialite", "Ville"), _
                 "30 pt;110 pt;80 pt;120 pt;80 pt", "", True
    f.Show vbModal
    If f.NouveauDemande Then
        Unload f
        Set fc = New ufCorrespEdit
        fc.ChargerNouveau
        fc.Show vbModal
        Unload fc
    ElseIf Not f.Annule Then
        Dim cor As Object
        Set cor = f.Resultat
        Unload f
        Set fc = New ufCorrespEdit
        fc.ChargerExistant cor
        fc.Show vbModal
        Unload fc
    Else
        Unload f
    End If
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_CourriersATraiter()
    On Error GoTo Erreur
    Dim courriers As Collection, f As ufListe, d As Object, fa As ufChoixActe
    Set courriers = modEchange.CourriersEnAttente()
    If courriers.Count = 0 Then
        MsgBox "Aucun courrier en attente.", vbInformation, "Cabinet"
        Exit Sub
    End If
    Set f = New ufListe
    f.Configurer "Courriers valides par le medecin", courriers, _
                 Array("Nom", "Prenom", "TypeCourrier", "DateValidation"), "110 pt;90 pt;110 pt;90 pt"
    f.Show vbModal
    If f.Annule Then Unload f: Exit Sub
    Set d = f.Resultat
    Unload f
    Set fa = New ufChoixActe
    fa.Charger d
    fa.Show vbModal
    Unload fa
    modEchange.VerifierEchange
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_Agenda()
    On Error GoTo Erreur
    modAgendaVue.AfficherAgenda Date
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_OuvrirJournal()
    modJournal.OuvrirJournal
End Sub

Public Sub UI_CalageCerfa()
    modCerfaPrint.CalageCerfa
End Sub

' Selection d'un patient (version Excel)
Public Function ChoisirPatientX() As Object
    Dim f As ufListe
    Set f = New ufListe
    f.Configurer "Choisir un patient (tapez pour filtrer)", _
                 modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS"), _
                 Array("Nom", "Prenom", "DDN", "Ville", "ID"), "120 pt;90 pt;70 pt;80 pt;40 pt"
    f.Show vbModal
    If Not f.Annule Then Set ChoisirPatientX = f.Resultat
    Unload f
End Function
