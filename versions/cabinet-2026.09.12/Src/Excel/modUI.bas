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
    modEchange.SignalerArrivee r
    MsgBox r("Prenom") & " " & r("Nom") & " : arrivee transmise au medecin.", vbInformation, "Cabinet"
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
    Dim recherche As String, f As ufListe
    recherche = Trim$(InputBox("Nom, prenom, date de naissance ou identifiant du patient :", "Rechercher un patient"))
    If Len(recherche) < 2 Then Exit Function
    Set f = New ufListe
    f.Configurer "Patients", modServiceNas.LireTable("PATIENTS", recherche), Array("Nom", "Prenom", "DDN"), "130 pt;130 pt;70 pt"
    f.Show vbModal
    If Not f.Annule Then Set ChoisirPatientX = f.Resultat
    Unload f
End Function

Public Sub UI_AssurePatient()
    On Error GoTo Echec
    Dim pat As Object, cle As Variant, valeur As Variant, p As Object, r As Object
    Set pat = ChoisirPatientX()
    If pat Is Nothing Then Exit Sub
    For Each cle In Array("AssureNom", "AssurePrenom", "AssureDDN", "AssureNIR")
        valeur = Application.InputBox("Assure distinct du patient : " & CStr(cle) & vbCrLf & "Vider les quatre champs pour revenir au patient assure.", "Assure", CStr(pat(CStr(cle))), Type:=2)
        If VarType(valeur) = vbBoolean Then Exit Sub
        pat(CStr(cle)) = Trim$(CStr(valeur))
    Next cle
    Set p = modServiceNas.Parametres(): p("genre") = "PATIENTS": Set p("data") = pat
    Set r = modServiceNas.Appeler("table.update", p)
    MsgBox "Assure enregistre.", vbInformation, "Cabinet"
    Exit Sub
Echec:
    MsgBox Err.Description, vbExclamation, "Assure"
End Sub

Public Sub UI_ParametresDestinataire()
    On Error GoTo Echec
    Dim f As ufListe, cor As Object, p As Object, r As Object, cle As Variant, valeur As Variant
    Set f = New ufListe
    f.Configurer "Parametres des courriers annexes", modServiceNas.LireTable("CORRESPONDANTS"), Array("Nom", "Prenom", "Ville"), "140 pt;100 pt;120 pt"
    f.Show vbModal
    If Not f.Annule Then Set cor = f.Resultat
    Unload f
    If cor Is Nothing Then Exit Sub
    For Each cle In Array("TypesExamen", "ParDefaut", "AValider", "ClesDestination")
        valeur = Application.InputBox(CStr(cle) & vbCrLf & "Types/alias separes par ;. ParDefaut/AValider : 0 ou 1.", "Destinataire", CStr(cor(CStr(cle))), Type:=2)
        If VarType(valeur) = vbBoolean Then Exit Sub
        cor(CStr(cle)) = Trim$(CStr(valeur))
    Next cle
    Set p = modServiceNas.Parametres(): Set p("data") = cor
    Set r = modServiceNas.Appeler("correspondent.save", p)
    MsgBox "Parametres enregistres.", vbInformation, "Cabinet"
    Exit Sub
Echec:
    MsgBox Err.Description, vbExclamation, "Destinataire"
End Sub

Public Sub UI_EncaisserSeance()
    On Error GoTo Echec
    Dim id As String, mode As String, p As Object, r As Object
    id = Trim$(InputBox("Identifiant SeanceID du journal :", "Encaissement"))
    If Len(id) = 0 Then Exit Sub
    mode = Trim$(InputBox("Mode du reglement integral recu :", "Encaissement"))
    If Len(mode) = 0 Then Exit Sub
    Set p = modServiceNas.Parametres(): p("id") = id: p("mode") = mode: p("date") = Format$(Date, "dd/mm/yyyy")
    Set r = modServiceNas.Appeler("payment", p)
    MsgBox "Reglement enregistre.", vbInformation, "Cabinet"
    Exit Sub
Echec:
    MsgBox Err.Description, vbExclamation, "Encaissement"
End Sub
