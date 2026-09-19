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

Public Function MontantReglement(ByVal texte As String) As Currency
    Dim re As Object, morceaux As Variant, fraction As String
    texte = Trim$(texte)
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[0-9]{1,7}([.,][0-9]{1,2})?$"
    If Not re.Test(texte) Then Err.Raise vbObjectError + 661, , "Montant invalide : chiffres et au plus deux decimales attendus."
    morceaux = Split(Replace(texte, ",", "."), ".")
    If UBound(morceaux) = 1 Then fraction = morceaux(1)
    fraction = Left$(fraction & "00", 2)
    MontantReglement = CCur(morceaux(0)) + CCur(fraction) / 100@
End Function

Public Function DateEncaissement(ByVal texte As String) As String
    If Not modTexte.DateFrValide(texte) Then Err.Raise vbObjectError + 661, , "Date invalide."
    ' Le separateur litteral ne depend pas des parametres regionaux Windows.
    DateEncaissement = Format$(modTexte.DateFr(texte), "dd""/""mm""/""yyyy")
End Function

Public Sub UI_EncaisserSeance()
    On Error GoTo Echec
    Dim id As String, mode As String, p As Object, r As Object, saved As Object, line As Object
    Dim patient As Currency, organisme As Currency, total As Currency, recu As Currency, saisie As String
    Dim detail As String, jour As String, payeur As String, choix As String
    Dim lignes As Collection, items As Collection, vus As Object, it As Object, f As ufListe
    jour = Trim$(InputBox("Date de la seance (JJ/MM/AAAA) :", "Rechercher un encaissement", Format$(Date, "dd""/""mm""/""yyyy")))
    If Len(jour) = 0 Then Exit Sub
    jour = DateEncaissement(jour)
    Set p = modServiceNas.Parametres(): p("date") = jour
    Set lignes = modServiceNas.LirePages("journal.read", p)
    Set items = New Collection: Set vus = CreateObject("Scripting.Dictionary")
    For Each line In lignes
        If CStr(line("Paye")) <> "O" Then
            id = CStr(line("SeanceID"))
            If Not vus.Exists(id) Then
                Set it = modServiceNas.Parametres(): it("ID") = id
                it("Patient") = CStr(line("Nom")) & " " & CStr(line("Prenom"))
                it("Naissance") = CStr(line("DDN")): it("Date") = CStr(line("Date"))
                items.Add it: vus(id) = True
            End If
        End If
    Next line
    If items.Count = 0 Then
        MsgBox "Aucune seance impayee pour cette date.", vbInformation, "Encaissement"
        Exit Sub
    End If
    Set f = New ufListe
    f.Configurer "Selectionner la seance", items, Array("Patient", "Naissance", "Date", "ID"), "180 pt;85 pt;85 pt;180 pt"
    f.Show vbModal
    id = "": If Not f.Annule Then id = CStr(f.Resultat("ID"))
    Unload f
    If Len(id) = 0 Then Exit Sub
    Set saved = modServiceNas.CommandeID("billing.get", id)
    For Each line In saved("lignes")
        detail = CStr(line("Nom")) & " " & CStr(line("Prenom")) & " - " & CStr(line("Date"))
        If CStr(line("Paye")) <> "O" Then
            If CStr(line("TiersPayant")) = "O" Then
                organisme = organisme + MontantReglement(CStr(line("Montant")))
            Else
                patient = patient + MontantReglement(CStr(line("Montant")))
            End If
        End If
    Next line
    choix = Trim$(InputBox(detail & vbCrLf & "1 : Patient, solde " & Format$(patient, "0.00") & " EUR" & vbCrLf & "2 : Organisme, solde " & Format$(organisme, "0.00") & " EUR" & vbCrLf & "Qui a effectue ce reglement ?", "Payeur"))
    If Len(choix) = 0 Then Exit Sub
    Select Case choix
        Case "1": payeur = "Patient": total = patient
        Case "2": payeur = "Organisme": total = organisme
        Case Else: Err.Raise vbObjectError + 661, , "Choisissez 1 ou 2."
    End Select
    If total <= 0 Then Err.Raise vbObjectError + 661, , "Aucun solde impaye pour ce payeur."
    saisie = Trim$(InputBox("Montant reellement recu de : " & payeur & vbCrLf & "Solde attendu : " & Format$(total, "0.00") & " EUR. Les reglements partiels ne sont pas pris en charge.", "Montant recu"))
    If Len(saisie) = 0 Then Exit Sub
    recu = MontantReglement(saisie)
    If recu <> total Then Err.Raise vbObjectError + 661, , "Le montant recu differe du solde. Aucun reglement enregistre."
    mode = Trim$(InputBox("Mode du reglement recu : CB, Cheque, Especes ou Virement", "Encaissement"))
    If Len(mode) = 0 Then Exit Sub
    If MsgBox(detail & vbCrLf & payeur & " : " & Format$(recu, "0.00") & " EUR par " & mode & vbCrLf & "Confirmer l encaissement a la date du jour ?", vbYesNo + vbQuestion, "Confirmer le reglement") <> vbYes Then Exit Sub
    Set p = modServiceNas.Parametres(): p("id") = id: p("mode") = mode: p("date") = Format$(Date, "dd""/""mm""/""yyyy")
    p("empreinte") = CStr(saved("empreinte")): p("payeur") = payeur: p("montant") = Replace(Format$(recu, "0.00"), ",", ".")
    Set r = modServiceNas.Appeler("payment", p)
    MsgBox "Reglement enregistre.", vbInformation, "Cabinet"
    Exit Sub
Echec:
    MsgBox Err.Description, vbExclamation, "Encaissement"
End Sub
