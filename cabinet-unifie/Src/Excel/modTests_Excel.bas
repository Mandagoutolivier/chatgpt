Attribute VB_Name = "modTests_Excel"
Option Explicit
Option Private Module
' =====================================================================
' modTests_Excel - Suites de tests du poste secretaire (donnees FICTIVES).
' Lancees par run_tests.ps1 -Cible Excel.
' =====================================================================

Public Sub Test_J1(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "J1 base secretaire"
    On Error GoTo Echec
    Dim d As Object, id As String, pats As Collection, p As Object, trouve As Object

    ' --- ajout d'un patient ---
    Set d = CreateObject("Scripting.Dictionary")
    d("Nom") = "TESTJ1"
    d("Prenom") = "Alice"
    d("DDN") = "02/02/1960"
    d("Sexe") = "F"
    d("Ville") = "Montpellier"
    d("DateCreation") = Format$(Now, "dd/mm/yyyy")
    id = modBaseIO.AjouterLigne(modConfig.FichierPatients(), "PATIENTS", d, "P")
    modLog.Verifier "patient ajoute avec ID", Left$(id, 1) = "P" And Len(id) = 6, id

    Set pats = modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
    Set trouve = Nothing
    For Each p In pats
        If p("ID") = id Then Set trouve = p
    Next p
    modLog.Verifier "patient relu", Not trouve Is Nothing
    If Not trouve Is Nothing Then
        modLog.Verifier "champs corrects", trouve("Nom") = "TESTJ1" And trouve("DDN") = "02/02/1960", _
                        trouve("Nom") & " " & trouve("DDN")
    End If

    ' --- modification ---
    Set d = CreateObject("Scripting.Dictionary")
    d("Ville") = "Beziers"
    d("DateModif") = Format$(Now, "dd/mm/yyyy")
    modBaseIO.ModifierLigne modConfig.FichierPatients(), "PATIENTS", "ID", id, d
    Set pats = modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
    For Each p In pats
        If p("ID") = id Then Set trouve = p
    Next p
    modLog.Verifier "modification relue", trouve("Ville") = "Beziers", trouve("Ville")

    ' --- concurrence : verrou pris -> l'ecriture echoue proprement ---
    modFichiers.AcquerirVerrou "Patients", 2000
    Dim erreurAttendue As Boolean
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    d("Nom") = "NEDOITPASEXISTER"
    modBaseIO.AjouterLigne modConfig.FichierPatients(), "PATIENTS", d, "P"
    erreurAttendue = (Err.Number <> 0)
    On Error GoTo Echec
    modFichiers.RelacherVerrou "Patients"
    modLog.Verifier "ecriture refusee sous verrou", erreurAttendue

    ' --- agenda ---
    Dim rdvID As String, rdvs As Collection, r As Object
    rdvID = modAgenda.AjouterRdv(id, Format$(Date, "dd/mm/yyyy"), "11:30", "30", "CSC", "test")
    modLog.Verifier "rdv cree", Left$(rdvID, 1) = "R", rdvID
    modAgenda.MarquerStatut rdvID, "Arrive"
    Set rdvs = modAgenda.RdvDuJour()
    Dim statutLu As String
    For Each r In rdvs
        If r("ID") = rdvID Then statutLu = r("Statut") & "/" & r("HeureArrivee")
    Next r
    modLog.Verifier "rdv arrive avec heure", InStr(statutLu, "Arrive") = 1 And Len(statutLu) > 7, statutLu

    ' --- nomenclature ---
    Dim nomen As Collection
    Set nomen = modActes.Nomenclature()
    modLog.Verifier "nomenclature lue", nomen.Count >= 5, nomen.Count & " actes"
    modLog.Verifier "acte CSC avec associe", Not modActes.ActeParCode("CSC") Is Nothing And _
                    Len(modActes.ActeParCode("CSC")("CodeAssocie")) > 0, _
                    "CSC + " & modActes.ActeParCode("CSC")("CodeAssocie")
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

Public Sub Test_AGENDA(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "AGENDA vue hebdomadaire"
    On Error GoTo Echec
    Dim ws As Worksheet, rdvID As String, auj As String, trouve As Boolean, c As Range, col As Long, j As Long
    auj = Format$(Date, "dd/mm/yyyy")
    rdvID = modAgenda.AjouterRdv("P00002", auj, "10:00", "30", "ETT", "test agenda")
    modLog.Verifier "rdv de test cree", Len(rdvID) > 0, rdvID
    modAgendaVue.AfficherAgenda Date
    Set ws = ThisWorkbook.Worksheets("Agenda")
    modLog.Verifier "feuille Agenda presente", Not ws Is Nothing
    modLog.Verifier "titre semaine", InStr(ws.Cells(2, 1).Value, "Semaine du") > 0, ws.Cells(2, 1).Value
    ' colonne du jour et ligne 10:00
    trouve = False
    For Each c In ws.Range(ws.Cells(4, 2), ws.Cells(60, 8)).Cells
        If InStr(c.Value, "SANTONJA") > 0 And InStr(c.Value, "ETT") > 0 Then trouve = True
    Next c
    modLog.Verifier "rdv SANTONJA 10:00 affiche dans la grille", trouve
    modLog.Verifier "creneau 10:00 occupe", Not modAgendaVue.CreneauLibre(Date, "10:00")
    modLog.Verifier "creneau 18:30 libre", modAgendaVue.CreneauLibre(Date, "18:30")
    modLog.Verifier "colonne des heures en texte (pas de decimal)", _
        ws.Cells(4, 1).Value = "08:00" And ws.Cells(5, 1).Value = "08:15" And VarType(ws.Cells(5, 1).Value) = vbString, _
        ws.Cells(4, 1).Value & " / " & ws.Cells(5, 1).Value
    modLog.Verifier "5 jours (lundi-vendredi)", Len(ws.Cells(3, 7).Value) = 0 And InStr(ws.Cells(3, 6).Value, "Ven") = 1, _
        ws.Cells(3, 6).Value & " | " & ws.Cells(3, 7).Value
    ' navigation
    modAgendaVue.AgendaSemaineSuivante
    modLog.Verifier "semaine suivante", InStr(ws.Cells(2, 1).Value, Format$(Date - Weekday(Date, vbMonday) + 8, "dd/mm/yyyy")) > 0, ws.Cells(2, 1).Value
    modAgendaVue.AgendaAujourdhui
    ' statut via l'agenda
    modAgenda.MarquerStatut rdvID, "Annule"
    modAgendaVue.Rendre
    trouve = False
    For Each c In ws.Range(ws.Cells(4, 2), ws.Cells(60, 8)).Cells
        If InStr(c.Value, "SANTONJA") > 0 And InStr(c.Value, "ETT") > 0 Then trouve = True
    Next c
    modLog.Verifier "rdv annule retire de la grille (creneau libere)", Not trouve
    modLog.Verifier "creneau 10:00 libre apres annulation", modAgendaVue.CreneauLibre(Date, "10:00")
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub

Public Sub Test_J5X(Optional ByVal racine As String = "")
    Err.Raise vbObjectError + 989, "Tests historiques", "Ces tests utilisent les anciennes fixtures. Utilisez les tests Audit et la recette Windows du depot."
    modLog.TestDebut "J5 flux secretaire (drapeau -> journal)"
    On Error GoTo Echec
    Dim d As Object, chemin As String, courriers As Collection, c As Object
    Dim actes As Collection, seanceID As String, lignes As Collection, l As Object, nCsc As Long

    ' drapeau fabrique
    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = "P00003"
    d("Nom") = "COMBALUZIER"
    d("Prenom") = "Pierre"
    d("DDN") = "07/03/1939"
    d("NIR") = "1 39 03 34 000 003 42"
    d("TypeCourrier") = "consultation"
    d("DateValidation") = Format$(Now, "dd/mm/yyyy hh:nn")
    chemin = modFichiers.EcrireDrapeau(modConfig.chemin("Echange") & "\AEnvoyer", _
                                       "_testJ5_" & modFichiers.IdUnique() & "_P00003", d)

    Set courriers = modEchange.CourriersEnAttente()
    Dim present As Boolean
    For Each c In courriers
        If c("_Chemin") = chemin Then present = True
    Next c
    modLog.Verifier "drapeau detecte en attente", present, courriers.Count & " en attente"

    ' enregistrement de la seance (acte CSC -> 2 lignes avec DEQP003)
    Set actes = New Collection
    actes.Add modActes.ActeParCode("CSC")
    Set d = modFichiers.LireDrapeau(chemin)
    seanceID = modActes.EnregistrerSeance(d, actes, "CB", False, True)
    modLog.Verifier "seance enregistree", Len(seanceID) > 0, seanceID

    Set lignes = modBaseIO.LireTableX(modConfig.FichierJournal(), "JOURNAL", "SeanceID")
    nCsc = 0
    For Each l In lignes
        If l("SeanceID") = seanceID Then
            nCsc = nCsc + 1
            If nCsc = 1 Then modLog.Verifier "ligne journal complete", _
                l("Nom") = "COMBALUZIER" And l("CodeActe") = "CSC" And l("ModePaiement") = "CB" And l("Paye") = "O", _
                l("Nom") & "/" & l("CodeActe") & "/" & l("ModePaiement")
        End If
    Next l
    modLog.Verifier "2 lignes (CSC + DEQP003)", nCsc = 2, nCsc & " lignes"

    ' feuille de soins vers PDF (controle visuel J6)
    Dim pdf As String
    pdf = modConfig.chemin("Logs") & "\cerfa_test.pdf"
    If Len(Dir$(pdf)) > 0 Then Kill pdf
    modCerfaPrint.ImprimerFeuille d, modActes.LignesPourImpression(actes), pdf
    modLog.Verifier "feuille de soins exportee en PDF", Len(Dir$(pdf)) > 0, pdf

    ' archivage du drapeau
    modEchange.DeplacerVersTraites chemin
    modLog.Verifier "drapeau archive", Len(Dir$(chemin)) = 0 And _
        Len(Dir$(modConfig.chemin("Echange") & "\Traites\" & mID$(chemin, InStrRev(chemin, "\") + 1))) > 0
    Exit Sub
Echec:
    modLog.TestResultat "exception " & Err.Number, False, Err.Description
End Sub
