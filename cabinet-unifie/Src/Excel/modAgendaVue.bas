Attribute VB_Name = "modAgendaVue"
Option Explicit
' =====================================================================
' modAgendaVue - Vue hebdomadaire de l'agenda (feuille "Agenda").
' Grille : une colonne par jour, une ligne par creneau (config [AGENDA]).
' Double-clic (evenement de la feuille) : creneau libre -> prise de RDV
' pre-remplie ; creneau occupe -> actions (arrive, absent, annule, fiche).
' Les donnees restent dans Agenda_AAAA.xlsx (modAgenda) ; la grille est
' un simple rendu, rafraichi apres chaque action.
' =====================================================================

Private mLundi As Date                 ' lundi de la semaine affichee
Private mPatientPreselect As Object    ' patient pour prise de RDV directe (dict) ou Nothing
Private mCarte As Object               ' "ligne|col" -> Collection de RDV (dict) du creneau
Private mPatients As Object            ' ID -> dict patient (cache du rendu)

Private Const LIG_TITRE As Long = 2
Private Const LIG_ENTETE As Long = 3
Private Const LIG_DEBUT As Long = 4
Private Const COL_HEURE As Long = 1
Private Const COL_PREMIER_JOUR As Long = 2

' --- parametres -------------------------------------------------------
Private Function NbJours() As Long
    NbJours = CLng(modConfig.ConfigNum("AGENDA", "Jours", 5))
    If NbJours < 1 Then NbJours = 1
    If NbJours > 7 Then NbJours = 7
End Function

Private Function PasMinutes() As Long
    PasMinutes = CLng(modConfig.ConfigNum("AGENDA", "PasMinutes", 15))
    If PasMinutes < 5 Then PasMinutes = 5
End Function

' Heure (texte hh:mm) correspondant a une ligne de la grille
Private Function HeureDeLigne(ByVal ligne As Long) As String
    Dim minutes As Long
    minutes = MinutesDebut() + (ligne - LIG_DEBUT) * PasMinutes()
    If ligne < LIG_DEBUT Or minutes >= MinutesFin() Then HeureDeLigne = "" Else HeureDeLigne = TexteHeure(minutes)
End Function

Private Function MinutesDebut() As Long
    MinutesDebut = MinutesDe(modConfig.Config("AGENDA", "HeureDebut", "08:00"))
End Function

Private Function MinutesFin() As Long
    MinutesFin = MinutesDe(modConfig.Config("AGENDA", "HeureFin", "19:00"))
End Function

Private Function MinutesDe(ByVal hhmm As String) As Long
    MinutesDe = modTexte.MinutesDepuisMinuit(hhmm)
    If MinutesDe < 0 Then Err.Raise vbObjectError + 660, "modAgendaVue", "Heure invalide dans l agenda."
End Function
Private Function TexteHeure(ByVal minutes As Long) As String
    TexteHeure = Format$(minutes \ 60, "00") & ":" & Format$(minutes Mod 60, "00")
End Function

' --- affichage --------------------------------------------------------
Public Sub AfficherAgenda(Optional ByVal dateRef As Date = 0, Optional ByVal pat As Object = Nothing)
    If dateRef = 0 Then dateRef = Date
    Set mPatientPreselect = pat
    mLundi = dateRef - Weekday(dateRef, vbMonday) + 1
    Rendre
End Sub

Public Sub AgendaSemainePrecedente()
    If mLundi = 0 Then mLundi = Date - Weekday(Date, vbMonday) + 1
    mLundi = mLundi - 7
    Rendre
End Sub

Public Sub AgendaSemaineSuivante()
    If mLundi = 0 Then mLundi = Date - Weekday(Date, vbMonday) + 1
    mLundi = mLundi + 7
    Rendre
End Sub

Public Sub AgendaAujourdhui()
    mLundi = Date - Weekday(Date, vbMonday) + 1
    Rendre
End Sub

Public Sub AgendaRetourAccueil()
    Set mPatientPreselect = Nothing
    ThisWorkbook.Worksheets("Accueil").Activate
End Sub

Private Function Feuille() As Worksheet
    On Error Resume Next
    Set Feuille = ThisWorkbook.Worksheets("Agenda")
    If Feuille Is Nothing Then
        Set Feuille = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Feuille.Name = "Agenda"
    End If
End Function

Public Sub Rendre()
    On Error GoTo Erreur
    Dim ws As Worksheet, j As Long, r As Long, minutes As Long, nbLignes As Long
    Dim rdvs As Collection, rdv As Object, p As Object
    Dim jours As Variant
    jours = Array("Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim")
    If mLundi = 0 Then mLundi = Date - Weekday(Date, vbMonday) + 1
    Set ws = Feuille()
    If MinutesFin() <= MinutesDebut() Or PasMinutes() > MinutesFin() - MinutesDebut() Then Err.Raise vbObjectError + 661, "modAgendaVue", "Plage horaire de l agenda invalide."
    Application.ScreenUpdating = False

    ' nettoyage de la grille
    ws.Range(ws.Cells(LIG_TITRE, 1), ws.Cells(Application.Max(400, ws.UsedRange.Rows.Count), COL_PREMIER_JOUR + 7)).Clear
    ws.Cells(LIG_TITRE, 1).Value = "Semaine du " & Format$(mLundi, "dd/mm/yyyy") & " au " & _
                                   Format$(mLundi + NbJours() - 1, "dd/mm/yyyy")
    ws.Cells(LIG_TITRE, 1).Font.Bold = True
    ws.Cells(LIG_TITRE, 1).Font.Size = 12
    If Not mPatientPreselect Is Nothing Then
        ws.Cells(LIG_TITRE, 4).Value = "RDV pour " & mPatientPreselect("Prenom") & " " & mPatientPreselect("Nom") & _
                                       " : double-cliquez un creneau libre"
        ws.Cells(LIG_TITRE, 4).Font.Color = RGB(192, 0, 0)
        ws.Cells(LIG_TITRE, 4).Font.Bold = True
    Else
        ws.Cells(LIG_TITRE, 4).Value = "Double-clic : creneau libre = prendre RDV | creneau occupe = arrive / absent / annuler / fiche"
        ws.Cells(LIG_TITRE, 4).Font.Color = RGB(90, 90, 90)
    End If

    ' en-tetes
    ws.Cells(LIG_ENTETE, COL_HEURE).Value = "Heure"
    For j = 0 To NbJours() - 1
        ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR + j).Value = jours(j) & " " & Format$(mLundi + j, "dd/mm")
        If mLundi + j = Date Then ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR + j).Interior.Color = RGB(255, 235, 156)
    Next j
    With ws.Range(ws.Cells(LIG_ENTETE, 1), ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR + NbJours() - 1))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' creneaux (colonne des heures en TEXTE : sinon Excel convertit "08:00" en fraction de jour)
    ws.Columns(COL_HEURE).NumberFormat = "@"
    nbLignes = 0
    For minutes = MinutesDebut() To MinutesFin() - PasMinutes() Step PasMinutes()
        r = LIG_DEBUT + nbLignes
        ws.Cells(r, COL_HEURE).Value = TexteHeure(minutes)
        If (minutes Mod 60) <> 0 Then ws.Cells(r, COL_HEURE).Font.Color = RGB(120, 120, 120)
        ws.Cells(r, COL_HEURE).HorizontalAlignment = xlCenter
        ws.Cells(r, COL_HEURE).Font.Bold = True
        With ws.Range(ws.Cells(r, COL_PREMIER_JOUR), ws.Cells(r, COL_PREMIER_JOUR + NbJours() - 1))
            .Interior.Color = RGB(235, 250, 235)      ' libre
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        nbLignes = nbLignes + 1
    Next minutes
    ws.Range(ws.Cells(LIG_DEBUT, 1), ws.Cells(LIG_DEBUT + nbLignes - 1, 1)).Borders.LineStyle = xlContinuous

    ' rendez-vous de la semaine
    Set mCarte = CreateObject("Scripting.Dictionary")
    ChargerPatients
    For Each rdv In RdvSemaine()
        PlacerRdv ws, rdv
    Next rdv

    ' mise en page
    ws.Columns(COL_HEURE).ColumnWidth = 7
    For j = 0 To NbJours() - 1
        ws.Columns(COL_PREMIER_JOUR + j).ColumnWidth = 26
    Next j
    ws.Rows(LIG_DEBUT & ":" & (LIG_DEBUT + nbLignes - 1)).RowHeight = IIf(PasMinutes() < 30, 16, 30)
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.FreezePanes = False
    ws.Cells(LIG_DEBUT, COL_PREMIER_JOUR).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(LIG_DEBUT, COL_PREMIER_JOUR).Select
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    MsgBox "Affichage de l'agenda impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub

Private Sub ChargerPatients()
    Dim p As Object
    Set mPatients = CreateObject("Scripting.Dictionary")
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        Set mPatients(p("ID")) = p
    Next p
End Sub

' RDV de la semaine affichee (les deux annees si la semaine chevauche)
Private Function RdvSemaine() As Collection
    Dim col As Collection, r As Object, d As Date, annees As Object, a As Variant, tous As Collection
    Set col = New Collection
    Set annees = CreateObject("Scripting.Dictionary")
    annees(Year(mLundi)) = 1
    annees(Year(mLundi + NbJours() - 1)) = 1
    For Each a In annees.Keys
        modAgenda.AssurerAgendaAnnee CLng(a)
        Set tous = modBaseIO.LireTableX(modConfig.FichierAgenda(CLng(a)), "RDV")
        If Not tous Is Nothing Then
            For Each r In tous
                d = DateDe(r("Date"))
                If d >= mLundi And d <= mLundi + NbJours() - 1 Then col.Add r
            Next r
        End If
    Next a
    Set RdvSemaine = col
End Function

Private Function DateDe(ByVal texte As String) As Date
    DateDe = modTexte.DateFr(texte)
End Function
Private Sub PlacerRdv(ByVal ws As Worksheet, ByVal rdv As Object)
    Dim d As Date, minutes As Long, ligne As Long, col As Long, cle As String
    Dim nom As String, texte As String, nbCreneaux As Long, k As Long, statut As String
    ' un RDV annule libere le creneau : il n'est plus affiche (il reste dans
    ' la base avec le statut Annule pour l'historique)
    If rdv("Statut") = "Annule" Then Exit Sub
    d = DateDe(rdv("Date"))
    minutes = MinutesDe(rdv("Heure"))
    If minutes < MinutesDebut() Or minutes >= MinutesFin() Then Exit Sub
    col = COL_PREMIER_JOUR + (d - mLundi)
    ligne = LIG_DEBUT + (minutes - MinutesDebut()) \ PasMinutes()
    cle = ligne & "|" & col
    If mPatients.Exists(rdv("PatientID")) Then
        nom = mPatients(rdv("PatientID"))("Nom") & " " & mPatients(rdv("PatientID"))("Prenom")
    Else
        nom = rdv("PatientID")
    End If
    statut = rdv("Statut")
    texte = nom
    If Len(rdv("TypeActe")) > 0 Then texte = texte & " - " & rdv("TypeActe")
    If statut = "Arrive" Then texte = texte & " (arrive " & rdv("HeureArrivee") & ")"
    If statut = "Absent" Then texte = texte & " (ABSENT)"
    If statut = "Honore" Then texte = texte & " (vu)"
    If Len(ws.Cells(ligne, col).Value) > 0 Then
        ws.Cells(ligne, col).Value = ws.Cells(ligne, col).Value & vbLf & texte
    Else
        ws.Cells(ligne, col).Value = texte
    End If
    ws.Cells(ligne, col).Interior.Color = CouleurStatut(statut)
    If Not mCarte.Exists(cle) Then Set mCarte(cle) = New Collection
    mCarte(cle).Add rdv
    ' creneaux suivants couverts par la duree
    nbCreneaux = (Val(rdv("DureeMin")) + PasMinutes() - 1) \ PasMinutes()
    For k = 1 To nbCreneaux - 1
        If Len(ws.Cells(ligne + k, col).Value) = 0 And ligne + k < LIG_DEBUT + (MinutesFin() - MinutesDebut()) \ PasMinutes() Then
            ws.Cells(ligne + k, col).Value = ChrW$(8627) & " " & nom
            ws.Cells(ligne + k, col).Interior.Color = CouleurStatut(statut)
            ws.Cells(ligne + k, col).Font.Color = RGB(120, 120, 120)
            Dim cle2 As String
            cle2 = (ligne + k) & "|" & col
            If Not mCarte.Exists(cle2) Then Set mCarte(cle2) = New Collection
            mCarte(cle2).Add rdv
        End If
    Next k
End Sub

Private Function CouleurStatut(ByVal statut As String) As Long
    Select Case statut
        Case "Arrive": CouleurStatut = RGB(197, 224, 255)   ' bleu clair
        Case "Honore": CouleurStatut = RGB(217, 217, 217)   ' gris
        Case "Absent": CouleurStatut = RGB(255, 199, 199)   ' rouge clair
        Case "Annule": CouleurStatut = RGB(230, 230, 230)
        Case Else:     CouleurStatut = RGB(255, 242, 180)   ' prevu : jaune clair
    End Select
End Function

' --- interaction (appelee par Worksheet_BeforeDoubleClick de la feuille) ---
Public Sub ClicCreneau(ByVal cible As Range)
    On Error GoTo Erreur
    Dim ligne As Long, col As Long, cle As String, d As Date, heure As String
    ligne = cible.Row: col = cible.Column
    If mLundi = 0 Or mCarte Is Nothing Then Rendre
    If ligne < LIG_DEBUT Or col < COL_PREMIER_JOUR Or col > COL_PREMIER_JOUR + NbJours() - 1 Then Exit Sub
    heure = HeureDeLigne(ligne)
    If Len(heure) = 0 Then Exit Sub
    d = mLundi + (col - COL_PREMIER_JOUR)
    cle = ligne & "|" & col
    If mCarte.Exists(cle) Then
        ActionsCreneau mCarte(cle)
    Else
        PrendreRdv d, heure
    End If
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Function CreneauLibre(ByVal d As Date, ByVal heure As String) As Boolean
    Dim r As Object
    CreneauLibre = True
    For Each r In modAgenda.RdvDuJour(Format$(d, "dd/mm/yyyy"))
        If r("Statut") <> "Annule" Then
            If modTexte.IntervallesSeChevauchent(MinutesDe(heure), 1, MinutesDe(r("Heure")), CLng(Val(r("DureeMin")))) Then
                CreneauLibre = False
                Exit Function
            End If
        End If
    Next r
End Function
Private Sub PrendreRdv(ByVal d As Date, ByVal heure As String)
    Dim f As ufRdvEdit
    Set f = New ufRdvEdit
    f.Prefixer Format$(d, "dd/mm/yyyy"), heure, mPatientPreselect
    f.Show vbModal
    Unload f
    Set mPatientPreselect = Nothing
    Rendre
End Sub

Private Sub ActionsCreneau(ByVal rdvs As Collection)
    Dim items As Collection, r As Object, it As Object, f As ufListe, nom As String
    Dim actions As Variant, a As Variant
    actions = Array(Array("ARRIVE", "Marquer ARRIVE"), Array("ABSENT", "Marquer ABSENT"), _
                    Array("ANNULE", "Annuler le rendez-vous"), Array("FICHE", "Ouvrir la fiche patient"))
    Set items = New Collection
    For Each r In rdvs
        If mPatients.Exists(r("PatientID")) Then
            nom = mPatients(r("PatientID"))("Prenom") & " " & mPatients(r("PatientID"))("Nom")
        Else
            nom = r("PatientID")
        End If
        For Each a In actions
            Set it = CreateObject("Scripting.Dictionary")
            it.CompareMode = 1
            it("ID") = r("ID") & "|" & a(0)
            it("Patient") = nom & " (" & r("Heure") & ")"
            it("Action") = a(1)
            items.Add it
        Next a
    Next r
    Set f = New ufListe
    f.Configurer "Rendez-vous : que faire ?", items, Array("Patient", "Action"), "170 pt;190 pt"
    f.Show vbModal
    If Not f.Annule Then
        Dim choix As String, rdvID As String, action As String
        choix = f.Resultat("ID")
        rdvID = Left$(choix, InStr(choix, "|") - 1)
        action = mID$(choix, InStr(choix, "|") + 1)
        Unload f
        ExecuterAction rdvs, rdvID, action
    Else
        Unload f
    End If
End Sub

Private Sub ExecuterAction(ByVal rdvs As Collection, ByVal rdvID As String, ByVal action As String)
    Dim r As Object, x As Object, annee As Long, fp As ufPatientEdit
    For Each x In rdvs
        If x("ID") = rdvID Then Set r = x
    Next x
    If r Is Nothing Then Exit Sub
    annee = Year(DateDe(r("Date")))
    Select Case action
        Case "ARRIVE"
            modEchange.SignalerArrivee r
        Case "ABSENT"
            modEchange.RetirerArrivee r
            modAgenda.MarquerStatut rdvID, "Absent", annee
        Case "ANNULE"
            If MsgBox("Annuler ce rendez-vous ?", vbYesNo + vbQuestion, "Cabinet") = vbYes Then
                modEchange.RetirerArrivee r
                modAgenda.MarquerStatut rdvID, "Annule", annee
            End If
        Case "FICHE"
            If mPatients.Exists(r("PatientID")) Then
                Set fp = New ufPatientEdit
                fp.ChargerExistant mPatients(r("PatientID"))
                fp.Show vbModal
                Unload fp
            End If
    End Select
    Rendre
End Sub
