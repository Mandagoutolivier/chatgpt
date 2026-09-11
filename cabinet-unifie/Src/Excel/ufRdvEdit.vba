Attribute VB_Name = "ufRdvEdit"
Attribute VB_Base = "0{4B212A4B-57F8-489B-8E6C-32AAED210234}{C5CEAA6D-ACF7-4A87-A929-8125013208C5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' Prise de rendez-vous (poste secretaire).
Private mPatient As Object

Private Sub UserForm_Initialize()
    Dim a As Object, duree As Variant
    txtDate.Text = Format$(Date, "dd/mm/yyyy")
    txtHeure.Text = ""
    ' valeurs par defaut du cabinet : config.ini [AGENDA] DureeDefautMin / ActeDefaut
    Dim dureeDef As String, acteDef As String, i As Long
    dureeDef = CStr(modConfig.ConfigNum("AGENDA", "DureeDefautMin", 15))
    acteDef = modConfig.Config("AGENDA", "ActeDefaut", "APC")
    For Each duree In Array("15", "20", "30", "45", "60")
        cmbDuree.AddItem CStr(duree)
    Next duree
    cmbDuree.ListIndex = 0
    For i = 0 To cmbDuree.ListCount - 1
        If cmbDuree.List(i) = dureeDef Then cmbDuree.ListIndex = i
    Next i
    cmbType.Clear
    On Error Resume Next
    For Each a In modActes.Nomenclature()
        cmbType.AddItem a("Code")
    Next a
    On Error GoTo 0
    If cmbType.ListCount > 0 Then cmbType.ListIndex = 0
    For i = 0 To cmbType.ListCount - 1
        If UCase$(cmbType.List(i)) = UCase$(acteDef) Then cmbType.ListIndex = i
    Next i
End Sub

' Pre-remplissage depuis l'agenda (double-clic sur un creneau libre)
Public Sub Prefixer(ByVal dateTxt As String, ByVal heureTxt As String, Optional ByVal pat As Object = Nothing)
    txtDate.Text = dateTxt
    txtHeure.Text = heureTxt
    If Not pat Is Nothing Then
        Set mPatient = pat
        lblPatient.Caption = pat("Prenom") & " " & pat("Nom") & " (" & pat("DDN") & ")"
    End If
End Sub

Private Sub btnAgenda_Click()
    Dim d As Date
    On Error Resume Next
    d = CDate(Trim$(txtDate.Text))
    If Err.Number <> 0 Then d = Date
    On Error GoTo 0
    Me.Hide
    modAgendaVue.AfficherAgenda d, mPatient
End Sub

Private Sub btnPatient_Click()
    Dim p As Object
    Set p = modUI.ChoisirPatientX()
    If Not p Is Nothing Then
        Set mPatient = p
        lblPatient.Caption = p("Prenom") & " " & p("Nom") & " (" & p("DDN") & ")"
    End If
End Sub

Private Sub btnOK_Click()
    On Error GoTo Erreur
    Dim id As String
    If mPatient Is Nothing Then
        MsgBox "Choisissez d'abord le patient.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    If Not modTexte.DateFrValide(txtDate.Text) Then
        MsgBox "Date invalide (jj/mm/aaaa).", vbExclamation, "Cabinet"
        Exit Sub
    End If
    If Not modTexte.HeureValide(txtHeure.Text) Then
        MsgBox "Heure invalide (hh:mm).", vbExclamation, "Cabinet"
        Exit Sub
    End If
    id = modAgenda.AjouterRdv(mPatient("ID"), Trim$(txtDate.Text), Trim$(txtHeure.Text), _
                              cmbDuree.Text, cmbType.Text, Trim$(txtNotes.Text))
    MsgBox "RDV enregistre (" & id & ") : " & mPatient("Prenom") & " " & mPatient("Nom") & _
           " le " & txtDate.Text & " a " & txtHeure.Text & ".", vbInformation, "Cabinet"
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Sub btnAnnuler_Click()
    Me.Hide
End Sub
