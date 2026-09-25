Attribute VB_Name = "ufArriveesMedecin"
Option Explicit
Private mItems As Collection
Private mAffiches As Collection
Private mRemplissage As Boolean
Private mDisponible As Boolean

Private Sub UserForm_Initialize()
    Set mItems = New Collection
    Set mAffiches = New Collection
    lstPatients.ColumnCount = 4
    lstPatients.ColumnWidths = "45 pt;45 pt;170 pt;75 pt"
    lblEtat.Caption = "Chargement..."
End Sub

Public Sub Charger(ByVal items As Collection, ByVal heure As String)
    Dim selectionID As String
    selectionID = IdentifiantSelection()
    Set mItems = items
    mDisponible = True
    Remplir selectionID
    lblEtat.Caption = CStr(mItems.Count) & " patient(s) arrive(s) - actualise a " & heure
End Sub

Public Sub Indisponible(ByVal message As String)
    mDisponible = False
    lstPatients.Enabled = False
    lblEtat.Caption = message
End Sub

Public Function IdentifiantSelection() As String
    If mAffiches Is Nothing Then Exit Function
    If lstPatients.ListIndex < 0 Or lstPatients.ListIndex >= mAffiches.Count Then Exit Function
    IdentifiantSelection = CStr(mAffiches(lstPatients.ListIndex + 1)("ConsultationID"))
End Function

Public Function NombreAffiche() As Long
    NombreAffiche = lstPatients.ListCount
End Function

Private Sub Remplir(ByVal selectionID As String)
    Dim it As Object, filtre As String, libelle As String
    mRemplissage = True
    lstPatients.Clear
    Set mAffiches = New Collection
    filtre = modTexte.Plier(Trim$(txtFiltre.Text))
    For Each it In mItems
        libelle = CStr(it("Nom")) & " " & CStr(it("Prenom"))
        If Len(filtre) = 0 Or InStr(modTexte.Plier(libelle & " " & CStr(it("DDN"))), filtre) > 0 Then
            mAffiches.Add it
            lstPatients.AddItem CStr(it("HeureArrivee"))
            lstPatients.List(lstPatients.ListCount - 1, 1) = CStr(it("HeureRdv"))
            lstPatients.List(lstPatients.ListCount - 1, 2) = libelle
            lstPatients.List(lstPatients.ListCount - 1, 3) = CStr(it("DDN"))
        End If
    Next it
    lstPatients.ListIndex = -1
    If Len(selectionID) > 0 Then
        Dim i As Long
        For i = 1 To mAffiches.Count
            If CStr(mAffiches(i)("ConsultationID")) = selectionID Then lstPatients.ListIndex = i - 1
        Next i
    End If
    lstPatients.Enabled = mDisponible
    mRemplissage = False
End Sub

Private Sub txtFiltre_Change()
    If mRemplissage Or mItems Is Nothing Then Exit Sub
    Remplir IdentifiantSelection()
End Sub

Private Sub lstPatients_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim choix As Object
    If mRemplissage Or Not mDisponible Then Exit Sub
    If lstPatients.ListIndex < 0 Or lstPatients.ListIndex >= mAffiches.Count Then Exit Sub
    Set choix = mAffiches(lstPatients.ListIndex + 1)
    Cancel = True
    modFileArrivees.DemarrerSelection choix
End Sub

Private Sub btnActualiser_Click()
    modFileArrivees.ActualiserFileArrivees
End Sub

Private Sub btnReprendre_Click()
    If modPowerMicUnifie.Unifie_OperationEnCours() Then Exit Sub
    modIntegrationUnifie.Unifie_ReprendreBrouillon
End Sub

Private Sub btnMasquer_Click()
    modFileArrivees.SuspendreFileArrivees
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    modFileArrivees.SuspendreFileArrivees
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub
