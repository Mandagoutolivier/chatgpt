Attribute VB_Name = "frmCTDestination"
Attribute VB_Base = "0{1D65294C-7845-4699-A689-11CA5258A836}{5CD8A24E-D11D-4BF5-AB06-4C1823D75490}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private mDestinations As Collection

Private Sub UserForm_Initialize()

    gCleDestinationChoisieCabinetTest = ""

    Set mDestinations = _
        CT_ListeDestinationsPourSelection()

    CT_RemplirListe ""

End Sub

Private Sub txtRecherche_Change()

    CT_RemplirListe txtRecherche.Text

End Sub

Private Sub lstCorrespondants_DblClick( _
    ByVal Cancel As MSForms.ReturnBoolean)

    CT_ValiderChoix

End Sub

Private Sub cmdUtiliser_Click()

    CT_ValiderChoix

End Sub

Private Sub cmdAnnuler_Click()

    gCleDestinationChoisieCabinetTest = ""
    Unload Me

End Sub

Private Sub CT_ValiderChoix()

    If lstCorrespondants.ListIndex < 0 Then

        MsgBox _
            "Sélectionnez d'abord un destinataire dans la liste.", _
            vbExclamation, _
            "Cabinet Test"

        Exit Sub

    End If

    gCleDestinationChoisieCabinetTest = _
        CStr(lstCorrespondants.List( _
            lstCorrespondants.ListIndex, 4))

    Unload Me

End Sub

Private Sub CT_RemplirListe( _
    ByVal filtre As String)

    Dim destination As Object
    Dim texteRecherche As String
    Dim numeroLigne As Long
    Dim mots() As String
    Dim mot As Variant
    Dim correspond As Boolean

    lstCorrespondants.Clear

    filtre = LCase$(Trim$(filtre))

    If mDestinations Is Nothing Then Exit Sub

    For Each destination In mDestinations

        texteRecherche = _
            LCase$( _
                CStr(destination("Nom")) & " " & _
                CStr(destination("Type")) & " " & _
                CStr(destination("Structure")) & " " & _
                CStr(destination("Ville")) & " " & _
                CStr(destination("Cle")))

        correspond = True

        If filtre <> "" Then

            mots = Split(filtre, " ")

            For Each mot In mots

                If Trim$(CStr(mot)) <> "" Then

                    If InStr( _
                        1, _
                        texteRecherche, _
                        Trim$(CStr(mot)), _
                        vbTextCompare) = 0 Then

                        correspond = False
                        Exit For

                    End If

                End If

            Next mot

        End If

        If correspond Then

            lstCorrespondants.AddItem _
                CStr(destination("Nom"))

            numeroLigne = _
                lstCorrespondants.ListCount - 1

            lstCorrespondants.List( _
                numeroLigne, 1) = _
                CStr(destination("Type"))

            lstCorrespondants.List( _
                numeroLigne, 2) = _
                CStr(destination("Structure"))

            lstCorrespondants.List( _
                numeroLigne, 3) = _
                CStr(destination("Ville"))

            lstCorrespondants.List( _
                numeroLigne, 4) = _
                CStr(destination("Cle"))

        End If

    Next destination

End Sub
