Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
' Cabinet.xlsm - poste secretaire

Private Sub Workbook_Open()
    On Error Resume Next
    modEchange.DemarrerScrutation
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modEchange.ArreterScrutation
End Sub

' Double-clic dans la grille de l'agenda : prise de RDV / actions sur un RDV
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Sh.Name = "Agenda" Then
        Cancel = True
        modAgendaVue.ClicCreneau Target
    End If
End Sub
