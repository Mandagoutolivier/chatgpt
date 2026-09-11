Attribute VB_Name = "clsNCCombo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public WithEvents ctl As MSForms.ComboBox
Attribute ctl.VB_VarHelpID = -1

Public role As String

Private Sub Ctl_Change()

    NC_EvenementCombo role

End Sub
