Attribute VB_Name = "clsNSBouton"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public WithEvents ctl As MSForms.CommandButton
Attribute ctl.VB_VarHelpID = -1

Public role As String

Private Sub Ctl_Click()

    NS_EvenementBouton role

End Sub
