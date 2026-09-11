Attribute VB_Name = "ufProgression"
Attribute VB_Base = "0{63DB833A-935D-4162-AC68-B23C994D99B7}{298B30B7-DE85-406B-AB4C-057716036D37}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' Petite fenetre non bloquante affichee pendant les appels a l'API.

Public Sub Definir(ByVal msg As String)
    lblMsg.Caption = msg
    Me.Repaint
    DoEvents
End Sub
