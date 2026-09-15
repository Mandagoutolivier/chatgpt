Attribute VB_Name = "modRecetteU1Excel"
Option Explicit

Public Function Executer() As String
    Dim col As Collection, cas As Variant, numero As Long, reussis As Long, description As String
    On Error GoTo Echec
    Set col = modCerfaPrint.AnalyserPositions("# TEST FICTIF" & vbLf & "ASSURE_NOM;20,5;30;100;10")
    If col.Count <> 1 Or CDbl(col(1)(1)) <> 20.5 Then Err.Raise 5, , "Positions CERFA valides"
    reussis = 1
    For Each cas In Array("ASSURE_NOM;abc;30;100;10", "ASSURE_NOM;20;30;-1;10", "ASSURE_NOM;200;30;100;10", "ASSURE_NOM;20;295;100;10", "ASSURE_NOM;20;30;100;0", "ASSURE_NOM;20;30;100", "ASSURE_NOM;20;30;100;10" & vbLf & "assure_nom;20;30;100;10", "# vide")
        numero = 0
        On Error Resume Next
        Set col = modCerfaPrint.AnalyserPositions(CStr(cas))
        numero = Err.Number: Err.Clear
        On Error GoTo Echec
        If numero <> vbObjectError + 710 Then Err.Raise 5, , "Position CERFA invalide acceptee"
        reussis = reussis + 1
    Next cas
    If modServiceNas.SHA256("abc") <> "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad" Then Err.Raise 5, , "SHA256 Excel"
    reussis = reussis + 1
    Executer = "{""reussis"":" & CStr(reussis) & ",""echec"":false}"
    Exit Function
Echec:
    description = Err.Description
    Executer = "{""reussis"":" & CStr(reussis) & ",""echec"":true,""description"":" & modServiceNas.JsonValeur(description) & "}"
End Function
