Attribute VB_Name = "modRuban"
Option Explicit
' =====================================================================
' modRuban - Rappels du ruban "Cabinet" (customUI14.xml injecte dans
' Cabinet.dotm par build.ps1). Un seul point d'entree : Ruban_Action.
' Les macros sont appelees par leur nom (Application.Run) : ce sont les
' memes que celles des raccourcis clavier et des commandes Dragon.
' =====================================================================

Public Sub Ruban_Action(ByVal control As IRibbonControl)
    On Error GoTo Erreur
    Dim macro As String
    Select Case control.id
        Case "cabNouveau":      macro = "Unifie_A_NouvelleLettre"
        Case "cabPatient":      macro = "Unifie_C_InsererPatient"
        Case "cabCorriger":     macro = "Unifie_D_Finaliser"
        Case "cabDerivee":      macro = "Unifie_D_Finaliser"
        Case "cabValider":      macro = "Unifie_D_Finaliser"
        Case "cabReprendre":     macro = "Unifie_ReprendreBrouillon"
        Case "cabAnciennes":    macro = "Unifie_AfficherArriveesAnciennes"
        Case "cabEcg":          macro = "EnvoyerECG"
        Case "cabGras":         macro = "MettreEnGras"
        Case "cabMedicaments":  macro = "OuvrirDictionnaireMedicaments"
        Case "cabExpressions":  macro = "OuvrirDictionnaireExpressions"
        Case "cabAide":         macro = "AideCabinet"
        Case "cabDiagnostic":   macro = "DiagnosticCabinet"
        Case Else
            MsgBox "Bouton inconnu : " & control.id, vbExclamation, "Cabinet"
            Exit Sub
    End Select
    Application.Run macro
    Exit Sub
Erreur:
    modLog.LogErreur "Ruban " & control.id & " : " & Err.Description
    MsgBox Err.Description, vbCritical, "Cabinet"
End Sub
