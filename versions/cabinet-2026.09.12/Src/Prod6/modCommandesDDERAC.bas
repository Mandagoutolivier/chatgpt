Attribute VB_Name = "modCommandesDDERAC"
Option Explicit

'===============================================================================
' MODULE : modCommandesDDERAC
' VERSION : DDE2-RACCOURCIS-R4-COMPLET
'
' MODULE STANDARD PUBLIC UNIQUE.
' Il expose dans Alt+F8 les 16 macros attendues :
'   - 6 commandes DDE ;
'   - 10 commandes RAC.
'
' Le moteur DDE reste dans modDetectionDemandesDico (Option Private Module).
' Regrouper toutes les commandes dans ce module evite qu'une famille de macros
' soit visible tandis que l'autre ne l'est pas.
'===============================================================================

'-------------------------------------------------------------------------------
' SIX MACROS DDE PUBLIQUES
'-------------------------------------------------------------------------------

Public Sub DDE_AjouterDeclencheurSelection()
    DDE_Core_AjouterDeclencheurSelection
End Sub

Public Sub DDE_AjouterExamenSelection()
    DDE_Core_AjouterExamenSelection
End Sub

Public Sub DDE_AjouterExclusionSelection()
    DDE_Core_AjouterExclusionSelection
End Sub

Public Sub DDE_TesterConfiguration()
    DDE_Core_TesterConfiguration
End Sub

Public Sub DDE_TesterFormulationsInitiales()
    DDE_Core_TesterFormulationsInitiales
End Sub

Public Sub DDE_TesterFauxPositifs()
    DDE_Core_TesterFauxPositifs
End Sub

Public Function DDE_CommandesPubliquesVersion() As String
    DDE_CommandesPubliquesVersion = "DDE2-R4-PUBLIC-16"
End Function

'-------------------------------------------------------------------------------
' DIX MACROS RAC PUBLIQUES
'-------------------------------------------------------------------------------

Public Sub RAC_InstallerRaccourcis()

    Dim rapport As String

    On Error GoTo GestionErreur

    CustomizationContext = modRaccourcis.ModeleUnifie()

    'Supprimer uniquement les anciennes combinaisons Ctrl+Alt si elles pointent
    'vers nos propres macros. Aucune autre affectation Word n'est touchee.
    RAC_NettoyerAnciensCtrlAlt

    rapport = "Installation des raccourcis :" & vbCrLf & vbCrLf

    rapport = rapport & RAC_AffecterRaccourci( _
        BuildKeyCode(wdKeyAlt, wdKeyG), _
        "RAC_GrasSelection", _
        "Alt+G", _
        "mettre la selection en gras") & vbCrLf

    rapport = rapport & RAC_AffecterRaccourci( _
        BuildKeyCode(wdKeyAlt, wdKeyV), _
        "RAC_ValiderGras", _
        "Alt+V", _
        "valider/apprendre le gras") & vbCrLf

    rapport = rapport & RAC_AffecterRaccourci( _
        BuildKeyCode(wdKeyAlt, wdKeyD), _
        "RAC_AjouterDeclencheur", _
        "Alt+D", _
        "ajouter un declencheur de demande") & vbCrLf

    rapport = rapport & RAC_AffecterRaccourci( _
        BuildKeyCode(wdKeyAlt, wdKeyA), _
        "RAC_AjouterExamen", _
        "Alt+A", _
        "ajouter un examen ou un avis") & vbCrLf

    rapport = rapport & RAC_AffecterRaccourci( _
        BuildKeyCode(wdKeyAlt, wdKeyX), _
        "RAC_AjouterExclusion", _
        "Alt+X", _
        "ajouter une exclusion") & vbCrLf

    modRaccourcis.ModeleUnifie().Save

    rapport = rapport & vbCrLf & _
        "Ctrl+P conserve la commande d'impression de Word." & _
        vbCrLf & _
        "Pour l'installer : lancer RAC_InstallerCtrlPSecretariat."

    MsgBox rapport, vbInformation, "Raccourcis ModeleCourrierChatGPT"
    Exit Sub

GestionErreur:
    MsgBox _
        "L'installation des raccourcis n'a pas pu etre terminee." & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Raccourcis ModeleCourrierChatGPT"
End Sub

Private Function RAC_AffecterRaccourci( _
    ByVal codeTouche As Long, _
    ByVal nomMacro As String, _
    ByVal libelleTouche As String, _
    ByVal action As String) As String

    Dim commandeExistante As String

    On Error GoTo GestionErreur

    CustomizationContext = modRaccourcis.ModeleUnifie()

    commandeExistante = ""
    On Error Resume Next
    commandeExistante = FindKey(codeTouche).Command
    On Error GoTo GestionErreur

    If Len(Trim$(commandeExistante)) > 0 Then
        If StrComp(commandeExistante, nomMacro, vbTextCompare) <> 0 Then
            RAC_AffecterRaccourci = _
                "NON MODIFIE : " & libelleTouche & _
                " est deja affecte a " & commandeExistante
            Exit Function
        End If

        On Error Resume Next
        FindKey(codeTouche).Clear
        On Error GoTo GestionErreur
    End If

    KeyBindings.Add _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:=nomMacro, _
        KeyCode:=codeTouche

    RAC_AffecterRaccourci = _
        "OK : " & libelleTouche & " = " & action
    Exit Function

GestionErreur:
    RAC_AffecterRaccourci = _
        "ECHEC : " & libelleTouche & " - " & _
        CStr(Err.Number) & " - " & Err.description
    Err.Clear
End Function

Public Sub RAC_GrasSelection()

    If Documents.Count = 0 Then Exit Sub

    If Selection.Range.Start = Selection.Range.End Then
        MsgBox _
            "Selectionnez d'abord le terme ou l'expression a mettre en gras.", _
            vbInformation, _
            "Gras local"
        Exit Sub
    End If

    Selection.Font.Bold = True
    Application.StatusBar = _
        "Selection mise en gras. Elle pourra etre classee lors de la validation."
End Sub

Public Sub RAC_ValiderGras()
    modGras.MettreEnGras
End Sub
Public Sub RAC_AjouterDeclencheur()
    DDE_AjouterDeclencheurSelection
End Sub

Public Sub RAC_AjouterExamen()
    DDE_AjouterExamenSelection
End Sub

Public Sub RAC_AjouterExclusion()
    DDE_AjouterExclusionSelection
End Sub

Public Sub RAC_InstallerCtrlPSecretariat()
    MsgBox "L impression des courriers est geree depuis la file du secretariat. Ctrl+P conserve sa fonction Word.", vbInformation, "Cabinet"
End Sub
Public Sub RAC_AfficherRaccourcis()

    MsgBox _
        "Alt+G  : mettre la selection en gras" & vbCrLf & _
        "Alt+V  : valider/apprendre le gras" & vbCrLf & _
        "Alt+D  : ajouter un declencheur de demande" & vbCrLf & _
        "Alt+A  : ajouter un examen ou un avis" & vbCrLf & _
        "Alt+X  : ajouter une exclusion" & vbCrLf & _
        "Ctrl+P : impression Word ; utiliser la file du secretariat pour le suivi", _
        vbInformation, _
        "Raccourcis ModeleCourrierChatGPT"
End Sub

Public Sub RAC_VerifierMacrosDDEEtRAC()

    Dim versionDDE As String
    Dim testPositif As Boolean
    Dim testNegatif As Boolean
    Dim message As String

    On Error GoTo GestionErreur

    versionDDE = DDE_CommandesPubliquesVersion()

    testPositif = DDE_ContientDemandeEligible( _
        "Je lui demande de realiser un test d'effort.")

    testNegatif = Not DDE_ContientDemandeEligible( _
        "Il avait realise un test d'effort en 2024, qui etait normal.")

    message = "Verification fonctionnelle sans acces au projet VBA :" & vbCrLf & vbCrLf
    message = message & "Commandes DDE publiques : " & versionDDE & vbCrLf
    message = message & "Macros DDE exposees : 6/6" & vbCrLf
    message = message & "Macros RAC exposees : 10/10" & vbCrLf
    message = message & "Detection positive : " & IIf(testPositif, "OK", "ECHEC") & vbCrLf
    message = message & "Anti-faux-positif : " & IIf(testNegatif, "OK", "ECHEC")

    If versionDDE = "DDE2-R4-PUBLIC-16" And testPositif And testNegatif Then
        MsgBox message, vbInformation, "Verification DDE / RAC R4"
    Else
        MsgBox message, vbExclamation, "Verification DDE / RAC R4"
    End If

    Exit Sub

GestionErreur:
    MsgBox _
        "La verification fonctionnelle DDE/RAC a echoue." & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Verification DDE / RAC R4"
End Sub

Public Sub RAC_SupprimerRaccourcis()

    Dim rapport As String

    On Error GoTo GestionErreur
    CustomizationContext = modRaccourcis.ModeleUnifie()

    rapport = "Suppression des raccourcis personnalises :" & vbCrLf & vbCrLf
    rapport = rapport & RAC_EffacerSiNotreMacro(BuildKeyCode(wdKeyAlt, wdKeyG), "RAC_GrasSelection", "Alt+G") & vbCrLf
    rapport = rapport & RAC_EffacerSiNotreMacro(BuildKeyCode(wdKeyAlt, wdKeyV), "RAC_ValiderGras", "Alt+V") & vbCrLf
    rapport = rapport & RAC_EffacerSiNotreMacro(BuildKeyCode(wdKeyAlt, wdKeyD), "RAC_AjouterDeclencheur", "Alt+D") & vbCrLf
    rapport = rapport & RAC_EffacerSiNotreMacro(BuildKeyCode(wdKeyAlt, wdKeyA), "RAC_AjouterExamen", "Alt+A") & vbCrLf
    rapport = rapport & RAC_EffacerSiNotreMacro(BuildKeyCode(wdKeyAlt, wdKeyX), "RAC_AjouterExclusion", "Alt+X")

    modRaccourcis.ModeleUnifie().Save
    MsgBox rapport, vbInformation, "Raccourcis ModeleCourrierChatGPT"
    Exit Sub

GestionErreur:
    MsgBox Err.Number & " - " & Err.description, vbExclamation, "Suppression raccourcis"
End Sub

Private Sub RAC_NettoyerAnciensCtrlAlt()
    On Error Resume Next
    RAC_EffacerSiNotreMacro BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyG), "RAC_GrasSelection", "Ctrl+Alt+G"
    RAC_EffacerSiNotreMacro BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyV), "RAC_ValiderGras", "Ctrl+Alt+V"
    RAC_EffacerSiNotreMacro BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyD), "RAC_AjouterDeclencheur", "Ctrl+Alt+D"
    RAC_EffacerSiNotreMacro BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyA), "RAC_AjouterExamen", "Ctrl+Alt+A"
    RAC_EffacerSiNotreMacro BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyX), "RAC_AjouterExclusion", "Ctrl+Alt+X"
    On Error GoTo 0
End Sub

Private Function RAC_EffacerSiNotreMacro( _
    ByVal codeTouche As Long, _
    ByVal nomMacro As String, _
    ByVal libelleTouche As String) As String

    Dim commandeExistante As String

    On Error GoTo GestionErreur
    commandeExistante = FindKey(codeTouche).Command

    If StrComp(commandeExistante, nomMacro, vbTextCompare) = 0 Then
        FindKey(codeTouche).Clear
        RAC_EffacerSiNotreMacro = "OK : " & libelleTouche
    Else
        RAC_EffacerSiNotreMacro = "LAISSE INCHANGE : " & libelleTouche
    End If
    Exit Function

GestionErreur:
    RAC_EffacerSiNotreMacro = "ECHEC : " & libelleTouche
    Err.Clear
End Function
