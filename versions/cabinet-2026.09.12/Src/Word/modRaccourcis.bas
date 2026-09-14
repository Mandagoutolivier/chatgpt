Attribute VB_Name = "modRaccourcis"
Option Explicit
' =====================================================================
' modRaccourcis - Aide et reinstallation des raccourcis clavier.
' Les raccourcis sont normalement enregistres dans CabinetUnifie.dotm par le
' build ; cette macro les repose au besoin, et AideCabinet les rappelle.
' Les commandes vocales Dragon envoient ces memes raccourcis.
' =====================================================================

Public Sub AideCabinet()
    MsgBox "PowerMic A : patients arrives puis destinataire Dragon." & vbCrLf & _
           "B : formule d appel, puis dictee du courrier." & vbCrLf & _
           "C / F6 : NOM Prenom, age, a la position du curseur." & vbCrLf & _
           "D : correction, annexes et gras ; relire puis D pour transmettre au secretariat." & vbCrLf & _
           "Reprendre : rouvrir un brouillon interrompu.", vbInformation, "Cabinet"
End Sub
' Execute automatiquement au chargement de CabinetUnifie.dotm (demarrage de Word) :
' verifie et repare les raccourcis sans jamais demander d'enregistrer le modele.
Public Sub AutoExec()
    On Error Resume Next
    modLog.LogInfo "AutoExec CabinetUnifie.dotm : demarrage"
    InstallerRaccourcisSession False
End Sub

' Sonde sans interface : prouve qu'une liaison de touche declenche bien une
' macro (ecrit un fichier temoin). Utilisee par les tests, jamais par l'utilisateur.
Public Sub SondeRaccourci()
    On Error Resume Next
    modFichiers.EnsureDossier Environ$("TEMP") & "\CabinetCardio"
    modFichiers.EcrireTexteAnsi Environ$("TEMP") & "\CabinetCardio\sonde_raccourci.txt", Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub ReinstallerRaccourcis()
    InstallerRaccourcisSession True
End Sub

Private Sub InstallerRaccourcisSession(ByVal verbeux As Boolean)
    On Error Resume Next
    Dim modele As Template, i As Long, repares As Long
    Dim touches As Variant, macros As Variant
    Set modele = TrouverModeleCabinet()
    If modele Is Nothing Then
        If verbeux Then MsgBox "CabinetUnifie.dotm n'est pas charge.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    ' noms NON qualifies : c'est la forme que Word resout pour un modele global
    ' Correction = Ctrl+Alt+MAJ+C : Ctrl+Alt+C est deja pris par l'ancien
    ' complement ModeleCourrierChatGPT_PROD.dotm, que le medecin utilise encore.
    Dim majuscule As Variant
    touches = Array(wdKeyN, wdKeyC, wdKeyD, wdKeyP, wdKeyG, wdKeyV, wdKeyB)
    majuscule = Array(False, True, False, False, False, False, False)
    macros = Array("Unifie_A_NouvelleLettre", "Unifie_D_Finaliser", "Unifie_D_Finaliser", _
                   "Unifie_C_InsererPatient", "EnvoyerECG", "Unifie_D_Finaliser", "MettreEnGras")
    Dim ancienContexte As Object
    Set ancienContexte = Application.CustomizationContext
    CustomizationContext = modele
    For i = 0 To UBound(touches)
        Dim code As Long
        If majuscule(i) Then
            code = BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, touches(i))
        Else
            code = BuildKeyCode(wdKeyControl, wdKeyAlt, touches(i))
        End If
        ' on repose systematiquement la liaison (la propriete Command d'un
        ' modele global se lit vide, elle ne permet pas de verifier)
        Err.Clear
        KeyBindings.Add wdKeyCategoryMacro, macros(i), code
        If Err.Number = 0 Then repares = repares + 1 Else modLog.LogErreur "raccourci " & macros(i) & " : " & Err.Description
    Next i
    ' boutons PowerMic : A = destinataire, B = appel, D = finaliser (Ctrl+Alt+Maj+...)
    Err.Clear: KeyBindings.Add wdKeyCategoryMacro, "Unifie_A_NouvelleLettre", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, wdKeyA)
    Err.Clear: KeyBindings.Add wdKeyCategoryMacro, "Unifie_B_FormuleAppel", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, wdKeyB)
    Err.Clear: KeyBindings.Add wdKeyCategoryMacro, "Unifie_D_Finaliser", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, wdKeyD)
    ' F6 = identite du patient (meme macro que Ctrl+Alt+P, plus rapide a la dictee)
    Err.Clear
    KeyBindings.Add wdKeyCategoryMacro, "InsererPatient", BuildKeyCode(wdKeyF6)
    If Err.Number = 0 Then repares = repares + 1
    CustomizationContext = ancienContexte
    modele.Saved = True          ' aucune invite d'enregistrement du modele
    If verbeux Then MsgBox "Raccourcis verifies : " & repares & " reinstalle(s) (Ctrl+Alt+N/D/P/G/V/B, Ctrl+Alt+Maj+C).", vbInformation, "Cabinet"
    modLog.LogInfo "Raccourcis (re)poses : " & repares & "/7"
End Sub

' Diagnostic a lancer depuis Alt+F8 en cas de probleme : etat de la configuration du poste
Public Sub DiagnosticCabinet()
    On Error GoTo Echec
    Dim r As Object, texte As String, t As Template
    Set r = modServiceNas.Appeler("whoami", modServiceNas.Parametres())
    texte = "Service NAS : " & CStr(r("revision")) & vbCrLf & "Protocole " & CStr(r("protocole")) & " / schema " & CStr(r("schema"))
    texte = texte & vbCrLf & "Modele de correction : " & modConfig.Config("API", "ModeleOpenAI", "(absent)")
    texte = texte & vbCrLf & "Modele de lettre : " & CStr(modFichiers.FichierExiste(modConfig.chemin("Modeles") & "\LETTRE TYPE.dot"))
    For Each t In Templates: texte = texte & vbCrLf & "Modele charge : " & t.Name: Next t
    MsgBox texte, vbInformation, "Diagnostic Cabinet"
    Exit Sub
Echec:
    modLog.Diagnostic "diagnostic_word", "echec", Err.Number
    MsgBox "Connexion ou configuration indisponible : " & Err.Description, vbExclamation, "Diagnostic Cabinet"
End Sub

Private Function TrouverModeleCabinet() As Template
    Dim t As Template
    For Each t In Templates
        If LCase$(t.Name) = "cabinetunifie.dotm" Then
            Set TrouverModeleCabinet = t
            Exit Function
        End If
    Next t
End Function

Public Function ModeleUnifie() As Template
    Set ModeleUnifie = TrouverModeleCabinet()
    If ModeleUnifie Is Nothing Then Err.Raise vbObjectError + 965, , "CabinetUnifie.dotm n est pas charge."
End Function
