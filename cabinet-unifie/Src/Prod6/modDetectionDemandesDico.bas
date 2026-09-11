Attribute VB_Name = "modDetectionDemandesDico"
Option Explicit
Option Private Module

'===============================================================================
' MODULE : modDetectionDemandesDico
' VERSION : DDE2-R4 CORE - lecture robuste CRLF/LF/CR
'
' Fichiers partages :
'   \\DS224\home\logiciel_chatgpt_cabinet\declencheurs_demandes.txt
'   \\DS224\home\logiciel_chatgpt_cabinet\examens_complementaires.txt
'   \\DS224\home\logiciel_chatgpt_cabinet\exclusions_demandes.txt
'
' Module interne. Les six macros visibles sont exposees par
' modDDECommandesPubliques afin de rester accessibles meme si le projet est protege.
' Le caractere * dans un motif signifie : n'importe quel texte intermediaire.
' Exemple : je complete le bilan*par*
'===============================================================================

Public Function DDE_ContientDemandeEligible(ByVal texte As String) As Boolean
    Dim phrases As Collection
    Dim element As Variant
    Dim p As String

    DDE_ContientDemandeEligible = False
    Set phrases = DDE_DecouperPassages(texte)

    For Each element In phrases
        p = DDE_Normaliser(CStr(element))
        If Len(p) > 0 Then
            If DDE_ContientDeclencheur(p) And DDE_ContientExamen(p) Then
                If Not DDE_ContientExclusion(p) Then
                    DDE_ContientDemandeEligible = True
                    Exit Function
                End If
            End If
        End If
    Next element
End Function

Public Function DDE_ContientDeclencheur(ByVal texte As String) As Boolean
    DDE_ContientDeclencheur = DDE_ContientMotifDepuisFichier( _
        DDE_Normaliser(texte), DDE_FICHIER_DECLENCHEURS, False)
End Function

Public Function DDE_ContientExamen(ByVal texte As String) As Boolean
    DDE_ContientExamen = DDE_ContientMotifDepuisFichier( _
        DDE_Normaliser(texte), DDE_FICHIER_EXAMENS, True)
End Function

Public Function DDE_ContientExclusion(ByVal texte As String) As Boolean
    DDE_ContientExclusion = DDE_ContientMotifDepuisFichier( _
        DDE_Normaliser(texte), DDE_FICHIER_EXCLUSIONS, False)
End Function

Private Function DDE_ContientMotifDepuisFichier( _
    ByVal texteNormalise As String, _
    ByVal chemin As String, _
    ByVal prendreAvantSeparateur As Boolean) As Boolean

    Dim contenu As String
    Dim lignes() As String
    Dim ligne As String
    Dim motif As String
    Dim p As Long
    Dim i As Long

    DDE_ContientMotifDepuisFichier = False
    If Len(texteNormalise) = 0 Then Exit Function
    If Not modFichiers.FichierExiste(chemin) Then Err.Raise vbObjectError + 961, "modDetectionDemandesDico", "Dictionnaire de demandes absent : " & chemin

    contenu = DDE_LireToutFichier(chemin)
    If Len(contenu) = 0 Then Exit Function

    ' Accepte indifferemment les fichiers Windows CRLF, Unix LF ou anciens CR.
    contenu = Replace(contenu, vbCrLf, vbLf)
    contenu = Replace(contenu, vbCr, vbLf)
    lignes = Split(contenu, vbLf)

    For i = LBound(lignes) To UBound(lignes)
        ligne = Trim$(lignes(i))
        If Len(ligne) > 0 Then
            If Left$(ligne, 1) <> "#" Then
                motif = ""
                If prendreAvantSeparateur Then
                    p = InStr(1, ligne, "|", vbBinaryCompare)
                    If p > 0 Then motif = Left$(ligne, p - 1) Else motif = ligne
                Else
                    motif = ligne
                End If

                motif = DDE_Normaliser(motif)
                If Len(motif) > 0 Then
                    If DDE_MotifCorrespond(texteNormalise, motif) Then
                        DDE_ContientMotifDepuisFichier = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Function DDE_LireToutFichier(ByVal chemin As String) As String
    DDE_LireToutFichier = modFichiers.LireTexteUTF8(chemin)
End Function
Private Function DDE_MotifCorrespond( _
    ByVal texte As String, _
    ByVal motif As String) As Boolean

    Dim morceaux() As String
    Dim i As Long
    Dim position As Long
    Dim trouve As Long
    Dim morceau As String

    DDE_MotifCorrespond = False
    If Len(motif) = 0 Then Exit Function

    If InStr(1, motif, "*", vbBinaryCompare) = 0 Then
        DDE_MotifCorrespond = (InStr(1, texte, motif, vbTextCompare) > 0)
        Exit Function
    End If

    morceaux = Split(motif, "*")
    position = 1

    For i = LBound(morceaux) To UBound(morceaux)
        morceau = Trim$(morceaux(i))
        If Len(morceau) > 0 Then
            trouve = InStr(position, texte, morceau, vbTextCompare)
            If trouve = 0 Then Exit Function
            position = trouve + Len(morceau)
        End If
    Next i

    DDE_MotifCorrespond = True
End Function

Private Function DDE_DecouperPassages(ByVal texte As String) As Collection
    Dim c As New Collection
    Dim t As String
    Dim morceaux() As String
    Dim i As Long

    t = texte
    t = Replace(t, vbCrLf, ".")
    t = Replace(t, vbCr, ".")
    t = Replace(t, vbLf, ".")
    t = Replace(t, "!", ".")
    t = Replace(t, "?", ".")
    t = Replace(t, ";", ".")

    morceaux = Split(t, ".")
    For i = LBound(morceaux) To UBound(morceaux)
        If Len(Trim$(morceaux(i))) > 0 Then c.Add Trim$(morceaux(i))
    Next i

    Set DDE_DecouperPassages = c
End Function

Private Function DDE_Normaliser(ByVal texte As String) As String
    texte = LCase$(texte)
    texte = Replace(texte, ChrW(8217), "'")
    texte = Replace(texte, ChrW(8216), "'")
    texte = Replace(texte, ChrW(160), " ")
    texte = Replace(texte, vbTab, " ")

    texte = Replace(texte, "à", "a")
    texte = Replace(texte, "â", "a")
    texte = Replace(texte, "ä", "a")
    texte = Replace(texte, "á", "a")
    texte = Replace(texte, "ã", "a")
    texte = Replace(texte, "ç", "c")
    texte = Replace(texte, "é", "e")
    texte = Replace(texte, "è", "e")
    texte = Replace(texte, "ê", "e")
    texte = Replace(texte, "ë", "e")
    texte = Replace(texte, "î", "i")
    texte = Replace(texte, "ï", "i")
    texte = Replace(texte, "í", "i")
    texte = Replace(texte, "ô", "o")
    texte = Replace(texte, "ö", "o")
    texte = Replace(texte, "ó", "o")
    texte = Replace(texte, "õ", "o")
    texte = Replace(texte, "ù", "u")
    texte = Replace(texte, "û", "u")
    texte = Replace(texte, "ü", "u")
    texte = Replace(texte, "ú", "u")
    texte = Replace(texte, "ÿ", "y")
    texte = Replace(texte, "œ", "oe")

    Do While InStr(1, texte, "  ", vbBinaryCompare) > 0
        texte = Replace(texte, "  ", " ")
    Loop

    DDE_Normaliser = Trim$(texte)
End Function

Public Sub DDE_Core_AjouterDeclencheurSelection()
    DDE_AjouterDepuisSelection DDE_FICHIER_DECLENCHEURS, _
        "Ajouter un declencheur", _
        "Modifiez si besoin le motif. Utilisez * pour le texte variable."
End Sub

Public Sub DDE_Core_AjouterExamenSelection()
    DDE_AjouterDepuisSelection DDE_FICHIER_EXAMENS, _
        "Ajouter un examen", _
        "Format conseille : expression|TYPE_EXAMEN"
End Sub

Public Sub DDE_Core_AjouterExclusionSelection()
    DDE_AjouterDepuisSelection DDE_FICHIER_EXCLUSIONS, _
        "Ajouter une exclusion", _
        "Cette formulation bloquera une demande si elle se trouve dans le meme passage."
End Sub

Private Sub DDE_AjouterDepuisSelection( _
    ByVal chemin As String, _
    ByVal titre As String, _
    ByVal instruction As String)

    Dim proposition As String
    Dim valeur As String

    proposition = Selection.Text
    proposition = Replace(proposition, vbCr, " ")
    proposition = Replace(proposition, vbLf, " ")
    proposition = Trim$(proposition)

    valeur = InputBox(instruction, titre, proposition)
    valeur = Trim$(valeur)
    If Len(valeur) = 0 Then Exit Sub

    If DDE_LigneExiste(chemin, valeur) Then
        MsgBox "Cette ligne existe deja dans le dictionnaire.", vbInformation, titre
        Exit Sub
    End If

    If Not DDE_AjouterLigne(chemin, valeur) Then
        MsgBox "Impossible de modifier :" & vbCrLf & chemin, vbExclamation, titre
        Exit Sub
    End If

    MsgBox "Ajout effectue :" & vbCrLf & valeur, vbInformation, titre
End Sub

Private Function DDE_AjouterLigne(ByVal chemin As String, ByVal valeur As String) As Boolean
    Dim verrou As String, temporaire As String, contenu As String, acquis As Boolean
    On Error GoTo Echec
    If InStr(valeur, vbCr) > 0 Or InStr(valeur, vbLf) > 0 Then Err.Raise vbObjectError + 963, , "Une seule ligne est attendue."
    verrou = "DDE_" & CreateObject("Scripting.FileSystemObject").GetBaseName(chemin)
    acquis = modFichiers.AcquerirVerrou(verrou)
    If Not acquis Then Err.Raise vbObjectError + 964, , "Dictionnaire occupe."
    If Not DDE_LigneExiste(chemin, valeur) Then
        contenu = DDE_LireToutFichier(chemin)
        temporaire = chemin & "." & modFichiers.IdUnique() & ".tmp"
        modFichiers.EcrireTexteUTF8 temporaire, contenu & vbCrLf & valeur & vbCrLf
        modFichiers.RenommerAtomique temporaire, chemin, True
    End If
    modFichiers.RelacherVerrou verrou
    DDE_AjouterLigne = True
    Exit Function
Echec:
    If acquis Then modFichiers.RelacherVerrou verrou
    modFichiers.SupprimerTemporaire temporaire
End Function
Private Function DDE_LigneExiste(ByVal chemin As String, ByVal valeur As String) As Boolean
    Dim contenu As String
    Dim lignes() As String
    Dim cible As String
    Dim i As Long

    DDE_LigneExiste = False
    If Not modFichiers.FichierExiste(chemin) Then Err.Raise vbObjectError + 961, "modDetectionDemandesDico", "Dictionnaire de demandes absent : " & chemin

    cible = DDE_Normaliser(valeur)
    contenu = DDE_LireToutFichier(chemin)
    If Len(contenu) = 0 Then Exit Function

    contenu = Replace(contenu, vbCrLf, vbLf)
    contenu = Replace(contenu, vbCr, vbLf)
    lignes = Split(contenu, vbLf)

    For i = LBound(lignes) To UBound(lignes)
        If DDE_Normaliser(Trim$(lignes(i))) = cible Then
            DDE_LigneExiste = True
            Exit Function
        End If
    Next i
End Function

Private Function DDE_CompterLignesUtiles(ByVal chemin As String) As Long
    Dim contenu As String
    Dim lignes() As String
    Dim ligne As String
    Dim i As Long

    DDE_CompterLignesUtiles = 0
    If Not modFichiers.FichierExiste(chemin) Then Err.Raise vbObjectError + 961, "modDetectionDemandesDico", "Dictionnaire de demandes absent : " & chemin

    contenu = DDE_LireToutFichier(chemin)
    If Len(contenu) = 0 Then Exit Function

    contenu = Replace(contenu, vbCrLf, vbLf)
    contenu = Replace(contenu, vbCr, vbLf)
    lignes = Split(contenu, vbLf)

    For i = LBound(lignes) To UBound(lignes)
        ligne = Trim$(lignes(i))
        If Len(ligne) > 0 Then
            If Left$(ligne, 1) <> "#" Then
                DDE_CompterLignesUtiles = DDE_CompterLignesUtiles + 1
            End If
        End If
    Next i
End Function

Public Sub DDE_Core_TesterConfiguration()
    Dim msg As String

    msg = "Dictionnaires de detection :" & vbCrLf & vbCrLf

    If Dir$(DDE_FICHIER_DECLENCHEURS) <> "" Then
        msg = msg & "OK (" & CStr(DDE_CompterLignesUtiles(DDE_FICHIER_DECLENCHEURS)) & _
            " lignes) - " & DDE_FICHIER_DECLENCHEURS & vbCrLf
    Else
        msg = msg & "ABSENT - " & DDE_FICHIER_DECLENCHEURS & vbCrLf
    End If

    If Dir$(DDE_FICHIER_EXAMENS) <> "" Then
        msg = msg & "OK (" & CStr(DDE_CompterLignesUtiles(DDE_FICHIER_EXAMENS)) & _
            " lignes) - " & DDE_FICHIER_EXAMENS & vbCrLf
    Else
        msg = msg & "ABSENT - " & DDE_FICHIER_EXAMENS & vbCrLf
    End If

    If Dir$(DDE_FICHIER_EXCLUSIONS) <> "" Then
        msg = msg & "OK (" & CStr(DDE_CompterLignesUtiles(DDE_FICHIER_EXCLUSIONS)) & _
            " lignes) - " & DDE_FICHIER_EXCLUSIONS
    Else
        msg = msg & "ABSENT - " & DDE_FICHIER_EXCLUSIONS
    End If

    MsgBox msg, vbInformation, "Detection demandes DDE-1 V2"
End Sub

Public Sub DDE_Core_TesterFormulationsInitiales()
    Dim tests As Variant
    Dim i As Long
    Dim ok As Long
    Dim msg As String

    tests = Array( _
        "Je lui demande de realiser un test d'effort.", _
        "Il va completer ce bilan par une scintigraphie myocardique.", _
        "Il completera par un score calcique.", _
        "Nous poursuivons les investigations en realisant un scanner thoracique sans injection.", _
        "Il va realiser un test d'effort.", _
        "Je lui prescris un coroscanner.", _
        "Je lui demande un avis neurologique.", _
        "Je l'adresse au gastro-enterologue pour avis.", _
        "Je le dirige vers un neurologue pour avis.", _
        "Je complete le bilan de ce terrain a risque par un test d'effort.")

    For i = LBound(tests) To UBound(tests)
        If DDE_ContientDemandeEligible(CStr(tests(i))) Then
            ok = ok + 1
        Else
            msg = msg & "NON DETECTE : " & CStr(tests(i)) & vbCrLf
        End If
    Next i

    If ok = UBound(tests) - LBound(tests) + 1 Then
        MsgBox "OK : toutes les formulations initiales sont detectees.", _
            vbInformation, "Detection demandes"
    Else
        MsgBox CStr(ok) & "/" & CStr(UBound(tests) - LBound(tests) + 1) & _
            " formulations detectees." & vbCrLf & vbCrLf & msg, _
            vbExclamation, "Detection demandes"
    End If
End Sub

Public Sub DDE_Core_TesterFauxPositifs()
    Dim t1 As String
    Dim t2 As String
    t1 = "Il avait realise un test d'effort en 2024, qui etait normal."
    t2 = "Une scintigraphie myocardique avait ete realisee en 2025."

    If Not DDE_ContientDemandeEligible(t1) And Not DDE_ContientDemandeEligible(t2) Then
        MsgBox "OK : les deux exemples historiques ne declenchent pas de demande.", _
            vbInformation, "Detection demandes"
    Else
        MsgBox "ECHEC : au moins un exemple historique declenche une demande.", _
            vbExclamation, "Detection demandes"
    End If
End Sub

Private Function DDE_FICHIER_DECLENCHEURS() As String
    DDE_FICHIER_DECLENCHEURS = modConfig.CheminNasConfigure("PROD6", "DossierDictionnaires", "Config\DDE") & "\declencheurs_demandes.txt"
End Function

Private Function DDE_FICHIER_EXAMENS() As String
    DDE_FICHIER_EXAMENS = modConfig.CheminNasConfigure("PROD6", "DossierDictionnaires", "Config\DDE") & "\examens_complementaires.txt"
End Function

Private Function DDE_FICHIER_EXCLUSIONS() As String
    DDE_FICHIER_EXCLUSIONS = modConfig.CheminNasConfigure("PROD6", "DossierDictionnaires", "Config\DDE") & "\exclusions_demandes.txt"
End Function
