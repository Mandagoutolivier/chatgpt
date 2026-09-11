Attribute VB_Name = "modNormalisationGrasLocal"
Option Explicit

'===============================================================================
' MODULE : modNormalisationGrasLocal
' VERSION : R5 - IDENTITE PATIENT EN GRAS LOCAL
'
' Traitement 100 % local, sans appel API :
' - lit les deux dictionnaires partagés du NAS ;
' - médicaments : forme canonique en MAJUSCULES + gras ;
' - expressions : gras uniquement, sans modifier la casse ;
' - s'applique au courrier principal ET aux lettres complémentaires déjà fusionnées.
' - impose localement le gras sur Monsieur/Madame NOM Prénom ;
' - l'identité du patient n'est jamais ajoutée aux dictionnaires NAS.
'
' Les dictionnaires restent sur le NAS afin que les apprentissages réalisés
' au secrétariat soient immédiatement utilisables sur les autres postes.
'===============================================================================

Private Const NGL_DOSSIER_CONFIG As String = _
    "\\DS224\home\logiciel_chatgpt_cabinet"

Private Const NGL_FICHIER_MEDICAMENTS As String = _
    "\\DS224\home\logiciel_chatgpt_cabinet\medicaments.txt"

Private Const NGL_FICHIER_EXPRESSIONS As String = _
    "\\DS224\home\logiciel_chatgpt_cabinet\termes_gras.txt"

Private Const NGL_VAR_IDENTITE_PATIENT As String = _
    "MCP_IDENTITE_PATIENT_V1"

Public Function NGL_NormaliserDocument(ByVal doc As Document) As Boolean
    Call modGras.AppliquerGrasDocumentComplet(doc)
    NGL_NormaliserDocument = True
End Function
Public Sub NGL_TesterConfiguration()
    MsgBox CStr(modGras.NbMedicaments()) & " medicaments et " & CStr(modGras.NbExpressions()) & " expressions dans les dictionnaires Cabinet.", vbInformation, "Gras"
End Sub
Private Function NGL_ConfigurationDisponible(ByVal afficherErreur As Boolean) As Boolean

    Dim fso As Object
    Dim message As String

    On Error GoTo GestionErreur

    Set fso = CreateObject("Scripting.FileSystemObject")
    NGL_ConfigurationDisponible = False

    If Not fso.FolderExists(NGL_DOSSIER_CONFIG) Then
        message = "Dossier des dictionnaires introuvable :" & vbCrLf & NGL_DOSSIER_CONFIG
        GoTo Echec
    End If

    If Not fso.FileExists(NGL_FICHIER_MEDICAMENTS) Then
        message = "Dictionnaire des médicaments introuvable :" & vbCrLf & NGL_FICHIER_MEDICAMENTS
        GoTo Echec
    End If

    If Not fso.FileExists(NGL_FICHIER_EXPRESSIONS) Then
        message = "Dictionnaire des expressions introuvable :" & vbCrLf & NGL_FICHIER_EXPRESSIONS
        GoTo Echec
    End If

    NGL_ConfigurationDisponible = True
    Exit Function

Echec:
    If afficherErreur Then
        MsgBox message, vbExclamation, "Mise en gras locale"
    End If
    Exit Function

GestionErreur:
    If afficherErreur Then
        MsgBox _
            "Impossible d'accéder aux dictionnaires :" & vbCrLf & _
            Err.Number & " - " & Err.description, _
            vbExclamation, _
            "Mise en gras locale"
    End If
End Function

Private Function NGL_ChargerDictionnaire( _
    ByVal chemin As String, _
    ByVal estMedicament As Boolean) As Object

    Dim d As Object
    Dim numero As Integer
    Dim ligne As String
    Dim parties As Variant
    Dim partie As Variant
    Dim canonique As String
    Dim aliasTexte As String
    Dim cle As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    numero = FreeFile
    Open chemin For Input As #numero

    Do While Not EOF(numero)
        Line Input #numero, ligne
        ligne = Trim$(ligne)

        If ligne <> "" Then
            If Left$(ligne, 1) <> "#" Then
                parties = Split(ligne, "|")
                canonique = Trim$(CStr(parties(LBound(parties))))

                If canonique <> "" Then
                    If estMedicament Then canonique = UCase$(canonique)

                    For Each partie In parties
                        aliasTexte = Trim$(CStr(partie))
                        If aliasTexte <> "" Then
                            cle = LCase$(aliasTexte)
                            If Not d.Exists(cle) Then d.Add cle, canonique
                        End If
                    Next partie

                    cle = LCase$(canonique)
                    If Not d.Exists(cle) Then d.Add cle, canonique
                End If
            End If
        End If
    Loop

    Close #numero
    Set NGL_ChargerDictionnaire = d
End Function

Private Sub NGL_AppliquerTerme( _
    ByVal doc As Document, _
    ByVal texteRecherche As String, _
    ByVal texteCanonique As String, _
    ByVal modifierCasse As Boolean)

    Dim rngRecherche As Range
    Dim debutSuivant As Long
    Dim finDocument As Long
    Dim texteTrouve As String

    If doc Is Nothing Then Exit Sub
    If Len(Trim$(texteRecherche)) = 0 Then Exit Sub

    debutSuivant = 0

    Do
        finDocument = doc.Content.End - 1
        If debutSuivant >= finDocument Then Exit Do

        Set rngRecherche = doc.Range(Start:=debutSuivant, End:=finDocument)

        With rngRecherche.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = texteRecherche
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
        End With

        If Not rngRecherche.Find.Execute Then Exit Do

        If NGL_BornesTermeValides(doc, rngRecherche, texteRecherche) Then
            texteTrouve = rngRecherche.Text

            If modifierCasse Then
                If StrComp(texteTrouve, texteCanonique, vbBinaryCompare) <> 0 Then
                    rngRecherche.Text = texteCanonique
                End If
            End If

            rngRecherche.Font.Bold = True
            debutSuivant = rngRecherche.End
        Else
            debutSuivant = rngRecherche.Start + 1
        End If
    Loop
End Sub

Private Function NGL_BornesTermeValides( _
    ByVal doc As Document, _
    ByVal rngTrouve As Range, _
    ByVal texteRecherche As String) As Boolean

    Dim avant As String
    Dim apres As String
    Dim premier As String
    Dim dernier As String

    NGL_BornesTermeValides = True
    If Len(texteRecherche) = 0 Then Exit Function

    premier = Left$(texteRecherche, 1)
    dernier = Right$(texteRecherche, 1)

    If NGL_EstCaractereMot(premier) Then
        If rngTrouve.Start > 0 Then
            avant = doc.Range(Start:=rngTrouve.Start - 1, End:=rngTrouve.Start).Text
            If NGL_EstCaractereMot(avant) Then
                NGL_BornesTermeValides = False
                Exit Function
            End If
        End If
    End If

    If NGL_EstCaractereMot(dernier) Then
        If rngTrouve.End < doc.Content.End - 1 Then
            apres = doc.Range(Start:=rngTrouve.End, End:=rngTrouve.End + 1).Text
            If NGL_EstCaractereMot(apres) Then
                NGL_BornesTermeValides = False
                Exit Function
            End If
        End If
    End If
End Function

Private Function NGL_EstCaractereMot(ByVal caractere As String) As Boolean
    Dim c As String

    If Len(caractere) = 0 Then Exit Function
    c = Left$(caractere, 1)

    ' Evite les classes Like avec plages accentuees, incompatibles sur
    ' certaines installations Word/VBA (erreur 93 : format de chaine incorrect).
    If c Like "[0-9A-Za-z]" Then
        NGL_EstCaractereMot = True
    ElseIf LCase$(c) <> UCase$(c) Then
        ' Toute lettre alphabetique Unicode, y compris accentuee.
        NGL_EstCaractereMot = True
    Else
        NGL_EstCaractereMot = False
    End If
End Function


'-------------------------------------------------------------------------------
' IDENTITE DU PATIENT : REGLE STRUCTURELLE LOCALE
'-------------------------------------------------------------------------------

Private Sub NGL_MettreIdentitePatientEnGras(ByVal doc As Document)

    Dim identite As String

    If doc Is Nothing Then Exit Sub

    On Error Resume Next
    identite = Trim$(gPatient.NomComplet)
    On Error GoTo 0

    If identite = "" Then
        identite = NGL_LireIdentitePatientDocument(doc)
    End If

    If identite = "" Then
        identite = NGL_DetecterIdentitePatientDocument(doc)
    End If

    If identite = "" Then Exit Sub

    identite = NGL_CanonicaliserIdentitePatient(identite)
    If identite = "" Then Exit Sub

    NGL_MemoriserIdentitePatient doc, identite
    NGL_MettreIdentiteDansDocumentEnGras doc, identite

End Sub

Private Function NGL_LireIdentitePatientDocument( _
    ByVal doc As Document) As String

    On Error Resume Next
    NGL_LireIdentitePatientDocument = _
        Trim$(doc.Variables(NGL_VAR_IDENTITE_PATIENT).Value)
    On Error GoTo 0

End Function

Private Sub NGL_MemoriserIdentitePatient( _
    ByVal doc As Document, _
    ByVal identite As String)

    If doc Is Nothing Then Exit Sub
    identite = Trim$(identite)
    If identite = "" Then Exit Sub

    On Error Resume Next
    doc.Variables(NGL_VAR_IDENTITE_PATIENT).Delete
    On Error GoTo 0

    On Error Resume Next
    doc.Variables.Add _
        Name:=NGL_VAR_IDENTITE_PATIENT, _
        Value:=identite
    On Error GoTo 0

End Sub

Private Function NGL_DetecterIdentitePatientDocument( _
    ByVal doc As Document) As String

    Dim p As Paragraph
    Dim identite As String

    If doc Is Nothing Then Exit Function

    For Each p In doc.Paragraphs
        identite = NGL_ExtraireIdentitePatientParagraphe(p.Range.Text)
        If identite <> "" Then
            NGL_DetecterIdentitePatientDocument = identite
            Exit Function
        End If
    Next p

End Function

Private Function NGL_ExtraireIdentitePatientParagraphe( _
    ByVal texte As String) As String

    Dim t As String
    Dim tMin As String
    Dim posCivilite As Long
    Dim longueurCivilite As Long
    Dim posFin As Long
    Dim candidat As String

    t = NGL_NettoyerTexteIdentite(texte)
    If t = "" Then Exit Function

    tMin = LCase$(t)

    posCivilite = InStr(1, tMin, "monsieur ", vbTextCompare)
    longueurCivilite = Len("Monsieur")

    If posCivilite = 0 Then
        posCivilite = InStr(1, tMin, "madame ", vbTextCompare)
        longueurCivilite = Len("Madame")
    End If

    If posCivilite = 0 Then
        posCivilite = InStr(1, tMin, "mademoiselle ", vbTextCompare)
        longueurCivilite = Len("Mademoiselle")
    End If

    If posCivilite = 0 Then Exit Function

    'Pour une détection de secours, on limite la recherche aux paragraphes
    'qui ressemblent réellement à une présentation du patient.
    If InStr(1, tMin, " ans", vbTextCompare) = 0 _
    And InStr(1, tMin, "je revois", vbTextCompare) = 0 _
    And InStr(1, tMin, "je vois", vbTextCompare) = 0 _
    And InStr(1, tMin, "merci de réaliser", vbTextCompare) = 0 _
    And InStr(1, tMin, "merci de realiser", vbTextCompare) = 0 _
    And InStr(1, tMin, "merci de prendre en charge", vbTextCompare) = 0 Then
        Exit Function
    End If

    posFin = InStr(posCivilite + longueurCivilite + 1, t, ",")
    If posFin = 0 Then posFin = InStr(posCivilite + longueurCivilite + 1, t, ";")
    If posFin = 0 Then posFin = InStr(posCivilite + longueurCivilite + 1, t, ".")
    If posFin = 0 Then Exit Function

    candidat = Trim$(Mid$(t, posCivilite, posFin - posCivilite))
    candidat = NGL_CanonicaliserIdentitePatient(candidat)

    If NGL_IdentiteCandidateValide(candidat) Then
        NGL_ExtraireIdentitePatientParagraphe = candidat
    End If

End Function

Private Function NGL_IdentiteCandidateValide( _
    ByVal candidat As String) As Boolean

    Dim morceaux As Variant

    candidat = Trim$(candidat)
    If Len(candidat) < 10 Or Len(candidat) > 100 Then Exit Function

    morceaux = Split(candidat, " ")
    If UBound(morceaux) < 2 Then Exit Function

    If InStr(1, LCase$(candidat), "docteur", vbTextCompare) > 0 Then Exit Function
    If InStr(1, candidat, "[[PATIENT]]", vbTextCompare) > 0 Then Exit Function

    NGL_IdentiteCandidateValide = True

End Function

Private Sub NGL_MettreIdentiteDansDocumentEnGras( _
    ByVal doc As Document, _
    ByVal identite As String)

    Dim civilite As String
    Dim nomPrenom As String
    Dim zone As Range
    Dim rngNom As Range
    Dim rngCivilite As Range
    Dim texteNormalise As String
    Dim nomNormalise As String
    Dim positionNom As Long
    Dim debutZone As Long
    Dim debutContexte As Long
    Dim contexte As String
    Dim positionCiviliteContexte As Long
    Dim positionCiviliteDocument As Long
    Dim separateur As String

    If doc Is Nothing Then Exit Sub

    If Not NGL_DecomposerIdentitePatient( _
        identite, civilite, nomPrenom) Then Exit Sub

    Set zone = doc.Content.Duplicate
    If zone.End > zone.Start Then zone.End = zone.End - 1

    debutZone = zone.Start
    texteNormalise = NGL_NormaliserLongueurConstante(zone.Text)
    nomNormalise = NGL_NormaliserLongueurConstante(nomPrenom)

    positionNom = 1

    Do
        positionNom = InStr(positionNom, texteNormalise, nomNormalise, vbTextCompare)
        If positionNom = 0 Then Exit Do

        debutContexte = positionNom - Len(civilite) - 8
        If debutContexte < 1 Then debutContexte = 1

        contexte = Mid$( _
            texteNormalise, _
            debutContexte, _
            positionNom - debutContexte)

        positionCiviliteContexte = _
            InStrRev(contexte, civilite, -1, vbTextCompare)

        If positionCiviliteContexte > 0 Then

            separateur = Mid$( _
                contexte, _
                positionCiviliteContexte + Len(civilite))

            If NGL_SeparateurIdentiteValide(separateur) Then

                positionCiviliteDocument = _
                    debutContexte + positionCiviliteContexte - 1

                Set rngCivilite = doc.Range( _
                    Start:=debutZone + positionCiviliteDocument - 1, _
                    End:=debutZone + positionCiviliteDocument - 1 + Len(civilite))

                Set rngNom = doc.Range( _
                    Start:=debutZone + positionNom - 1, _
                    End:=debutZone + positionNom - 1 + Len(nomPrenom))

                rngCivilite.Font.Bold = True
                rngNom.Font.Bold = True
            End If
        End If

        positionNom = positionNom + Len(nomPrenom)
    Loop

End Sub

Private Function NGL_DecomposerIdentitePatient( _
    ByVal identite As String, _
    ByRef civilite As String, _
    ByRef nomPrenom As String) As Boolean

    Dim t As String

    t = NGL_CanonicaliserIdentitePatient(identite)
    civilite = ""
    nomPrenom = ""

    If StrComp(Left$(t, 9), "Monsieur ", vbTextCompare) = 0 Then
        civilite = "Monsieur"
        nomPrenom = Trim$(Mid$(t, 10))
    ElseIf StrComp(Left$(t, 7), "Madame ", vbTextCompare) = 0 Then
        civilite = "Madame"
        nomPrenom = Trim$(Mid$(t, 8))
    ElseIf StrComp(Left$(t, 13), "Mademoiselle ", vbTextCompare) = 0 Then
        civilite = "Mademoiselle"
        nomPrenom = Trim$(Mid$(t, 14))
    End If

    If civilite = "" Or nomPrenom = "" Then Exit Function

    NGL_DecomposerIdentitePatient = True

End Function

Private Function NGL_SeparateurIdentiteValide( _
    ByVal texte As String) As Boolean

    Dim i As Long
    Dim c As String

    NGL_SeparateurIdentiteValide = True

    For i = 1 To Len(texte)
        c = Mid$(texte, i, 1)

        If c <> " " _
        And c <> Chr(160) _
        And c <> ChrW(8239) _
        And c <> "«" _
        And c <> "»" _
        And c <> Chr(34) _
        And c <> "“" _
        And c <> "”" Then

            NGL_SeparateurIdentiteValide = False
            Exit Function
        End If
    Next i

End Function

Private Function NGL_CanonicaliserIdentitePatient( _
    ByVal identite As String) As String

    identite = NGL_NormaliserLongueurConstante(identite)
    identite = Replace(identite, "«", " ")
    identite = Replace(identite, "»", " ")
    identite = Replace(identite, Chr(34), " ")
    identite = Replace(identite, "“", " ")
    identite = Replace(identite, "”", " ")

    Do While InStr(identite, "  ") > 0
        identite = Replace(identite, "  ", " ")
    Loop

    NGL_CanonicaliserIdentitePatient = Trim$(identite)

End Function

Private Function NGL_NormaliserLongueurConstante( _
    ByVal texte As String) As String

    texte = Replace(texte, Chr(160), " ")
    texte = Replace(texte, ChrW(8239), " ")
    texte = Replace(texte, vbTab, " ")
    texte = Replace(texte, vbCr, " ")
    texte = Replace(texte, vbLf, " ")

    NGL_NormaliserLongueurConstante = texte

End Function

Private Function NGL_NettoyerTexteIdentite( _
    ByVal texte As String) As String

    texte = NGL_NormaliserLongueurConstante(texte)

    Do While InStr(texte, "  ") > 0
        texte = Replace(texte, "  ", " ")
    Loop

    NGL_NettoyerTexteIdentite = Trim$(texte)

End Function
