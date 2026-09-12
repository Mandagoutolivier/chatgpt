Attribute VB_Name = "modAnonymisation"
Option Explicit

Public Function EstFormuleAppel(ByVal s As String) As Boolean

    Dim t As String

    t = LCase(NettoyerTexteParagraphe(s))

    EstFormuleAppel = False

    If t = "" Then Exit Function
    If Len(t) > 120 Then Exit Function

    If CommencePar(t, "cher") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "chère") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "chere") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "mon cher") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "ma chère") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "ma chere") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "mes chers") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "mes chères") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "mes cheres") Then EstFormuleAppel = True: Exit Function

    If CommencePar(t, "madame,") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "monsieur,") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "madame le docteur") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "monsieur le docteur") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "maître") Then EstFormuleAppel = True: Exit Function
    If CommencePar(t, "maitre") Then EstFormuleAppel = True: Exit Function

End Function

Public Function EstParagraphePatient(ByVal s As String) As Boolean

    Dim t As String

    t = LCase(NettoyerTexteParagraphe(s))

    EstParagraphePatient = False

    If t = "" Then Exit Function

    If InStr(1, t, "monsieur", vbTextCompare) > 0 Then
        EstParagraphePatient = True
        Exit Function
    End If

    If InStr(1, t, "madame", vbTextCompare) > 0 Then
        EstParagraphePatient = True
        Exit Function
    End If

End Function

Public Function EstSignature(ByVal s As String) As Boolean

    Dim t As String

    t = LCase(NettoyerTexteParagraphe(s))

    EstSignature = False

    If InStr(1, t, "docteur olivier mandagout", vbTextCompare) > 0 Then
        EstSignature = True
        Exit Function
    End If

    If InStr(1, t, "dr olivier mandagout", vbTextCompare) > 0 Then
        EstSignature = True
        Exit Function
    End If

End Function

Public Function LocaliserCorpsCourrier(ByVal doc As Document, _
                                        ByRef rngCorps As Range, _
                                        ByRef premierParagraphePatient As Range) As Boolean

    Dim i As Long
    Dim idxDate As Long
    Dim idxAppel As Long
    Dim idxDebutRecherche As Long
    Dim idxDebutCorps As Long
    Dim idxSignature As Long
    Dim idxPolitesse As Long
    Dim idxFinCorps As Long
    Dim textePara As String

    LocaliserCorpsCourrier = False
    If Len(Trim$(modIntegrationUnifie.VariableDoc(doc, "PatientID"))) > 0 Then
        LocaliserCorpsCourrier = modIntegrationUnifie.LocaliserCorpsUnifie(doc, rngCorps, premierParagraphePatient)
        Exit Function
    End If

    idxDate = 0
    idxAppel = 0
    idxDebutRecherche = 0
    idxDebutCorps = 0
    idxSignature = 0
    idxPolitesse = 0
    idxFinCorps = 0

    '====================================================
    ' 1. Méthode principale : recherche de la ligne de date
    '    Beaumont ... le ... année.
    '====================================================

    For i = 1 To doc.Paragraphs.Count

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(i).Range.Text)

        If EstParagrapheDateCourrier(textePara) Then
            idxDate = i
            Exit For
        End If

    Next i

    If idxDate > 0 Then
        idxDebutRecherche = idxDate + 1
    End If

    '====================================================
    ' 2. Méthode de secours : ancienne recherche par
    '    formule d'appel, si la date n'a pas été trouvée.
    '====================================================

    If idxDebutRecherche = 0 Then

        For i = 1 To doc.Paragraphs.Count

            textePara = NettoyerTexteParagraphe(doc.Paragraphs(i).Range.Text)

            If EstFormuleAppel(textePara) Then
                idxAppel = i
                Exit For
            End If

        Next i

        If idxAppel = 0 Then Exit Function

        idxDebutRecherche = idxAppel + 1

    End If

    '====================================================
    ' 3. Recherche du premier paragraphe patient après
    '    la date ou après la formule d'appel de secours.
    '====================================================

    For i = idxDebutRecherche To doc.Paragraphs.Count

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(i).Range.Text)

        If EstParagraphePatient(textePara) Then
            idxDebutCorps = i
            Exit For
        End If

        If EstSignature(textePara) Then Exit Function

    Next i

    If idxDebutCorps = 0 Then Exit Function

    '====================================================
    ' 4. Recherche de la signature.
    '====================================================

    For i = idxDebutCorps To doc.Paragraphs.Count

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(i).Range.Text)

        If EstSignature(textePara) Then
            idxSignature = i
            Exit For
        End If

    Next i

    If idxSignature = 0 Then Exit Function

    '====================================================
    ' 5. Le paragraphe de politesse est le paragraphe non
    '    vide immédiatement avant la signature.
    '====================================================

    idxPolitesse = idxSignature - 1

    Do While idxPolitesse >= idxDebutCorps

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(idxPolitesse).Range.Text)

        If textePara <> "" Then Exit Do

        idxPolitesse = idxPolitesse - 1

    Loop

    If idxPolitesse < idxDebutCorps Then Exit Function

    '====================================================
    ' 6. Le corps s'arrête au paragraphe non vide avant
    '    la formule de politesse.
    '====================================================

    idxFinCorps = idxPolitesse - 1

    Do While idxFinCorps >= idxDebutCorps

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(idxFinCorps).Range.Text)

        If textePara <> "" Then Exit Do

        idxFinCorps = idxFinCorps - 1

    Loop

    If idxFinCorps < idxDebutCorps Then Exit Function

    Set rngCorps = doc.Range( _
        Start:=doc.Paragraphs(idxDebutCorps).Range.Start, _
        End:=doc.Paragraphs(idxFinCorps).Range.End)

    Set premierParagraphePatient = doc.Paragraphs(idxDebutCorps).Range

    LocaliserCorpsCourrier = True

End Function

Public Function ExtrairePatientDepuisParagraphe(ByVal texte As String) As Boolean

    Dim s As String
    Dim sMin As String
    Dim posMonsieur As Long
    Dim posMadame As Long
    Dim posCivilite As Long
    Dim posDebutIdentite As Long
    Dim posFinIdentite As Long
    Dim blocIdentite As String
    Dim mots() As String
    Dim i As Long
    Dim resteNom As String

    ExtrairePatientDepuisParagraphe = False

    s = NettoyerTexteParagraphe(texte)
    sMin = LCase$(s)

    gPatient.civilite = ""
    gPatient.nom = ""
    gPatient.prenom = ""
    gPatient.Age = ""
    gPatient.NomComplet = ""
    gPatient.marqueur = MARQUEUR_PATIENT

    posMonsieur = InStr(1, sMin, "monsieur", vbTextCompare)
    posMadame = InStr(1, sMin, "madame", vbTextCompare)

    If posMonsieur = 0 And posMadame = 0 Then Exit Function

    If posMonsieur > 0 And (posMadame = 0 Or posMonsieur < posMadame) Then
        posCivilite = posMonsieur
        gPatient.civilite = "Monsieur"
    Else
        posCivilite = posMadame
        gPatient.civilite = "Madame"
    End If

    posDebutIdentite = posCivilite + Len(gPatient.civilite)

    Do While posDebutIdentite <= Len(s)
        If Mid$(s, posDebutIdentite, 1) <> " " _
        And Mid$(s, posDebutIdentite, 1) <> Chr(160) Then Exit Do

        posDebutIdentite = posDebutIdentite + 1
    Loop

    If posDebutIdentite > Len(s) Then Exit Function

    posFinIdentite = TrouverFinIdentitePatientSelonCasse(s, posDebutIdentite)

    If posFinIdentite <= posDebutIdentite Then Exit Function

    blocIdentite = Mid$(s, posDebutIdentite, posFinIdentite - posDebutIdentite)
    blocIdentite = NettoyerTexteParagraphe(blocIdentite)

    If blocIdentite = "" Then Exit Function

    mots = Split(blocIdentite, " ")

    If UBound(mots) = 0 Then

        gPatient.nom = blocIdentite
        gPatient.prenom = ""

    ElseIf EstPrenomProbable(mots(0)) Then

        gPatient.prenom = mots(0)

        resteNom = ""
        For i = 1 To UBound(mots)
            If mots(i) <> "" Then
                If resteNom <> "" Then resteNom = resteNom & " "
                resteNom = resteNom & mots(i)
            End If
        Next i

        gPatient.nom = resteNom

    Else

        gPatient.nom = blocIdentite
        gPatient.prenom = ""

    End If

    gPatient.NomComplet = gPatient.civilite & " " & blocIdentite

    ExtrairePatientDepuisParagraphe = True

End Function

Private Function TrouverFinIdentitePatientSelonCasse(ByVal s As String, _
                                                     ByVal posDebut As Long) As Long

    Dim i As Long
    Dim debutMot As Long
    Dim mot As String
    Dim c As String

    TrouverFinIdentitePatientSelonCasse = 0

    i = posDebut

    Do While i <= Len(s)

        c = Mid$(s, i, 1)

        If c = "," Or c = ";" Or c = ":" Or c = "." Then
            TrouverFinIdentitePatientSelonCasse = i
            Exit Function
        End If

        Do While i <= Len(s)
            c = Mid$(s, i, 1)

            If c <> " " And c <> Chr(160) Then Exit Do

            i = i + 1
        Loop

        If i > Len(s) Then Exit Do

        debutMot = i

        Do While i <= Len(s)
            c = Mid$(s, i, 1)

            If EstSeparateurMot(c) Then Exit Do

            i = i + 1
        Loop

        mot = Mid$(s, debutMot, i - debutMot)
        If i = debutMot Then i = i + 1
        mot = NettoyerMot(mot)

        If mot <> "" Then

            If EstMotNumerique(mot) Then
                TrouverFinIdentitePatientSelonCasse = debutMot - 1
                Exit Function
            End If

            If EstMotEntierementMinuscule(mot) Then
                TrouverFinIdentitePatientSelonCasse = debutMot - 1
                Exit Function
            End If

        End If

    Loop

End Function

Private Function EstPrenomProbable(ByVal mot As String) As Boolean

    Dim parties() As String
    Dim p As Variant
    Dim morceau As String

    mot = NettoyerMot(mot)

    EstPrenomProbable = False

    If mot = "" Then Exit Function

    parties = Split(mot, "-")

    For Each p In parties

        morceau = CStr(p)

        If Len(morceau) < 2 Then Exit Function

        If Left$(morceau, 1) <> UCase$(Left$(morceau, 1)) Then Exit Function

        If Mid$(morceau, 2) <> LCase$(Mid$(morceau, 2)) Then Exit Function

        If morceau = UCase$(morceau) Then Exit Function

    Next p

    EstPrenomProbable = True

End Function

Private Function EstMotEntierementMinuscule(ByVal mot As String) As Boolean

    mot = NettoyerMot(mot)

    EstMotEntierementMinuscule = False

    If mot = "" Then Exit Function
    If Not contientLettre(mot) Then Exit Function

    If mot = LCase$(mot) And mot <> UCase$(mot) Then
        EstMotEntierementMinuscule = True
    End If

End Function

Private Function EstMotNumerique(ByVal mot As String) As Boolean

    mot = NettoyerMot(mot)

    If mot = "" Then
        EstMotNumerique = False
    Else
        EstMotNumerique = IsNumeric(mot)
    End If

End Function

Private Function contientLettre(ByVal mot As String) As Boolean

    Dim i As Long
    Dim c As String

    contientLettre = False

    For i = 1 To Len(mot)
        c = Mid$(mot, i, 1)

        If LCase$(c) <> UCase$(c) Then
            contientLettre = True
            Exit Function
        End If
    Next i

End Function

Private Function EstSeparateurMot(ByVal c As String) As Boolean

    EstSeparateurMot = False

    If c = " " Then EstSeparateurMot = True: Exit Function
    If c = Chr(160) Then EstSeparateurMot = True: Exit Function
    If c = "," Then EstSeparateurMot = True: Exit Function
    If c = ";" Then EstSeparateurMot = True: Exit Function
    If c = ":" Then EstSeparateurMot = True: Exit Function
    If c = "." Then EstSeparateurMot = True: Exit Function
    If c = "(" Then EstSeparateurMot = True: Exit Function
    If c = ")" Then EstSeparateurMot = True: Exit Function
    If c = vbCr Then EstSeparateurMot = True: Exit Function
    If c = vbLf Then EstSeparateurMot = True: Exit Function

End Function

Private Function NettoyerMot(ByVal mot As String) As String

    mot = Replace(mot, ",", "")
    mot = Replace(mot, ";", "")
    mot = Replace(mot, ":", "")
    mot = Replace(mot, ".", "")
    mot = Replace(mot, "(", "")
    mot = Replace(mot, ")", "")
    mot = Replace(mot, Chr(160), " ")

    NettoyerMot = Trim$(mot)

End Function

Public Sub AnonymiserPatientDansRange(ByVal rng As Range)

    Dim texte As String
    Dim texteAnonymise As String

    texte = rng.Text

    If gPatient.NomComplet = "" Then Exit Sub

    'HF1 CONFIDENTIALITE : anonymiser toutes les occurrences du nom complet.
    'Le parametre Count=-1 est imperatif avant tout appel API.
    texteAnonymise = Replace(texte, gPatient.NomComplet, _
                             gPatient.civilite & " " & MARQUEUR_PATIENT, _
                             1, -1, vbTextCompare)

    rng.Text = texteAnonymise

End Sub

Public Sub TesterAnonymisationToutesOccurrencesHF1()

    Dim civiliteAvant As String
    Dim nomAvant As String
    Dim prenomAvant As String
    Dim ageAvant As String
    Dim nomCompletAvant As String
    Dim marqueurAvant As String
    Dim docTest As Document
    Dim texteObtenu As String

    civiliteAvant = gPatient.civilite
    nomAvant = gPatient.nom
    prenomAvant = gPatient.prenom
    ageAvant = gPatient.Age
    nomCompletAvant = gPatient.NomComplet
    marqueurAvant = gPatient.marqueur

    On Error GoTo GestionErreur

    gPatient.civilite = "Madame"
    gPatient.nom = "DUPONT"
    gPatient.prenom = "Jeanne"
    gPatient.Age = "70"
    gPatient.NomComplet = "Madame DUPONT Jeanne"
    gPatient.marqueur = MARQUEUR_PATIENT

    Set docTest = Documents.Add(Visible:=False)
    docTest.Content.Text = _
        "Je revois Madame DUPONT Jeanne." & vbCr & _
        "Madame DUPONT Jeanne reste asymptomatique."

    AnonymiserPatientDansRange docTest.Content
    texteObtenu = docTest.Content.Text

    If InStr(1, texteObtenu, "DUPONT", vbTextCompare) = 0 _
    And UBound(Split(texteObtenu, MARQUEUR_PATIENT)) = 2 Then

        MsgBox _
            "OK : les deux occurrences du nom complet sont anonymisees.", _
            vbInformation, _
            "Test confidentialite HF1"
    Else
        MsgBox _
            "ECHEC : le texte anonymise obtenu est :" & vbCrLf & vbCrLf & _
            texteObtenu, _
            vbExclamation, _
            "Test confidentialite HF1"
    End If

Sortie:
    On Error Resume Next
    If Not docTest Is Nothing Then docTest.Close SaveChanges:=wdDoNotSaveChanges
    gPatient.civilite = civiliteAvant
    gPatient.nom = nomAvant
    gPatient.prenom = prenomAvant
    gPatient.Age = ageAvant
    gPatient.NomComplet = nomCompletAvant
    gPatient.marqueur = marqueurAvant
    On Error GoTo 0
    Exit Sub

GestionErreur:
    MsgBox _
        "Erreur pendant le test d'anonymisation :" & vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Test confidentialite HF1"
    Resume Sortie

End Sub

Public Function RestaurerPatientDansTexte(ByVal texte As String) As String

    RestaurerPatientDansTexte = Replace( _
        texte, _
        gPatient.civilite & " " & MARQUEUR_PATIENT, _
        gPatient.NomComplet, _
        1, -1, vbTextCompare)

    RestaurerPatientDansTexte = Replace( _
        RestaurerPatientDansTexte, _
        MARQUEUR_PATIENT, _
        Replace(gPatient.NomComplet, gPatient.civilite & " ", ""), _
        1, -1, vbTextCompare)

End Function

Public Function EstParagrapheDateCourrier(ByVal s As String) As Boolean

    Dim t As String

    t = LCase(NettoyerTexteParagraphe(s))

    EstParagrapheDateCourrier = False

    If t = "" Then Exit Function

    'Évite de confondre avec l'adresse : 95260 Beaumont sur Oise.
    If InStr(1, t, "beaumont", vbTextCompare) = 0 Then Exit Function
    If InStr(1, t, " le ", vbTextCompare) = 0 Then Exit Function

    'On exige au moins une année probable à 4 chiffres.
    If ContientAnneeQuatreChiffres(t) Then
        EstParagrapheDateCourrier = True
    End If

End Function

Private Function ContientAnneeQuatreChiffres(ByVal s As String) As Boolean

    Dim i As Long
    Dim fragment As String

    ContientAnneeQuatreChiffres = False

    For i = 1 To Len(s) - 3

        fragment = Mid$(s, i, 4)

        If IsNumeric(fragment) Then
            If CLng(fragment) >= 1990 And CLng(fragment) <= 2099 Then
                ContientAnneeQuatreChiffres = True
                Exit Function
            End If
        End If

    Next i

End Function
