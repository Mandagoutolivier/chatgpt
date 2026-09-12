Attribute VB_Name = "modDemandes"
Option Explicit
' =====================================================================
' modDemandes - Reperage des demandes d'examen ou d'avis dans le courrier
' principal, et profils (canevas) de redaction des lettres de demande.
'
' 1. DETECTION (Config\demandes\, fichiers texte modifiables par le
'    medecin au Bloc-notes) :
'      declencheurs.txt : formules qui expriment une demande ("je prescris*")
'      examens.txt      : expression|CODE_DU_PROFIL
'      exclusions.txt   : formules qui annulent ("avait realise*")
'    Une demande est reconnue quand, dans la MEME phrase, on trouve un
'    declencheur ET un examen, sans exclusion. Comparaison sans casse ni
'    accents ; * = texte variable ; l'expression d'examen la plus longue
'    l'emporte (irm de stress avant irm, coroscanner avant scanner).
'
' 2. PROFILS (Config\profils\CODE.ini, issus des canevas R12) : premier
'    paragraphe, rubriques dans l'ordre, obligatoires / si presentes /
'    optionnelles / interdites, prefixes de phrase, longueur, objectif,
'    modalites, TYPE_BASE (codes du classeur des specialistes).
'
' 3. GENERATION AUTOMATIQUE a la validation du courrier principal
'    (modValidation) : une lettre par profil detecte, destinataire de
'    priorite 1 du classeur des specialistes, sans intervention.
' =====================================================================

' ---------------------------------------------------------------------
' Chemins
' ---------------------------------------------------------------------
Public Function DossierDemandes() As String
    DossierDemandes = modConfig.racine() & "\" & modConfig.Config("DERIVEES", "DossierDemandes", "Config\demandes")
End Function

Public Function DossierProfils() As String
    DossierProfils = modConfig.racine() & "\" & modConfig.Config("DERIVEES", "DossierProfils", "Config\profils")
End Function

' ---------------------------------------------------------------------
' Normalisation : minuscules, sans accents, apostrophes droites, espaces simples
' ---------------------------------------------------------------------
Public Function Normaliser(ByVal t As String) As String
    t = modTexte.Plier(t)
    t = Replace(t, ChrW$(8217), "'")
    t = Replace(t, ChrW$(8216), "'")
    t = Replace(t, ChrW$(160), " ")
    t = Replace(t, vbTab, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    Normaliser = Trim$(t)
End Function

' Lignes utiles d'un fichier de dictionnaire (sans commentaires ni vides)
Private Function LignesDictionnaire(ByVal nomFichier As String) As Collection
    Dim res As Collection, chemin As String, contenu As String, lignes() As String, i As Long, l As String
    Set res = New Collection
    chemin = DossierDemandes() & "\" & nomFichier
    If Not modFichiers.FichierExiste(chemin) Then
        modLog.LogErreur "Dictionnaire de demandes absent : " & chemin
        Set LignesDictionnaire = res
        Exit Function
    End If
    contenu = modFichiers.LireTexteUTF8(chemin)
    contenu = Replace(contenu, vbCrLf, vbLf)
    contenu = Replace(contenu, vbCr, vbLf)
    lignes = Split(contenu, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        l = Trim$(lignes(i))
        If Len(l) > 0 And Left$(l, 1) <> "#" And Left$(l, 1) <> ";" Then res.Add l
    Next i
    Set LignesDictionnaire = res
End Function

' Motif avec * : chaque morceau doit apparaitre dans l'ordre
Public Function MotifCorrespond(ByVal texteNormalise As String, ByVal motif As String) As Boolean
    Dim morceaux() As String, i As Long, position As Long, trouve As Long, m As String
    motif = Normaliser(motif)
    If Len(motif) = 0 Then Exit Function
    If InStr(motif, "*") = 0 Then
        MotifCorrespond = (InStr(1, texteNormalise, motif, vbTextCompare) > 0)
        Exit Function
    End If
    morceaux = Split(motif, "*")
    position = 1
    For i = LBound(morceaux) To UBound(morceaux)
        m = Trim$(morceaux(i))
        If Len(m) > 0 Then
            trouve = InStr(position, texteNormalise, m, vbTextCompare)
            If trouve = 0 Then Exit Function
            position = trouve + Len(m)
        End If
    Next i
    MotifCorrespond = True
End Function

Private Function ContientMotif(ByVal texteNormalise As String, ByVal nomFichier As String) As Boolean
    Dim l As Variant
    For Each l In LignesDictionnaire(nomFichier)
        If MotifCorrespond(texteNormalise, CStr(l)) Then ContientMotif = True: Exit Function
    Next l
End Function

' Decoupage en phrases (. ! ? ; et fins de paragraphe)
Public Function DecouperPhrases(ByVal texte As String) As Collection
    Dim c As Collection, t As String, morceaux() As String, i As Long
    Set c = New Collection
    t = Replace(texte, vbCrLf, ".")
    t = Replace(t, vbCr, ".")
    t = Replace(t, vbLf, ".")
    t = Replace(t, "!", ".")
    t = Replace(t, "?", ".")
    t = Replace(t, ";", ".")
    morceaux = Split(t, ".")
    For i = LBound(morceaux) To UBound(morceaux)
        If Len(Trim$(morceaux(i))) > 0 Then c.Add Trim$(morceaux(i))
    Next i
    Set DecouperPhrases = c
End Function

Private Function EstLettre(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    Select Case AscW(c)
        Case 48 To 57, 97 To 122: EstLettre = True
        Case Else: EstLettre = (c = "-")
    End Select
End Function

' Position du premier mot entier "expr" dans "texte" (tous deux normalises), 0 si absent
Private Function PositionMotEntier(ByVal texte As String, ByVal expr As String, Optional ByVal depart As Long = 1) As Long
    Dim pos As Long, avant As String, apres As String
    pos = depart
    Do
        pos = InStr(pos, texte, expr, vbBinaryCompare)
        If pos = 0 Then Exit Function
        If pos > 1 Then avant = Mid$(texte, pos - 1, 1) Else avant = ""
        If pos + Len(expr) <= Len(texte) Then apres = Mid$(texte, pos + Len(expr), 1) Else apres = ""
        If Not EstLettre(avant) And Not EstLettre(apres) Then
            PositionMotEntier = pos
            Exit Function
        End If
        pos = pos + 1
    Loop
End Function

' Examens presents dans une phrase normalisee : Collection de dictionnaires
' (Code, Expression). Les expressions sont testees de la plus longue a la
' plus courte et une expression incluse dans une zone deja reconnue est ignoree.
Public Function ExamensDansPhrase(ByVal phraseNormalisee As String) As Collection
    Dim res As Collection, lignes As Collection, exprs() As String, codes() As String
    Dim n As Long, i As Long, j As Long, l As Variant, p As Long, tmpE As String, tmpC As String
    Dim zones As String, pos As Long, libre As Boolean, k As Long, d As Object
    Set res = New Collection
    Set lignes = LignesDictionnaire("examens.txt")
    n = lignes.Count
    If n = 0 Then Set ExamensDansPhrase = res: Exit Function
    ReDim exprs(0 To n - 1): ReDim codes(0 To n - 1)
    i = 0
    For Each l In lignes
        p = InStr(CStr(l), "|")
        If p > 0 Then
            exprs(i) = Normaliser(Left$(CStr(l), p - 1))
            codes(i) = UCase$(Trim$(Mid$(CStr(l), p + 1)))
        Else
            exprs(i) = Normaliser(CStr(l)): codes(i) = "AUTRE_EXAMEN"
        End If
        i = i + 1
    Next l
    ' tri par longueur decroissante (insertion)
    For i = 1 To n - 1
        tmpE = exprs(i): tmpC = codes(i): j = i - 1
        Do While j >= 0
            If Len(exprs(j)) >= Len(tmpE) Then Exit Do
            exprs(j + 1) = exprs(j): codes(j + 1) = codes(j): j = j - 1
        Loop
        exprs(j + 1) = tmpE: codes(j + 1) = tmpC
    Next i
    ' zones : chaine de meme longueur que la phrase, "x" = deja reconnu
    zones = String$(Len(phraseNormalisee), " ")
    For i = 0 To n - 1
        If Len(exprs(i)) > 0 Then
            pos = PositionMotEntier(phraseNormalisee, exprs(i))
            Do While pos > 0
                libre = (InStr(Mid$(zones, pos, Len(exprs(i))), "x") = 0)
                If libre Then
                    Mid$(zones, pos, Len(exprs(i))) = String$(Len(exprs(i)), "x")
                    Set d = CreateObject("Scripting.Dictionary")
                    d.CompareMode = 1
                    d("Code") = codes(i)
                    d("Expression") = exprs(i)
                    d("Position") = pos
                    res.Add d
                End If
                pos = PositionMotEntier(phraseNormalisee, exprs(i), pos + Len(exprs(i)))
            Loop
        End If
    Next i
    Set ExamensDansPhrase = res
End Function

' Demandes reconnues dans un corps de courrier : Collection de dictionnaires
' (Code, Expression, Phrase) ; un code n'apparait qu'une fois, avec la
' premiere phrase qui le declenche (les autres phrases sont ajoutees dans
' "Phrases" pour l'API).
Public Function DetecterDemandes(ByVal corps As String) As Collection
    Dim res As Collection, index As Object, ph As Variant, pn As String, ex As Variant, d As Object
    Set res = New Collection
    Set index = CreateObject("Scripting.Dictionary")
    index.CompareMode = 1
    For Each ph In DecouperPhrases(corps)
        pn = Normaliser(CStr(ph))
        If Len(pn) > 0 Then
            If ContientMotif(pn, "declencheurs.txt") And Not ContientMotif(pn, "exclusions.txt") Then
                For Each ex In ExamensDansPhrase(pn)
                    If index.Exists(ex("Code")) Then
                        Set d = index(ex("Code"))
                        d("Phrases") = d("Phrases") & vbLf & Trim$(CStr(ph))
                    Else
                        Set d = CreateObject("Scripting.Dictionary")
                        d.CompareMode = 1
                        d("Code") = ex("Code")
                        d("Expression") = ex("Expression")
                        d("Phrase") = Trim$(CStr(ph))
                        d("Phrases") = Trim$(CStr(ph))
                        Set index(ex("Code")) = d
                        res.Add d
                    End If
                Next ex
            End If
        End If
    Next ph
    Set DetecterDemandes = res
End Function

' ---------------------------------------------------------------------
' Profils (.ini) : dictionnaire "SECTION|CLE" -> valeur, + "Code"
' ---------------------------------------------------------------------
Public Function ProfilExiste(ByVal code As String) As Boolean
    ProfilExiste = modFichiers.FichierExiste(DossierProfils() & "\" & UCase$(code) & ".ini")
End Function

Public Function ChargerProfil(ByVal code As String) As Object
    Dim p As Object, chemin As String, contenu As String, lignes() As String
    Dim i As Long, l As String, section As String, eq As Long
    code = UCase$(Trim$(code))
    chemin = DossierProfils() & "\" & code & ".ini"
    If Not modFichiers.FichierExiste(chemin) Then
        If code <> "AUTRE_EXAMEN" And modFichiers.FichierExiste(DossierProfils() & "\AUTRE_EXAMEN.ini") Then
            modLog.LogInfo "Profil " & code & " absent : profil AUTRE_EXAMEN utilise"
            Dim repli As Object
            Set repli = ChargerProfil("AUTRE_EXAMEN")
            repli("Code") = code
            Set ChargerProfil = repli
            Exit Function
        End If
        Err.Raise vbObjectError + 510, "modDemandes", "Profil de demande introuvable : " & chemin
    End If
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = 1
    p("Code") = code
    contenu = modFichiers.LireTexteUTF8(chemin)
    contenu = Replace(contenu, vbCrLf, vbLf)
    contenu = Replace(contenu, vbCr, vbLf)
    lignes = Split(contenu, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        l = Trim$(lignes(i))
        If Len(l) > 0 And Left$(l, 1) <> ";" And Left$(l, 1) <> "#" Then
            If Left$(l, 1) = "[" And Right$(l, 1) = "]" Then
                section = UCase$(Mid$(l, 2, Len(l) - 2))
            Else
                eq = InStr(l, "=")
                If eq > 1 And Len(section) > 0 Then
                    If section = "MODALITES" Then
                        ' liste ordonnee "libelle=texte|..." conservee telle quelle
                        If Len(p("MODALITES|LISTE")) > 0 Then p("MODALITES|LISTE") = p("MODALITES|LISTE") & "|"
                        p("MODALITES|LISTE") = p("MODALITES|LISTE") & l
                    Else
                        p(section & "|" & UCase$(Trim$(Left$(l, eq - 1)))) = Trim$(Mid$(l, eq + 1))
                    End If
                End If
            End If
        End If
    Next i
    If Len(p("IDENTITE|LIBELLE")) = 0 Then p("IDENTITE|LIBELLE") = LCase$(Replace(code, "_", " "))
    If Len(p("IDENTITE|ARTICLE")) = 0 Then p("IDENTITE|ARTICLE") = "un"
    If Len(p("IDENTITE|TYPE_BASE")) = 0 Then p("IDENTITE|TYPE_BASE") = code
    Set ChargerProfil = p
End Function

Public Function ValeurProfil(ByVal p As Object, ByVal section As String, ByVal cle As String, _
                             Optional ByVal defaut As String = "") As String
    Dim k As String
    k = UCase$(section) & "|" & UCase$(cle)
    If p.Exists(k) Then ValeurProfil = p(k) Else ValeurProfil = defaut
End Function

' Tous les profils du dossier, tries par libelle : Collection de dictionnaires
' (ID=Code, Code, Libelle, Nature)
Public Function ListerProfils(Optional ByVal filtreCodes As String = "") As Collection
    Dim res As Collection, f As Variant, code As String, p As Object, d As Object, i As Long, inserer As Boolean
    Set res = New Collection
    For Each f In modFichiers.ListerFichiers(DossierProfils(), ".ini")
        code = UCase$(Mid$(CStr(f), InStrRev(CStr(f), "\") + 1))
        code = Left$(code, Len(code) - 4)
        If Len(filtreCodes) = 0 Or InStr(";" & UCase$(filtreCodes) & ";", ";" & code & ";") > 0 Then
            On Error Resume Next
            Set p = ChargerProfil(code)
            On Error GoTo 0
            If Not p Is Nothing Then
                Set d = CreateObject("Scripting.Dictionary")
                d.CompareMode = 1
                d("ID") = code: d("Code") = code
                d("Libelle") = ValeurProfil(p, "IDENTITE", "LIBELLE")
                d("Nature") = ValeurProfil(p, "IDENTITE", "NATURE")
                inserer = False
                For i = 1 To res.Count
                    If StrComp(res(i)("Libelle"), d("Libelle"), vbTextCompare) > 0 Then
                        res.Add d, , i: inserer = True: Exit For
                    End If
                Next i
                If Not inserer Then res.Add d
                Set p = Nothing
            End If
        End If
    Next f
    Set ListerProfils = res
End Function

' Modalites d'un profil : Collection de dictionnaires (ID, Libelle, Texte)
Public Function ModalitesProfil(ByVal p As Object) As Collection
    Dim res As Collection, liste As String, items() As String, i As Long, eq As Long, d As Object
    Set res = New Collection
    liste = ValeurProfil(p, "MODALITES", "LISTE")
    If Len(liste) = 0 Then Set ModalitesProfil = res: Exit Function
    items = Split(liste, "|")
    For i = LBound(items) To UBound(items)
        eq = InStr(items(i), "=")
        If eq > 1 Then
            Set d = CreateObject("Scripting.Dictionary")
            d.CompareMode = 1
            d("ID") = CStr(i)
            d("Libelle") = Trim$(Left$(items(i), eq - 1))
            d("Texte") = Trim$(Mid$(items(i), eq + 1))
            res.Add d
        End If
    Next i
    Set ModalitesProfil = res
End Function

' Texte de l'examen pour la premiere phrase : modalite choisie, sinon
' "article libelle" du profil
Public Function TexteExamen(ByVal p As Object, Optional ByVal modalite As String = "") As String
    Dim m As Variant
    If Len(modalite) > 0 Then
        For Each m In ModalitesProfil(p)
            If StrComp(m("Libelle"), modalite, vbTextCompare) = 0 Then
                TexteExamen = m("Texte")
                Exit Function
            End If
        Next m
    End If
    TexteExamen = ValeurProfil(p, "IDENTITE", "ARTICLE") & " " & ValeurProfil(p, "IDENTITE", "LIBELLE")
End Function

Public Function LibelleProfil(ByVal p As Object, Optional ByVal modalite As String = "") As String
    LibelleProfil = ValeurProfil(p, "IDENTITE", "LIBELLE")
    If Len(modalite) > 0 Then LibelleProfil = LibelleProfil & " " & modalite
End Function

' ---------------------------------------------------------------------
' Consignes de redaction pour l'API, construites depuis le profil
' ---------------------------------------------------------------------
Public Function ConsignesProfil(ByVal p As Object, ByVal modalite As String, ByVal tutoiement As Boolean, _
                                ByVal identitePatient As String, ByVal motif As String, _
                                ByVal phrasesPrescription As String) As String
    Dim s As String, premier As String, ordre() As String, i As Long, r As String
    Dim oblig As String, siPresent As String, optionnel As String, interdit As String
    Dim prefixe As String, selection As String, statut As String, longueur As String, objectif As String
    Dim tv As String, txtMod As String, examen As String, variantes As String, exemples As String

    tv = IIf(tutoiement, "te", "vous")
    examen = ValeurProfil(p, "IDENTITE", "ARTICLE") & " " & ValeurProfil(p, "IDENTITE", "LIBELLE")
    txtMod = TexteModalite(p, modalite)
    ' 1. premier paragraphe impose (modele du profil, ou modalite qui le remplace)
    premier = ValeurProfil(p, "PREMIER_PARAGRAPHE", "MODELE", "Merci de réaliser {ARTICLE_EXAMEN} à {PATIENT}{AGE_SEGMENT}, {MOTIF}.")
    If Len(txtMod) > 0 And LCase$(Left$(Trim$(txtMod), 5)) = "merci" Then
        premier = txtMod                                   ' la modalite est une premiere phrase complete
    ElseIf InStr(1, premier, "{MODALITE}", vbTextCompare) > 0 Or InStr(1, premier, "{DEGRE_URGENCE}", vbTextCompare) > 0 Then
        premier = Replace(premier, "{MODALITE}", IIf(Len(txtMod) > 0, " " & txtMod, ""), , , vbTextCompare)
        premier = Replace(premier, "{DEGRE_URGENCE}", IIf(Len(txtMod) > 0, " " & txtMod, ""), , , vbTextCompare)
    ElseIf Len(txtMod) > 0 Then
        examen = txtMod                                    ' la modalite remplace l'examen
    End If
    premier = Replace(premier, "{ARTICLE_EXAMEN}", examen, , , vbTextCompare)
    premier = Replace(premier, "{EXAMEN}", examen, , , vbTextCompare)
    premier = Replace(premier, "{PATIENT}{AGE_SEGMENT}", identitePatient, , , vbTextCompare)
    premier = Replace(premier, "{PATIENT}", identitePatient, , , vbTextCompare)
    premier = Replace(premier, "{AGE_SEGMENT}", "", , , vbTextCompare)
    If Len(Trim$(motif)) > 0 Then
        premier = Replace(premier, "{MOTIF}", Trim$(motif), , , vbTextCompare)
    ElseIf Len(ValeurProfil(p, "MOTIF", "DEFAUT")) > 0 Then
        premier = Replace(premier, "{MOTIF}", ValeurProfil(p, "MOTIF", "DEFAUT"), , , vbTextCompare)
    End If
    ' sans motif saisi ni defaut, {MOTIF} reste dans la phrase : l'API le
    ' remplit d'apres la phrase de prescription (jamais un defaut clinique)
    s = "Première phrase, à reprendre telle quelle : " & vbLf & premier & vbLf
    s = s & "Dans cette phrase, le motif ou le terrain se rattache directement au patient, en apposition (« hypertendu, dyslipidémique ») ou par une relative (« qui présente ... », « aux antécédents de ... »), d'après la phrase de prescription et le courrier ; les champs restants entre accolades sont à remplir ainsi : {TERRAIN} = terrain ou antécédent majeur en apposition suivi d'une virgule, vide sinon ; {PRONOM_OBJET} = le / la / l' selon le patient ; {SPECIALITE} = la spécialité demandée ; {VALEUR} = la valeur citée dans le courrier ; {MOTIF} = le terrain, l'antécédent ou le motif tel que la phrase de prescription et le courrier le donnent, sinon rien du tout (supprimer alors la virgule qui précède). Ne jamais laisser d'accolade dans la lettre, ne jamais inventer un motif." & vbLf
    exemples = ValeurProfil(p, "MOTIF", "EXEMPLES")
    If Len(exemples) = 0 Then exemples = ValeurProfil(p, "MOTIF", "FORMES")
    If Len(exemples) > 0 Then s = s & "Formes réelles du motif chez ce médecin : " & Replace(exemples, ";", " / ") & vbLf
    If Len(ValeurProfil(p, "INSTRUCTIONS", "MOTIF")) > 0 Then s = s & "Consigne pour le motif : " & ValeurProfil(p, "INSTRUCTIONS", "MOTIF") & vbLf
    s = s & vbLf
    If Len(phrasesPrescription) > 0 Then
        s = s & "Phrase(s) du courrier de consultation qui formulent cette demande (elles donnent l'examen précis, sa modalité et le motif) :" & vbLf & _
                phrasesPrescription & vbLf & vbLf
    End If

    ' 2. rubriques du corps
    longueur = UCase$(ValeurProfil(p, "STYLE", "LONGUEUR", "MOYENNE"))
    Select Case longueur
        Case "COURTE": s = s & "Longueur : COURTE, une à trois lignes en tout, première phrase comprise ; une seule phrase suffit souvent." & vbLf
        Case "DETAILLEE": s = s & "Longueur : DÉTAILLÉE, jusqu'à " & ValeurProfil(p, "STYLE", "PARAGRAPHES_MAX", "10") & " lignes, une rubrique par ligne." & vbLf
        Case Else: s = s & "Longueur : MOYENNE, quatre à six lignes en tout, jamais plus de " & ValeurProfil(p, "STYLE", "PARAGRAPHES_MAX", "7") & "." & vbLf
    End Select
    oblig = ";" & UCase$(ValeurProfil(p, "STRUCTURE", "OBLIGATOIRE")) & ";"
    siPresent = ";" & UCase$(ValeurProfil(p, "STRUCTURE", "OBLIGATOIRE_SI_PRESENT")) & ";"
    optionnel = ";" & UCase$(ValeurProfil(p, "STRUCTURE", "OPTIONNEL")) & ";"
    interdit = ValeurProfil(p, "STRUCTURE", "INTERDIT")
    s = s & "Rubriques, dans cet ordre, une ligne chacune, en troisième personne (jamais d'adresse directe au confrère dans ces lignes) :" & vbLf
    ordre = Split(ValeurProfil(p, "STRUCTURE", "ORDRE"), ";")
    For i = LBound(ordre) To UBound(ordre)
        r = UCase$(Trim$(ordre(i)))
        If Len(r) > 0 And r <> "OBJECTIF" Then
            If InStr(oblig, ";" & r & ";") > 0 Then
                statut = "obligatoire"
            ElseIf InStr(siPresent, ";" & r & ";") > 0 Then
                statut = "obligatoire si le courrier source contient l'information, sinon omise"
            ElseIf InStr(optionnel, ";" & r & ";") > 0 Then
                statut = "optionnelle, seulement si utile au confrère"
            Else
                statut = "si présente dans le courrier source"
            End If
            prefixe = ValeurProfil(p, "PREFIXES", r)
            selection = ValeurProfil(p, "SELECTION", r)
            s = s & "- " & LibelleRubrique(r) & " : " & statut
            If Len(prefixe) > 0 Then s = s & " ; commencer par « " & Replace(prefixe, "|", " » ou « ") & " »"
            If Len(selection) > 0 Then s = s & " ; éléments à chercher : " & Replace(selection, ";", ", ")
            If Len(ValeurProfil(p, "INSTRUCTIONS", r)) > 0 Then s = s & " ; consigne : " & ValeurProfil(p, "INSTRUCTIONS", r)
            s = s & "." & vbLf
        End If
    Next i
    If Len(interdit) > 0 Then s = s & "Rubriques INTERDITES (ne pas les rédiger) : " & Replace(interdit, ";", ", ") & "." & vbLf

    ' 3. derniere phrase : l'objectif du profil (« Merci de confirmer ... ») tient
    '    lieu de phrase finale ; a defaut une phrase de courtoisie selon la nature
    objectif = ValeurProfil(p, "OBJECTIF", "MODELE")
    objectif = Replace(objectif, "{DEGRE_URGENCE}", IIf(Len(txtMod) > 0, " " & txtMod, ""), , , vbTextCompare)
    objectif = Replace(objectif, "{MODALITE}", IIf(Len(txtMod) > 0, " " & txtMod, ""), , , vbTextCompare)
    variantes = ValeurProfil(p, "OBJECTIF", "VARIANTES")
    If Len(objectif) > 0 Then
        s = s & "Dernière phrase (la demande précise, à reprendre ou adapter sobrement au contexte) : " & objectif
        If Len(variantes) > 0 Then s = s & " ; variantes selon le contexte : " & Replace(variantes, ";", " / ")
        If Len(ValeurProfil(p, "INSTRUCTIONS", "OBJECTIF")) > 0 Then s = s & " ; consigne : " & ValeurProfil(p, "INSTRUCTIONS", "OBJECTIF")
        s = s & vbLf & "Aucune autre phrase de courtoisie (pas de « Je te serais reconnaissant », pas de « et de m'adresser le compte rendu ») : la formule de politesse existe déjà dans le document."
    Else
        Select Case UCase$(ValeurProfil(p, "IDENTITE", "NATURE", "EXAMEN"))
            Case "HOSPITALISATION"
                s = s & "Dernière phrase : « Merci de le prendre en charge » (ou « la »)."
            Case "CONSULTATION"
                s = s & "Dernière phrase : « Merci de l'évaluer » ou la question précise posée au confrère."
            Case Else
                s = s & "Dernière phrase : « Je " & tv & " serais reconnaissant de bien vouloir réaliser cet examen »."
        End Select
    End If
    ConsignesProfil = s
End Function

' Texte de la modalite choisie ("" si aucune)
Private Function TexteModalite(ByVal p As Object, ByVal modalite As String) As String
    Dim m As Variant
    If Len(modalite) = 0 Then Exit Function
    For Each m In ModalitesProfil(p)
        If StrComp(m("Libelle"), modalite, vbTextCompare) = 0 Then
            TexteModalite = m("Texte")
            Exit Function
        End If
    Next m
End Function

Private Function LibelleRubrique(ByVal r As String) As String
    Select Case r
        Case "ANAMNESE": LibelleRubrique = "Anamnèse (symptômes ou leur absence, motif clinique)"
        Case "FACTEURS_RISQUE": LibelleRubrique = "Facteurs de risque cardiovasculaire"
        Case "TRAITEMENTS": LibelleRubrique = "Traitement en cours"
        Case "EXAMEN_CLINIQUE": LibelleRubrique = "Examen clinique"
        Case "ECG": LibelleRubrique = "Électrocardiogramme"
        Case "ECHOGRAPHIE": LibelleRubrique = "Échocardiographie"
        Case "EXAMENS_ANTERIEURS", "EXAMENS_ISCHEMIQUES_ANTERIEURS", "IMAGERIE": LibelleRubrique = "Examens antérieurs"
        Case "CONSIGNE_MEDICAMENTEUSE": LibelleRubrique = "Consigne médicamenteuse avant l'examen"
        Case "CORONAROPATHIE": LibelleRubrique = "Contexte coronarien"
        Case "REVASCULARISATION": LibelleRubrique = "Revascularisation"
        Case "BIOLOGIE_RENALE": LibelleRubrique = "Fonction rénale"
        Case "BIOLOGIE": LibelleRubrique = "Biologie"
        Case "INDICATION": LibelleRubrique = "Indication"
        Case "ANTECEDENTS", "ANTECEDENTS_CARDIOLOGIQUES": LibelleRubrique = "Antécédents"
        Case "HOLTER": LibelleRubrique = "Holter"
        Case "QUESTION": LibelleRubrique = "Question posée au confrère"
        Case "MOTIF_HOSPITALISATION": LibelleRubrique = "Motif d'hospitalisation"
        Case "ANTICOAGULATION": LibelleRubrique = "Anticoagulation"
        Case "PRISE_EN_CHARGE_ATTENDUE": LibelleRubrique = "Prise en charge attendue"
        Case "SYMPTOMES_SAOS": LibelleRubrique = "Symptômes évocateurs d'apnées du sommeil"
        Case "TERRAIN": LibelleRubrique = "Terrain"
        Case "HTA_MAPA": LibelleRubrique = "Profil tensionnel (MAPA)"
        Case "RYTHMOLOGIE": LibelleRubrique = "Contexte rythmologique"
        Case Else: LibelleRubrique = Replace(LCase$(r), "_", " ")
    End Select
End Function

' ---------------------------------------------------------------------
' Generation automatique a la validation du courrier principal
' Renvoie un compte rendu texte (une ligne par lettre) pour l'information
' du medecin ; les erreurs sont journalisees et n'interrompent pas la
' validation du courrier principal.
' ---------------------------------------------------------------------
Public Function GenererDemandesAutomatiques(ByVal docSource As Document) As String
    Err.Raise vbObjectError + 988, "modDemandes", "Moteur historique desactive. Utilisez Unifie_D_Finaliser."
End Function
' Ajoute le contenu d'une lettre de demande a la fin du document principal,
' apres un saut de page. Le debut de la zone ajoutee est marque par le
' signet DEMANDES_DEBUT (pose une seule fois) pour pouvoir la remplacer.

' Retire les demandes ajoutees par une finalisation precedente

Private Function VariableDoc(ByVal doc As Document, ByVal nom As String) As String
    On Error Resume Next
    VariableDoc = doc.Variables(nom).Value
    Err.Clear
End Function
