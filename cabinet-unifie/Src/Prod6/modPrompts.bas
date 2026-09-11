Attribute VB_Name = "modPrompts"
Option Explicit

'===============================================================================
' MODULE : modPrompts
' VERSION : DOMICILE-DOUBLE-API-DETAILLE-1 / gras local
'
' Prompt simplifié à partir de 4 307 courriers récents du cabinet.
' Le prompt principal produit UNIQUEMENT le bloc CORPS_COURRIER.
' Les demandes complémentaires sont produites par un second appel API séparé.
'
' L'ancien format ---DEMANDE_EXAMEN--- n'est plus demandé par le prompt v24.
' Les fonctions publiques historiques sont conservées afin de ne pas casser
' les appels provenant des autres modules pendant la transition.
'===============================================================================
' La casse et le gras sont désormais appliqués localement par modNormalisationGrasLocal.
' L'API ne reçoit plus aucune consigne de mise en gras / Markdown.

Public Function ConstruirePromptReecritureMedicale( _
    ByVal texteAnonymise As String) As String

    Dim p As String

    p = ""

    AjouterLignePrompt p, "Tu réécris le corps d'un courrier médical de cardiologie dicté par le Docteur Olivier Mandagout."
    AjouterLignePrompt p, "Produis un français médical correct tout en restant au plus près de la formulation, du rythme, de la concision et de la logique du texte source."
    AjouterLignePrompt p, "Corrige les erreurs de dictée, d'orthographe, d'accord, de conjugaison, de syntaxe et de ponctuation, sans transformer le courrier en texte académique ou administratif."
    AjouterLignePrompt p, "Ne résume pas, ne simplifie pas excessivement, ne supprime aucune donnée médicale et n'ajoute aucune information absente du texte source."
    AjouterLignePrompt p, "Conserve exactement le marqueur [[PATIENT]] et n'invente jamais de nom, de prénom, d'âge, de diagnostic, de résultat, de traitement ou de destinataire."
    AjouterLignePrompt p, "Dans le courrier principal, le confrère destinataire peut être désigné par tu ou vous ; le patient doit rester à la troisième personne : il, elle, le, la, lui, l' ou Monsieur/Madame [[PATIENT]]."
    AjouterLignePrompt p, "Conserve le tutoiement ou le vouvoiement du texte source et son degré de proximité confraternelle."
    AjouterLignePrompt p, "N'utilise ni liste à puces, ni tableau, ni titre visible, ni commentaire avant ou après les blocs demandés."
    AjouterLignePrompt p, "Conserve les formulations caractéristiques du texte source lorsqu'elles sont médicalement et grammaticalement correctes ; n'introduis pas de tournures administratives ou emphatiques inutiles."
    AjouterLignePrompt p, "Rédige des paragraphes courts et denses, généralement composés d'une ou deux phrases, avec une seule unité médico-logique par paragraphe."
    AjouterLignePrompt p, "Ne réorganise pas un courrier déjà cohérent ; lorsque le texte est désordonné, respecte l'ordre médico-logique défini ci-dessous sans inventer de rubrique absente."
    AjouterLignePrompt p, "Le premier paragraphe conserve l'amorce naturelle du courrier, notamment Je revois Monsieur/Madame [[PATIENT]] ou Je vois pour la première fois, lorsqu'elle figure dans la source."
    AjouterLignePrompt p, "Présente les antécédents et l'évolution clinique dans un ou plusieurs paragraphes distincts selon leur longueur et leur cohérence."
    AjouterLignePrompt p, "Présente les symptômes actuels ou le caractère asymptomatique dans un paragraphe distinct lorsqu'ils sont décrits."
    AjouterLignePrompt p, "Présente le traitement actuel et ses modifications dans un paragraphe distinct lorsqu'ils sont décrits."
    AjouterLignePrompt p, "Toute donnée d'examen clinique doit former un paragraphe distinct."
    AjouterLignePrompt p, "Toute description d'électrocardiogramme doit former un paragraphe distinct et ne doit jamais être collée aux symptômes, aux antécédents ou à la conclusion."
    AjouterLignePrompt p, "Toute échographie cardiaque, échocardiographie, échoscopie, exploration Holter, MAPA, imagerie, test d'effort, scintigraphie, coroscanner ou autre examen décrit doit former son propre paragraphe lorsqu'il comporte un résultat."
    AjouterLignePrompt p, "La biologie ou le bilan biologique doit former un paragraphe distinct lorsqu'il est décrit."
    AjouterLignePrompt p, "Lorsqu'une synthèse existe dans le texte source, introduis-la de préférence par Au total : et place-la dans un paragraphe distinct."
    AjouterLignePrompt p, "Place la conduite à tenir, les modifications thérapeutiques, les examens demandés et le délai de suivi dans un ou plusieurs paragraphes distincts après la synthèse."
    AjouterLignePrompt p, "N'invente jamais un examen clinique, un ECG, une échographie, un Holter, une biologie, un antécédent ou une conclusion pour compléter artificiellement la structure."
    AjouterLignePrompt p, "Conserve strictement les artères, segments, pourcentages de sténose, dates, gestes de revascularisation, valeurs chiffrées, doses, délais et consignes présents dans le texte source."
    AjouterLignePrompt p, "Retourne le courrier principal exclusivement entre les balises ---CORPS_COURRIER--- et ---FIN_CORPS_COURRIER---."

    AjouterLignePrompt p, "Ne rédige aucune lettre complémentaire et aucun bloc DEMANDE_DESTINATION dans ce premier appel."
    AjouterLignePrompt p, ""
    AjouterLignePrompt p, "COURRIER MÉDICAL ANONYMISÉ À TRAITER :"
    AjouterLignePrompt p, "--------------------"
    p = p & texteAnonymise & vbCrLf
    AjouterLignePrompt p, "--------------------"

    ConstruirePromptReecritureMedicale = p

End Function

Public Function TexteContientDemandeExamenEligible( _
    ByVal texte As String) As Boolean

    Dim t As String

    t = NormaliserTexteDetectionDemande(texte)
    TexteContientDemandeExamenEligible = False

    'La rythmologie conserve sa securite specifique anti-faux-positif.
    If ContientDemandeRythmologiqueExplicite(t) Then
        TexteContientDemandeExamenEligible = True
        Exit Function
    End If

    'Tous les autres examens utilisent les dictionnaires NAS extensibles.
    If DDE_ContientDemandeEligible(texte) Then
        TexteContientDemandeExamenEligible = True
    End If

End Function

Public Sub TesterDetectionJeProposeEchodopplerVeineux()

    Dim texteTest As String

    texteTest = _
        "Par ailleurs, je propose la réalisation d'un échodoppler veineux " & _
        "des membres inférieurs afin d'exclure une origine veineuse à cet oedème."

    If TexteContientDemandeExamenEligible(texteTest) Then
        MsgBox _
            "OK : la formulation 'je propose la réalisation de' est reconnue " & _
            "comme une demande d'examen éligible.", _
            vbInformation, _
            "Test détection demande"
    Else
        MsgBox _
            "ECHEC : la demande d'écho-Doppler veineux n'a pas été détectée.", _
            vbExclamation, _
            "Test détection demande"
    End If

End Sub


Public Sub TesterDetectionDictionnairesDDE1()
    DDE_TesterFormulationsInitiales
End Sub

Private Function NormaliserTexteDetectionDemande( _
    ByVal texte As String) As String

    texte = LCase$(texte)
    texte = Replace(texte, "'", "'")
    texte = Replace(texte, Chr(160), " ")
    texte = Replace(texte, vbTab, " ")
    texte = Replace(texte, vbCr, " ")
    texte = Replace(texte, vbLf, " ")

    Do While InStr(texte, "  ") > 0
        texte = Replace(texte, "  ", " ")
    Loop

    NormaliserTexteDetectionDemande = Trim$(texte)

End Function

Private Function ContientFormulationDemandeExamen( _
    ByVal t As String) As Boolean

    ContientFormulationDemandeExamen = False

    If InStr(1, t, "je lui demande", vbTextCompare) > 0 Then
        ContientFormulationDemandeExamen = True
        Exit Function
    End If

    If InStr(1, t, "je demande", vbTextCompare) > 0 Then
        ContientFormulationDemandeExamen = True
        Exit Function
    End If

    If InStr(1, t, "je sollicite", vbTextCompare) > 0 Then
        ContientFormulationDemandeExamen = True
        Exit Function
    End If

    If InStr(1, t, "je complète par", vbTextCompare) > 0 _
    Or InStr(1, t, "je complete par", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "je complète le bilan par", vbTextCompare) > 0 _
    Or InStr(1, t, "je complete le bilan par", vbTextCompare) > 0 _
    Or InStr(1, t, "je compléterai par", vbTextCompare) > 0 _
    Or InStr(1, t, "je completerai par", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "je prescris", vbTextCompare) > 0 _
    Or InStr(1, t, "prescription de", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "merci de réaliser", vbTextCompare) > 0 _
    Or InStr(1, t, "merci de realiser", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "il faut réaliser", vbTextCompare) > 0 _
    Or InStr(1, t, "il faut realiser", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "je l'adresse", vbTextCompare) > 0 _
    Or InStr(1, t, "je l'adresse", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "je te laisse le soin de le diriger", vbTextCompare) > 0 _
    Or InStr(1, t, "je vous laisse le soin de le diriger", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "sera complété par", vbTextCompare) > 0 _
    Or InStr(1, t, "sera complete par", vbTextCompare) > 0 _
    Or InStr(1, t, "nous compléterons par", vbTextCompare) > 0 _
    Or InStr(1, t, "nous completerons par", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    'LETTRES-2N : orientation / adressage vers un centre ou un spécialiste.
    'Exemple : "Nous allons la diriger vers le CCN pour une ablation par radiofréquence".
    If InStr(1, t, "nous allons le diriger", vbTextCompare) > 0 _
    Or InStr(1, t, "nous allons la diriger", vbTextCompare) > 0 _
    Or InStr(1, t, "je vais le diriger", vbTextCompare) > 0 _
    Or InStr(1, t, "je vais la diriger", vbTextCompare) > 0 _
    Or InStr(1, t, "nous allons l'adresser", vbTextCompare) > 0 _
    Or InStr(1, t, "nous allons l'adresser", vbTextCompare) > 0 _
    Or InStr(1, t, "je vais l'adresser", vbTextCompare) > 0 _
    Or InStr(1, t, "je vais l'adresser", vbTextCompare) > 0 _
    Or InStr(1, t, "je l'oriente", vbTextCompare) > 0 _
    Or InStr(1, t, "je l'oriente", vbTextCompare) > 0 _
    Or InStr(1, t, "nous l'orientons", vbTextCompare) > 0 _
    Or InStr(1, t, "nous l'orientons", vbTextCompare) > 0 _
    Or InStr(1, t, "je le confie", vbTextCompare) > 0 _
    Or InStr(1, t, "je la confie", vbTextCompare) > 0 _
    Or InStr(1, t, "nous le confions", vbTextCompare) > 0 _
    Or InStr(1, t, "nous la confions", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    'LETTRES-2L : une proposition explicite de réalisation constitue une demande.
    'Cette règle couvre notamment : "je propose la réalisation d'un écho-Doppler".
    If InStr(1, t, "je propose", vbTextCompare) > 0 _
    Or InStr(1, t, "nous proposons", vbTextCompare) > 0 _
    Or InStr(1, t, "je préconise", vbTextCompare) > 0 _
    Or InStr(1, t, "je preconise", vbTextCompare) > 0 _
    Or InStr(1, t, "nous préconisons", vbTextCompare) > 0 _
    Or InStr(1, t, "nous preconisons", vbTextCompare) > 0 _
    Or InStr(1, t, "je recommande", vbTextCompare) > 0 _
    Or InStr(1, t, "nous recommandons", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    'LETTRES-2D : formulations prospectives de prescription.
    If InStr(1, t, "je prévois", vbTextCompare) > 0 _
    Or InStr(1, t, "je prevois", vbTextCompare) > 0 _
    Or InStr(1, t, "nous prévoyons", vbTextCompare) > 0 _
    Or InStr(1, t, "nous prevoyons", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "je programme", vbTextCompare) > 0 _
    Or InStr(1, t, "nous programmons", vbTextCompare) > 0 _
    Or InStr(1, t, "je vais programmer", vbTextCompare) > 0 _
    Or InStr(1, t, "nous allons programmer", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "je souhaite réaliser", vbTextCompare) > 0 _
    Or InStr(1, t, "je souhaite realiser", vbTextCompare) > 0 _
    Or InStr(1, t, "je souhaite faire réaliser", vbTextCompare) > 0 _
    Or InStr(1, t, "je souhaite faire realiser", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "je vais lui faire réaliser", vbTextCompare) > 0 _
    Or InStr(1, t, "je vais lui faire realiser", vbTextCompare) > 0 _
    Or InStr(1, t, "nous allons réaliser", vbTextCompare) > 0 _
    Or InStr(1, t, "nous allons realiser", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

    If InStr(1, t, "avis spécialisé", vbTextCompare) > 0 _
    Or InStr(1, t, "avis specialise", vbTextCompare) > 0 Then

        ContientFormulationDemandeExamen = True
        Exit Function

    End If

End Function

Private Function ContientExamenEligible( _
    ByVal t As String) As Boolean

    ContientExamenEligible = False

    If InStr(1, t, "test d'effort", vbTextCompare) > 0 _
    Or InStr(1, t, "épreuve d'effort", vbTextCompare) > 0 _
    Or InStr(1, t, "epreuve d'effort", vbTextCompare) > 0 Then

        ContientExamenEligible = True
        Exit Function

    End If

    If InStr(1, t, "scintigraphie", vbTextCompare) > 0 _
    Or InStr(1, t, "coroscanner", vbTextCompare) > 0 _
    Or InStr(1, t, "coro-scanner", vbTextCompare) > 0 _
    Or InStr(1, t, "score calcique", vbTextCompare) > 0 Then

        ContientExamenEligible = True
        Exit Function

    End If

    If InStr(1, t, "scanner", vbTextCompare) > 0 _
    Or InStr(1, t, "angioscanner", vbTextCompare) > 0 _
    Or InStr(1, t, "irm", vbTextCompare) > 0 _
    Or InStr(1, t, "coronarographie", vbTextCompare) > 0 Then

        ContientExamenEligible = True
        Exit Function

    End If

    If InStr(1, t, "écho-doppler", vbTextCompare) > 0 _
    Or InStr(1, t, "echo-doppler", vbTextCompare) > 0 _
    Or InStr(1, t, "echodoppler", vbTextCompare) > 0 _
    Or InStr(1, t, "doppler des vaisseaux", vbTextCompare) > 0 _
    Or InStr(1, t, "troncs supra-aortiques", vbTextCompare) > 0 _
    Or InStr(1, t, "bilan vasculaire", vbTextCompare) > 0 Then

        ContientExamenEligible = True
        Exit Function

    End If

    If InStr(1, t, "avis neurologique", vbTextCompare) > 0 _
    Or InStr(1, t, "neurologue", vbTextCompare) > 0 _
    Or InStr(1, t, "avis vasculaire", vbTextCompare) > 0 _
    Or InStr(1, t, "chirurgien vasculaire", vbTextCompare) > 0 _
    Or InStr(1, t, "avis spécialisé", vbTextCompare) > 0 _
    Or InStr(1, t, "avis specialise", vbTextCompare) > 0 Then

        ContientExamenEligible = True
        Exit Function

    End If

End Function

'===============================================================================
' Fonction conservée sous son nom historique afin de ne pas casser les appels
' existants. En v22, elle demande désormais les nouveaux blocs structurés.
'===============================================================================

Private Function ContientDemandeRythmologiqueExplicite( _
    ByVal t As String) As Boolean

    Dim textePhrases As String
    Dim phrases() As String
    Dim phrase As String
    Dim i As Long

    ContientDemandeRythmologiqueExplicite = False

    textePhrases = Replace(t, "!", ".")
    textePhrases = Replace(textePhrases, "?", ".")
    phrases = Split(textePhrases, ".")

    For i = LBound(phrases) To UBound(phrases)

        phrase = Trim$(CStr(phrases(i)))

        If phrase <> "" Then
            If PRM_ContientCibleRythmologique(phrase) _
            And PRM_ContientIntentionRythmologique(phrase) Then

                ContientDemandeRythmologiqueExplicite = True
                Exit Function

            End If
        End If

    Next i

End Function

Private Function PRM_ContientCibleRythmologique( _
    ByVal phrase As String) As Boolean

    PRM_ContientCibleRythmologique = False

    If InStr(1, phrase, "ablation par radiofréquence", vbTextCompare) > 0 _
    Or InStr(1, phrase, "ablation par radiofrequence", vbTextCompare) > 0 _
    Or InStr(1, phrase, "ablation de fibrillation auriculaire", vbTextCompare) > 0 _
    Or InStr(1, phrase, "ablation de fa", vbTextCompare) > 0 _
    Or InStr(1, phrase, "ablation de flutter", vbTextCompare) > 0 _
    Or InStr(1, phrase, "ablation rythmologique", vbTextCompare) > 0 _
    Or InStr(1, phrase, "exploration électrophysiologique", vbTextCompare) > 0 _
    Or InStr(1, phrase, "exploration electrophysiologique", vbTextCompare) > 0 _
    Or InStr(1, phrase, "électrophysiologie", vbTextCompare) > 0 _
    Or InStr(1, phrase, "electrophysiologie", vbTextCompare) > 0 _
    Or InStr(1, phrase, "avis rythmologique", vbTextCompare) > 0 _
    Or InStr(1, phrase, "rythmologue", vbTextCompare) > 0 _
    Or InStr(1, phrase, "rythmologie", vbTextCompare) > 0 _
    Or InStr(1, phrase, "ccn", vbTextCompare) > 0 Then

        PRM_ContientCibleRythmologique = True

    End If

End Function

Private Function PRM_ContientIntentionRythmologique( _
    ByVal phrase As String) As Boolean

    Dim futurPremierePersonne As Boolean
    Dim verbeOrientation As Boolean

    PRM_ContientIntentionRythmologique = False

    'Intentions explicites suffisamment spécifiques à elles seules.
    If InStr(1, phrase, "je lui demande", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je demande", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je sollicite", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je prescris", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je propose", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous proposons", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je préconise", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je preconise", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous préconisons", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous preconisons", vbTextCompare) > 0 Then

        PRM_ContientIntentionRythmologique = True
        Exit Function

    End If

    If InStr(1, phrase, "je recommande", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous recommandons", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je prévois", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je prevois", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous prévoyons", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous prevoyons", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je programme", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous programmons", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je souhaite", vbTextCompare) > 0 Then

        PRM_ContientIntentionRythmologique = True
        Exit Function

    End If

    If InStr(1, phrase, "je l'adresse", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je l'oriente", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous l'orientons", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je le confie", vbTextCompare) > 0 _
    Or InStr(1, phrase, "je la confie", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous le confions", vbTextCompare) > 0 _
    Or InStr(1, phrase, "nous la confions", vbTextCompare) > 0 Then

        PRM_ContientIntentionRythmologique = True
        Exit Function

    End If

    'Pour "nous allons" / "je vais", exiger en plus un verbe d'action afin
    'd'éviter qu'une phrase comme "nous allons poursuivre le traitement ; une
    'ablation avait été discutée" ne soit prise pour une prescription.
    futurPremierePersonne = _
        (InStr(1, phrase, "nous allons", vbTextCompare) > 0) _
        Or (InStr(1, phrase, "je vais", vbTextCompare) > 0)

    verbeOrientation = _
        (InStr(1, phrase, "diriger", vbTextCompare) > 0) _
        Or (InStr(1, phrase, "adresser", vbTextCompare) > 0) _
        Or (InStr(1, phrase, "orienter", vbTextCompare) > 0) _
        Or (InStr(1, phrase, "confier", vbTextCompare) > 0) _
        Or (InStr(1, phrase, "réaliser", vbTextCompare) > 0) _
        Or (InStr(1, phrase, "realiser", vbTextCompare) > 0) _
        Or (InStr(1, phrase, "programmer", vbTextCompare) > 0)

    If futurPremierePersonne And verbeOrientation Then
        PRM_ContientIntentionRythmologique = True
    End If

End Function

Public Sub TesterDetectionDirigerCCNAblation()

    Dim texteTest As String

    texteTest = _
        "Nous allons donc la diriger vers le CCN pour une " & _
        "ablation par radiofréquence."

    If TexteContientDemandeExamenEligible(texteTest) Then
        MsgBox _
            "OK : l'orientation vers le CCN pour ablation est reconnue " & _
            "comme une demande de prise en charge spécialisée.", _
            vbInformation, _
            "Test détection rythmologie"
    Else
        MsgBox _
            "ÉCHEC : la formulation n'est pas reconnue.", _
            vbExclamation, _
            "Test détection rythmologie"
    End If

End Sub

Public Sub TesterAbsenceDeclenchementMentionRythmologie()

    Dim texteTest As String

    texteTest = _
        "Une ablation par radiofréquence avait été discutée antérieurement au CCN, " & _
        "sans indication retenue à ce jour. Je poursuis le traitement médical."

    If TexteContientDemandeExamenEligible(texteTest) Then
        MsgBox _
            "ECHEC : une simple mention historique déclenche encore une demande.", _
            vbExclamation, _
            "Test absence de faux positif"
    Else
        MsgBox _
            "OK : la mention historique d'ablation / CCN ne déclenche aucune lettre.", _
            vbInformation, _
            "Test absence de faux positif"
    End If

End Sub

Public Function ConstruirePromptDemandeExamenSeule( _
    ByVal texteAnonymise As String) As String

    Dim p As String

    p = ""

    AjouterLignePrompt p, "À partir du courrier médical anonymisé fourni, produis uniquement les demandes d'examens ou de consultations spécialisées explicitement décidées."
    AjouterLignePrompt p, "Ne produis pas de bloc ---CORPS_COURRIER--- et n'écris aucun commentaire hors des blocs de demande."
    AjouterLignePrompt p, "Conserve exactement [[PATIENT]] et n'invente aucun nom, âge, résultat, traitement, indication ou destinataire."
    AjouterLignePrompt p, "Une demande existe seulement si une décision actuelle ou future de prescrire, programmer, faire réaliser, adresser, orienter, diriger ou confier le patient est explicitement liée à l'examen ou à l'avis dans le même passage."
    AjouterLignePrompt p, "Une mention historique, un résultat, une hypothèse, une discussion, une possibilité conditionnelle ou un projet non décidé ne doit produire aucun bloc."
    AjouterLignePrompt p, "Une ablation, une exploration électrophysiologique, un avis rythmologique, un rythmologue ou le CCN ne déclenche une demande que si l'orientation est explicitement décidée dans le même passage."
    AjouterLignePrompt p, "Ne produis pas de courrier séparé pour une échographie cardiaque, une ETT, un Holter rythmique, un Holter ECG, un Holter tensionnel ou une MAPA."
    AjouterLignePrompt p, "Regroupe seulement les examens ayant exactement la même clé de destination réelle et sépare les destinations différentes."
    AjouterLignePrompt p, "Utilise uniquement les clés fournies ; utilise A_COMPLETER si aucune clé ne convient ou si un correspondant précis est requis mais absent."
    AjouterLignePrompt p, "Pour une consultation, un avis ou un geste thérapeutique, un centre ou un établissement cité sans médecin nommé impose CLE_DESTINATION=A_COMPLETER."

    AjouterLignePrompt p, ""
    p = p & ConstruireReglesDestinationsPrompt()
    AjouterLignePrompt p, ""

    AjouterLignePrompt p, "Tu travailles à partir du COURRIER PRINCIPAL DÉJÀ CORRIGÉ. Rédige des courriers complémentaires détaillés, directement utilisables comme brouillons médicaux par la secrétaire."
    AjouterLignePrompt p, "Une demande d'examen commence par Merci de réaliser un/une NOM EXACT DE L'EXAMEN à Monsieur/Madame [[PATIENT]], [âge], [motif précis]."
    AjouterLignePrompt p, "Une demande de consultation spécialisée commence par Merci de prendre en charge Monsieur/Madame [[PATIENT]], [âge], pour [motif précis]."
    AjouterLignePrompt p, "Le premier paragraphe contient obligatoirement le patient, son âge lorsqu'il est disponible, l'examen ou l'avis demandé et le motif réel."
    AjouterLignePrompt p, "Après ce premier paragraphe, reprends systématiquement, lorsqu'ils existent dans le courrier principal, le contexte anamnestique pertinent, les symptômes actuels, les antécédents utiles et les traitements pertinents."
    AjouterLignePrompt p, "L'EXAMEN CLINIQUE doit être repris lorsqu'il existe dans le courrier principal, dans un paragraphe autonome, même s'il paraît peu spécifique de l'examen demandé."
    AjouterLignePrompt p, "L'ÉLECTROCARDIOGRAMME doit être repris lorsqu'il existe dans le courrier principal, dans un paragraphe autonome, même s'il paraît peu spécifique de l'examen demandé."
    AjouterLignePrompt p, "Reprends également, lorsqu'ils existent et sont utiles, l'échographie cardiaque, le Holter, la biologie et les autres examens antérieurs, chacun dans un paragraphe distinct."
    AjouterLignePrompt p, "Ne réduis jamais une demande complexe au seul premier paragraphe Merci de réaliser.... Une demande complexe doit fournir au correspondant le contexte médical nécessaire."
    AjouterLignePrompt p, "Conserve exactement les détails techniques, les valeurs, les dates, les traitements et les consignes d'arrêt ou de reprise utiles."
    AjouterLignePrompt p, "N'invente aucune donnée absente du courrier principal corrigé. Si une rubrique n'est pas présente, omets-la simplement."
    AjouterLignePrompt p, "Le résultat est un BROUILLON MÉDICAL modifiable par la secrétaire avant impression ; il doit néanmoins être suffisamment complet pour être utilisable sans réécriture systématique."
    AjouterLignePrompt p, "Le corps ne contient ni adresse, ni bloc destinataire, ni formule d'appel, ni formule de politesse, ni signature."

    AjouterLignePrompt p, "Chaque demande respecte exactement le format indiqué ci-dessous."
    AjouterFormatDemandeAuPrompt p
    AjouterLignePrompt p, "Si aucune demande éligible n'est explicitement décidée, ne produis aucun bloc."
    AjouterLignePrompt p, ""
    AjouterLignePrompt p, "COURRIER MÉDICAL ANONYMISÉ À TRAITER :"
    AjouterLignePrompt p, "--------------------"
    p = p & texteAnonymise & vbCrLf
    AjouterLignePrompt p, "--------------------"

    ConstruirePromptDemandeExamenSeule = p

End Function

Public Function ConstruireConsignesDemandesExamensDetaillees() As String

    Dim p As String

    p = ""

    AjouterLignePrompt p, "Le corps d'une demande d'examen commence normalement par : Merci de réaliser un/une NOM DE L'EXAMEN à Monsieur/Madame [[PATIENT]], [âge], [motif ou contexte utile]."
    AjouterLignePrompt p, "Le corps d'une demande de consultation spécialisée commence normalement par : Merci de prendre en charge Monsieur/Madame [[PATIENT]], [âge], pour [motif]."
    AjouterLignePrompt p, "Le premier paragraphe de chaque demande doit contenir l'identité du patient, son âge lorsqu'il est présent, l'examen ou l'avis demandé et le motif réel de la demande."
    AjouterLignePrompt p, "Les paragraphes suivants ne reprennent que les antécédents, symptômes, traitements et résultats nécessaires à la compréhension de l'indication ; ne recopie pas tout le courrier principal."
    AjouterLignePrompt p, "Reprends systématiquement le contexte anamnestique pertinent lorsqu'il existe."
    AjouterLignePrompt p, "Reprends systématiquement l'examen clinique dans un paragraphe autonome lorsqu'il existe dans le courrier principal."
    AjouterLignePrompt p, "Reprends systématiquement l'électrocardiogramme dans un paragraphe autonome lorsqu'il existe dans le courrier principal."
    AjouterLignePrompt p, "Dans une demande complémentaire, présente séparément les données d'examen clinique, d'électrocardiogramme, d'échographie cardiaque, de Holter et de bilan biologique lorsqu'elles sont reprises."
    AjouterLignePrompt p, "Fais commencer ces paragraphes de façon naturelle, par exemple L'électrocardiogramme..., L'échographie cardiaque..., Le Holter rythmique... ou Le bilan biologique..., sans créer de titre artificiel."
    AjouterLignePrompt p, "Conserve exactement toute consigne d'arrêt ou de reprise d'un médicament avant l'examen, avec le nom du médicament et le délai indiqués."
    AjouterLignePrompt p, "Termine, si cela apporte une information utile et non redondante, par une formule brève du type Merci de vérifier..., Merci de confirmer... ou Merci d'évaluer...."
    AjouterLignePrompt p, "Une demande de score calcique doit rester très brève et reprendre seulement le terrain de risque ou l'objectif préventif réellement présent."
    AjouterLignePrompt p, "Une demande d'écho-Doppler doit reprendre seulement le terrain vasculaire, les symptômes, les anomalies de pouls, les antécédents vasculaires et l'objectif réellement utiles."
    AjouterLignePrompt p, "Une demande de test d'effort ou de scintigraphie myocardique doit reprendre les symptômes, le contexte coronarien ou la revascularisation, l'ECG pertinent, les examens antérieurs utiles et toute consigne d'arrêt médicamenteux présente."
    AjouterLignePrompt p, "Une demande de coroscanner ou d'angioscanner coronaire doit reprendre les symptômes, les facteurs de risque, les examens antérieurs et la question anatomique réellement posée."
    AjouterLignePrompt p, "Une demande de consultation spécialisée doit être concise, exposer la problématique et les éléments utiles, puis formuler clairement la question ou la prise en charge attendue."
    AjouterLignePrompt p, "Le corps d'une demande ne doit contenir ni adresse, ni bloc destinataire, ni formule d'appel, ni formule de politesse, ni signature."

    ConstruireConsignesDemandesExamensDetaillees = p

End Function


Public Function ConstruireConsignesDemandesStructureesAPI() As String

    Dim p As String

    p = ""

    AjouterLignePrompt p, "Après le bloc du courrier principal, produis un bloc de demande complémentaire uniquement lorsqu'une décision explicite et actuelle ou future de prescrire, programmer, faire réaliser, adresser, orienter, diriger ou confier le patient est réellement exprimée."
    AjouterLignePrompt p, "Une simple mention historique, un résultat d'examen, une possibilité discutée, une hypothèse, une formulation conditionnelle ou un projet non décidé ne doit déclencher aucun courrier complémentaire."
    AjouterLignePrompt p, "Sont des formulations explicites de demande, lorsqu'elles sont liées à l'examen ou à l'avis dans le même passage : Je prévois, Je programme, Je lui demande, Je complète par, Je propose la réalisation de, Merci de réaliser, Nous allons l'adresser, Nous allons le/la diriger, Je l'oriente ou Je le/la confie."
    AjouterLignePrompt p, "Pour une ablation, une exploration électrophysiologique, un avis rythmologique, un rythmologue ou le CCN, ne produis une demande que si la décision d'orientation est explicitement exprimée dans la même phrase ou le même passage immédiat."
    AjouterLignePrompt p, "Ne produis pas de courrier complémentaire séparé pour une échographie cardiaque, une échographie transthoracique, une ETT, un Holter rythmique, un Holter ECG, un Holter tensionnel ou une MAPA."
    AjouterLignePrompt p, "Regroupe dans un même bloc les examens ayant exactement la même clé de destination réelle ; produis des blocs distincts pour des destinations différentes."
    AjouterLignePrompt p, "Utilise exclusivement les clés de destination fournies ci-dessous ; utilise A_COMPLETER si aucune clé ne convient ou si le correspondant précis n'est pas connu."
    AjouterLignePrompt p, "Pour une consultation, un avis ou un geste thérapeutique, si le texte ne cite qu'un établissement, un centre, un service ou un acronyme institutionnel sans nommer de médecin correspondant, utilise CLE_DESTINATION=A_COMPLETER."

    AjouterLignePrompt p, ""
    p = p & ConstruireReglesDestinationsPrompt()
    AjouterLignePrompt p, ""
    p = p & ConstruireConsignesDemandesExamensDetaillees()
    AjouterLignePrompt p, "Chaque demande doit respecter exactement le format indiqué ci-dessous."
    AjouterFormatDemandeAuPrompt p
    AjouterLignePrompt p, "N'écris aucun bloc DEMANDE_DESTINATION si aucune demande éligible n'est explicitement décidée dans le courrier source."
    AjouterLignePrompt p, "N'écris rien en dehors du bloc CORPS_COURRIER et des éventuels blocs DEMANDE_DESTINATION."

    ConstruireConsignesDemandesStructureesAPI = p

End Function

Private Function ConstruireReglesDestinationsPrompt() As String

    Dim p As String

    p = ""
    AjouterLignePrompt p, "Règles de destination issues de la base Excel :"

    AjouterRegleDestinationDepuisBase p, "TEST_EFFORT", "Test d'effort ou épreuve d'effort"
    AjouterRegleDestinationDepuisBase p, "SCINTIGRAPHIE_MYOCARDIQUE", "Scintigraphie myocardique"
    AjouterRegleDestinationDepuisBase p, "COROSCANNER", "Coroscanner ou scanner coronaire"
    AjouterRegleDestinationDepuisBase p, "SCORE_CALCIQUE", "Score calcique"
    AjouterRegleDestinationDepuisBase p, "ECHODOPPLER_VAISSEAUX_DU_COU", "Écho-Doppler des vaisseaux du cou ou des troncs supra-aortiques"
    AjouterRegleDestinationDepuisBase p, "ECHODOPPLER_MEMBRES_INFERIEURS", "Écho-Doppler des membres inférieurs"
    AjouterRegleDestinationDepuisBase p, "ECHODOPPLER_ARTERIEL", "Écho-Doppler artériel"
    AjouterRegleDestinationDepuisBase p, "ECHODOPPLER_VEINEUX", "Écho-Doppler veineux"
    AjouterRegleDestinationDepuisBase p, "AVIS_NEUROLOGIE", "Avis neurologique"
    AjouterRegleDestinationDepuisBase p, "AVIS_RYTHMOLOGIE", "Avis rythmologique"
    AjouterRegleDestinationDepuisBase p, "AVIS_VASCULAIRE", "Avis vasculaire"

    AjouterLignePrompt p, "- Tout autre examen ou avis spécialisé : " & PREFIXE_CLE_DESTINATION & "A_COMPLETER"

    ConstruireReglesDestinationsPrompt = p

End Function

Private Sub AjouterFormatDemandeAuPrompt(ByRef p As String)

    AjouterLignePrompt p, BALISE_DEBUT_DEMANDE_DESTINATION
    AjouterLignePrompt p, PREFIXE_CLE_DESTINATION & "CLE_EXACTE"
    AjouterLignePrompt p, BALISE_DEBUT_CORPS_DESTINATION
    AjouterLignePrompt p, "Corps médical de la demande"
    AjouterLignePrompt p, BALISE_FIN_CORPS_DESTINATION
    AjouterLignePrompt p, BALISE_FIN_DEMANDE_DESTINATION

End Sub

Private Sub AjouterLignePrompt( _
    ByRef p As String, _
    ByVal ligne As String)

    p = p & ligne & vbCrLf

End Sub

Public Sub TesterPromptCorpus4307()

    Dim p As String
    Dim texteTest As String

    texteTest = _
        "Je revois Madame [[PATIENT]], 70 ans. " & _
        "Je lui demande une scintigraphie myocardique."

    p = ConstruirePromptReecritureMedicale(texteTest)

    If InStr(1, p, "---CORPS_COURRIER---", vbTextCompare) > 0 _
    And InStr(1, p, BALISE_DEBUT_DEMANDE_DESTINATION, vbTextCompare) > 0 _
    And InStr(1, p, "Merci de réaliser", vbTextCompare) > 0 _
    And InStr(1, p, texteTest, vbTextCompare) > 0 Then

        MsgBox _
            "OK : le prompt corpus 4307 est construit." & vbCrLf & _
            "Longueur : " & Len(p) & " caractères.", _
            vbInformation, _
            "Test prompt corpus 4307"

    Else

        MsgBox _
            "ÉCHEC : le prompt corpus 4307 est incomplet.", _
            vbExclamation, _
            "Test prompt corpus 4307"

    End If

End Sub

Private Sub AjouterRegleDestinationDepuisBase( _
    ByRef prompt As String, _
    ByVal typeExamen As String, _
    ByVal libelle As String)

    prompt = prompt & _
        "- " & libelle & " [" & typeExamen & "] : " & _
        PREFIXE_CLE_DESTINATION & _
        CleDestinationPourPrompt(typeExamen) & vbCrLf

End Sub

Private Function CleDestinationPourPrompt( _
    ByVal typeExamen As String) As String

    Dim recordDestination As String
    Dim cleDestination As String

    On Error GoTo GestionErreur

    recordDestination = _
        DestinationParDefautPourTypeExamen(typeExamen)

    If Trim$(recordDestination) = "" Then
        CleDestinationPourPrompt = "A_COMPLETER"
        Exit Function
    End If

    cleDestination = _
        Trim$(ChampDestination(recordDestination, 10))

    If cleDestination = "" Then
        cleDestination = "A_COMPLETER"
    End If

    CleDestinationPourPrompt = UCase$(cleDestination)
    Exit Function

GestionErreur:

    CleDestinationPourPrompt = "A_COMPLETER"

End Function
