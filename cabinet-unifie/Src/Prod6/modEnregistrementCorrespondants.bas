Attribute VB_Name = "modEnregistrementCorrespondants"
Option Explicit

'===============================================================================
' MODULE : modEnregistrementCorrespondants
' VERSION : STRUCT-9A
'
' Objet :
' - enregistrer un nouveau généraliste ou spécialiste depuis
'   frmNouveauCorrespondant ;
' - recontrôler les doublons juste avant écriture ;
' - sauvegarder physiquement la base avant modification ;
' - générer automatiquement GEN_#### / SPE_#### / SPT_#### ;
' - écrire le classeur une seule fois, à la fin ;
' - en cas d'erreur avant Save, fermer sans sauvegarder ;
' - si un spécialiste existe déjà, ne pas dupliquer sa fiche :
'   ajouter uniquement les nouveaux TypeExamen dans Specialistes_ParType
'   et compléter TypesExamensPossibles.
'
' Sécurité de priorité :
' - un nouveau destinataire n'écrase pas un choix existant ;
' - si un TypeExamen possède déjà une destination active, la nouvelle
'   association reçoit Priorite=100 ;
' - si aucune destination active n'existe pour ce type, Priorite=1.
'===============================================================================

Private Const NCE_XL_UP As Long = -4162
Private Const NCE_XL_TO_LEFT As Long = -4159
Private Const NCE_ID_NOUVELLE_STRUCTURE As String = "__NOUVELLE_STRUCTURE__"

Public Sub NCE_EnregistrerCorrespondant()

    Dim verrou As String, verrouAcquis As Boolean
    Dim frm As Object

    Dim typeCorrespondant As String
    Dim sexe As String
    Dim prenom As String
    Dim nom As String
    Dim idStructure As String
    Dim adresse1 As String
    Dim adresse2 As String
    Dim codePostal As String
    Dim ville As String
    Dim telephone As String
    Dim relation As String
    Dim typesSelectionnes As String

    Dim xlApp As Object
    Dim wb As Object
    Dim wsGen As Object
    Dim wsSpe As Object
    Dim wsSPT As Object
    Dim wsSaisie As Object

    Dim ligneGenExacte As Long
    Dim ligneSpeExacte As Long
    Dim idSpecialisteExistant As String
    Dim typesNouveaux As String
    Dim doublonsProbables As String
    Dim idSpecialisteAIgnorer As String

    Dim donnees As Object
    Dim resumeFiche As String
    Dim reponse As VbMsgBoxResult

    Dim cheminSauvegarde As String

    Dim numeroErreur As Long
    Dim descriptionErreur As String

    On Error GoTo GestionErreur

    Set frm = frmNouveauCorrespondant

    '===========================================================================
    ' 1. Lecture et validation de la saisie
    '===========================================================================
    typeCorrespondant = _
        Trim$(frm.Controls("cmbTypeCorrespondant").Text)

    sexe = _
        Trim$(frm.Controls("cmbSexe").Text)

    prenom = _
        Trim$(frm.Controls("txtPrenom").Text)

    nom = _
        Trim$(frm.Controls("txtNom").Text)

    idStructure = _
        NCE_IDStructureSelectionnee(frm)

    adresse1 = _
        Trim$(frm.Controls("txtAdresse1").Text)

    adresse2 = _
        Trim$(frm.Controls("txtAdresse2").Text)

    codePostal = _
        Trim$(frm.Controls("txtCodePostal").Text)

    ville = _
        Trim$(frm.Controls("txtVille").Text)

    telephone = _
        Trim$(frm.Controls("txtTelephone").Text)

    relation = _
        Trim$(frm.Controls("cmbTutoiement").Text)

    If relation = "" Then relation = "vous"

    typesSelectionnes = _
        NCE_TypesSelectionnes(frm)

    If Not NCE_ValiderSaisie( _
        frm, _
        typeCorrespondant, _
        sexe, _
        prenom, _
        nom, _
        idStructure, _
        adresse1, _
        codePostal, _
        ville, _
        typesSelectionnes) Then

        If verrouAcquis Then modFichiers.RelacherVerrou verrou
    Exit Sub

    End If

    '===========================================================================
    ' 2. Précontrôle en lecture seule
    '===========================================================================
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.AutomationSecurity = 3
    xlApp.EnableEvents = False

    Set wb = xlApp.Workbooks.Open( _
        CHEMIN_BASE_CORRESPONDANTS, _
        0, _
        True)

    Set wsGen = wb.Worksheets("Generalistes")
    Set wsSpe = wb.Worksheets("Specialistes")
    Set wsSPT = wb.Worksheets("Specialistes_ParType")
    Set wsSaisie = wb.Worksheets("Saisie_Correspondants")

    ligneGenExacte = _
        NCE_TrouverLigneExacte( _
            wsGen, _
            prenom, _
            nom, _
            "NomAffichage")

    ligneSpeExacte = _
        NCE_TrouverLigneExacte( _
            wsSpe, _
            prenom, _
            nom, _
            "NomDestinataire")

    If StrComp( _
        typeCorrespondant, _
        "Généraliste", _
        vbTextCompare) = 0 Then

        If ligneGenExacte > 0 _
        Or ligneSpeExacte > 0 Then

            NC_VerifierDoublonsFormulaire

            MsgBox _
                "Enregistrement refusé : ce correspondant existe déjà " & _
                "dans la base.", _
                vbExclamation, _
                "Doublon correspondant"

            GoTo SortieLectureSeule

        End If

    Else

        'Un spécialiste exact déjà présent sera réutilisé.
        'En revanche, une identité identique déjà classée comme généraliste
        'est volontairement bloquée : cette situation mérite une vérification.
        If ligneGenExacte > 0 Then

            NC_VerifierDoublonsFormulaire

            MsgBox _
                "Enregistrement interrompu : cette identité existe déjà " & _
                "dans Generalistes." & vbCrLf & vbCrLf & _
                "Vérifiez la fiche avant de créer une fiche spécialiste.", _
                vbExclamation, _
                "Correspondant déjà présent"

            GoTo SortieLectureSeule

        End If

    End If

    'Correspondances probables hors éventuel spécialiste exact réutilisé.
    idSpecialisteAIgnorer = ""

    If ligneSpeExacte > 0 Then
        idSpecialisteAIgnorer = _
            NCE_LireCelluleParEntete( _
                wsSpe, _
                ligneSpeExacte, _
                "ID")
    End If

    doublonsProbables = _
        NCE_ListerDoublonsProbables( _
            wsGen, _
            wsSpe, _
            prenom, _
            nom, _
            telephone, _
            idSpecialisteAIgnorer)

    If Trim$(doublonsProbables) <> "" Then

        reponse = MsgBox( _
            "Des correspondants proches existent déjà :" & _
            vbCrLf & vbCrLf & _
            doublonsProbables & _
            vbCrLf & _
            "Confirmez-vous qu'il s'agit bien de la fiche à traiter ?", _
            vbExclamation + vbYesNo, _
            "Correspondants proches")

        If reponse <> vbYes Then
            GoTo SortieLectureSeule
        End If

    End If

    If StrComp( _
        typeCorrespondant, _
        "Spécialiste", _
        vbTextCompare) = 0 _
    And ligneSpeExacte > 0 Then

        idSpecialisteExistant = _
            Trim$(NCE_LireCelluleParEntete( _
                wsSpe, _
                ligneSpeExacte, _
                "ID"))

        typesNouveaux = _
            NCE_TypesNonAssocies( _
                wsSPT, _
                idSpecialisteExistant, _
                typesSelectionnes)

        If Trim$(typesNouveaux) = "" Then

            MsgBox _
                "Ce spécialiste existe déjà et tous les types sélectionnés " & _
                "lui sont déjà associés." & vbCrLf & vbCrLf & _
                "Aucune modification n'est nécessaire.", _
                vbInformation, _
                "Spécialiste déjà complet"

            GoTo SortieLectureSeule

        End If

        resumeFiche = _
            "Le spécialiste existe déjà :" & vbCrLf & _
            NCE_LireCelluleParEntete( _
                wsSpe, _
                ligneSpeExacte, _
                "NomDestinataire") & _
            vbCrLf & vbCrLf & _
            "La fiche existante ne sera pas dupliquée." & vbCrLf & _
            "Seuls les nouveaux types suivants seront ajoutés :" & vbCrLf & _
            NCE_TypesEnLignes(typesNouveaux)

        reponse = MsgBox( _
            resumeFiche & vbCrLf & _
            "Continuer ?", _
            vbQuestion + vbYesNo, _
            "Compléter un spécialiste existant")

        If reponse <> vbYes Then
            GoTo SortieLectureSeule
        End If

    Else

        Set donnees = _
            NCE_CalculerDonneesDepuisSaisie( _
                xlApp, _
                wsSaisie, _
                typeCorrespondant, _
                sexe, _
                prenom, _
                nom, _
                idStructure, _
                adresse1, _
                adresse2, _
                codePostal, _
                ville, _
                telephone, _
                relation, _
                typesSelectionnes)

        resumeFiche = _
            NCE_ResumeNouvelleFiche( _
                typeCorrespondant, _
                donnees)

        reponse = MsgBox( _
            resumeFiche & vbCrLf & _
            "Enregistrer cette fiche ?", _
            vbQuestion + vbYesNo, _
            "Confirmer l'enregistrement")

        If reponse <> vbYes Then
            GoTo SortieLectureSeule
        End If

    End If

    'Fermer la lecture seule avant la sauvegarde physique.
    wb.Close SaveChanges:=False
    Set wb = Nothing

    xlApp.Quit
    Set xlApp = Nothing

    Set wsGen = Nothing
    Set wsSpe = Nothing
    Set wsSPT = Nothing
    Set wsSaisie = Nothing

    '===========================================================================
    ' 3. Sauvegarde physique de la base AVANT toute écriture
    '===========================================================================
    verrou = CreateObject("Scripting.FileSystemObject").GetBaseName(CHEMIN_BASE_CORRESPONDANTS)
    If Not modFichiers.AcquerirVerrou(verrou) Then Err.Raise vbObjectError + 726, , "Base correspondants occupee."
    verrouAcquis = True
    cheminSauvegarde = _
        NCE_CreerSauvegardeBase()

    '===========================================================================
    ' 4. Ouverture en écriture + recontrôle immédiat
    '===========================================================================
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.AutomationSecurity = 3
    xlApp.EnableEvents = False

    Set wb = xlApp.Workbooks.Open( _
        CHEMIN_BASE_CORRESPONDANTS, _
        0, _
        False)

    If wb.ReadOnly Then

        MsgBox _
            "La base Excel est actuellement en lecture seule." & vbCrLf & _
            "Fermez-la dans Excel puis recommencez." & vbCrLf & vbCrLf & _
            "Une sauvegarde a déjà été créée :" & vbCrLf & _
            cheminSauvegarde, _
            vbExclamation, _
            "Base correspondants"

        GoTo SortieSansSauver

    End If

    Set wsGen = wb.Worksheets("Generalistes")
    Set wsSpe = wb.Worksheets("Specialistes")
    Set wsSPT = wb.Worksheets("Specialistes_ParType")
    Set wsSaisie = wb.Worksheets("Saisie_Correspondants")

    ligneGenExacte = _
        NCE_TrouverLigneExacte( _
            wsGen, _
            prenom, _
            nom, _
            "NomAffichage")

    ligneSpeExacte = _
        NCE_TrouverLigneExacte( _
            wsSpe, _
            prenom, _
            nom, _
            "NomDestinataire")

    If StrComp( _
        typeCorrespondant, _
        "Généraliste", _
        vbTextCompare) = 0 Then

        If ligneGenExacte > 0 _
        Or ligneSpeExacte > 0 Then

            MsgBox _
                "L'enregistrement est annulé : un doublon est apparu " & _
                "dans la base depuis le premier contrôle.", _
                vbExclamation, _
                "Recontrôle des doublons"

            GoTo SortieSansSauver

        End If

        Set donnees = _
            NCE_CalculerDonneesDepuisSaisie( _
                xlApp, _
                wsSaisie, _
                typeCorrespondant, _
                sexe, _
                prenom, _
                nom, _
                idStructure, _
                adresse1, _
                adresse2, _
                codePostal, _
                ville, _
                telephone, _
                relation, _
                typesSelectionnes)

        NCE_AjouterGeneraliste _
            wsGen, _
            donnees

    Else

        If ligneGenExacte > 0 Then

            MsgBox _
                "L'enregistrement est annulé : cette identité est maintenant " & _
                "présente dans Generalistes.", _
                vbExclamation, _
                "Recontrôle des doublons"

            GoTo SortieSansSauver

        End If

        If idSpecialisteExistant <> "" Then

            'Le spécialiste devait déjà exister au premier contrôle.
            If ligneSpeExacte = 0 Then

                MsgBox _
                    "La fiche spécialiste a changé depuis le premier contrôle." & _
                    vbCrLf & _
                    "Aucune écriture n'a été effectuée. Relancez la saisie.", _
                    vbExclamation, _
                    "Base modifiée"

                GoTo SortieSansSauver

            End If

            If StrComp( _
                Trim$(NCE_LireCelluleParEntete( _
                    wsSpe, _
                    ligneSpeExacte, _
                    "ID")), _
                idSpecialisteExistant, _
                vbTextCompare) <> 0 Then

                MsgBox _
                    "La fiche spécialiste ne correspond plus au premier contrôle." & _
                    vbCrLf & _
                    "Aucune écriture n'a été effectuée.", _
                    vbExclamation, _
                    "Base modifiée"

                GoTo SortieSansSauver

            End If

            typesNouveaux = _
                NCE_TypesNonAssocies( _
                    wsSPT, _
                    idSpecialisteExistant, _
                    typesSelectionnes)

            If Trim$(typesNouveaux) = "" Then

                MsgBox _
                    "Aucune modification à enregistrer : les types sont " & _
                    "déjà présents.", _
                    vbInformation, _
                    "Spécialiste existant"

                GoTo SortieSansSauver

            End If

            NCE_AjouterTypesSpecialisteExistant _
                wsSpe, _
                wsSPT, _
                ligneSpeExacte, _
                idSpecialisteExistant, _
                typesNouveaux

        Else

            'Un spécialiste exact créé entre les deux contrôles bloque l'opération.
            If ligneSpeExacte > 0 Then

                MsgBox _
                    "L'enregistrement est annulé : ce spécialiste vient " & _
                    "d'être ajouté à la base." & vbCrLf & _
                    "Relancez la saisie afin de compléter sa fiche existante.", _
                    vbExclamation, _
                    "Recontrôle des doublons"

                GoTo SortieSansSauver

            End If

            Set donnees = _
                NCE_CalculerDonneesDepuisSaisie( _
                    xlApp, _
                    wsSaisie, _
                    typeCorrespondant, _
                    sexe, _
                    prenom, _
                    nom, _
                    idStructure, _
                    adresse1, _
                    adresse2, _
                    codePostal, _
                    ville, _
                    telephone, _
                    relation, _
                    typesSelectionnes)

            NCE_AjouterNouveauSpecialiste _
                wsSpe, _
                wsSPT, _
                donnees

        End If

    End If

    'La zone technique ne conserve pas la saisie personnelle.
    NCE_ReinitialiserZoneSaisie wsSaisie

    'Une seule écriture disque, après toutes les opérations réussies.
    wb.Save
    If Not wb.Saved Then Err.Raise vbObjectError + 727, , "Sauvegarde des correspondants non confirmee."

    wb.Close SaveChanges:=False
    Set wb = Nothing

    xlApp.Quit
    Set xlApp = Nothing

    MsgBox _
        "Enregistrement terminé." & vbCrLf & vbCrLf & _
        "Sauvegarde préalable :" & vbCrLf & _
        cheminSauvegarde, _
        vbInformation, _
        "Correspondant enregistré"

    On Error Resume Next
    Unload frmNouveauCorrespondant
    On Error GoTo 0

    If verrouAcquis Then modFichiers.RelacherVerrou verrou
    Exit Sub

SortieLectureSeule:

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If

    Set wsGen = Nothing
    Set wsSpe = Nothing
    Set wsSPT = Nothing
    Set wsSaisie = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Set donnees = Nothing

    On Error GoTo 0
    If verrouAcquis Then modFichiers.RelacherVerrou verrou
    Exit Sub

SortieSansSauver:

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If

    Set wsGen = Nothing
    Set wsSpe = Nothing
    Set wsSPT = Nothing
    Set wsSaisie = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Set donnees = Nothing

    On Error GoTo 0
    If verrouAcquis Then modFichiers.RelacherVerrou verrou
    Exit Sub

GestionErreur:

    numeroErreur = Err.Number
    descriptionErreur = Err.description

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If

    Set wsGen = Nothing
    Set wsSpe = Nothing
    Set wsSPT = Nothing
    Set wsSaisie = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Set donnees = Nothing

    On Error GoTo 0

    MsgBox _
        "Le correspondant n'a pas été enregistré." & vbCrLf & vbCrLf & _
        CStr(numeroErreur) & " - " & descriptionErreur, _
        vbExclamation, _
        "Enregistrement correspondant"

    If verrouAcquis Then modFichiers.RelacherVerrou verrou
End Sub

'===============================================================================
' VALIDATION
'===============================================================================
Private Function NCE_ValiderSaisie( _
    ByVal frm As Object, _
    ByVal typeCorrespondant As String, _
    ByVal sexe As String, _
    ByVal prenom As String, _
    ByVal nom As String, _
    ByVal idStructure As String, _
    ByVal adresse1 As String, _
    ByVal codePostal As String, _
    ByVal ville As String, _
    ByVal typesSelectionnes As String) As Boolean

    NCE_ValiderSaisie = False

    If StrComp(typeCorrespondant, "Généraliste", vbTextCompare) <> 0 _
    And StrComp(typeCorrespondant, "Spécialiste", vbTextCompare) <> 0 Then

        MsgBox _
            "Choisissez Généraliste ou Spécialiste.", _
            vbExclamation, _
            "Nouveau correspondant"

        Exit Function

    End If

    If sexe = "" Then

        MsgBox _
            "Le sexe du correspondant est obligatoire.", _
            vbExclamation, _
            "Nouveau correspondant"

        frm.Controls("cmbSexe").SetFocus
        Exit Function

    End If

    If prenom = "" Then

        MsgBox _
            "Le prénom est obligatoire.", _
            vbExclamation, _
            "Nouveau correspondant"

        frm.Controls("txtPrenom").SetFocus
        Exit Function

    End If

    If nom = "" Then

        MsgBox _
            "Le nom est obligatoire.", _
            vbExclamation, _
            "Nouveau correspondant"

        frm.Controls("txtNom").SetFocus
        Exit Function

    End If

    If StrComp( _
        idStructure, _
        NCE_ID_NOUVELLE_STRUCTURE, _
        vbTextCompare) = 0 Then

        MsgBox _
            "Créez ou sélectionnez d'abord la structure.", _
            vbExclamation, _
            "Nouveau correspondant"

        Exit Function

    End If

    If idStructure = "" Then

        If adresse1 = "" _
        Or codePostal = "" _
        Or ville = "" Then

            MsgBox _
                "Sans structure sélectionnée, renseignez au minimum " & _
                "l'adresse, le code postal et la ville.", _
                vbExclamation, _
                "Nouveau correspondant"

            Exit Function

        End If

    End If

    If codePostal <> "" Then

        If Len(NCE_GarderChiffres(codePostal)) <> 5 Then

            If MsgBox( _
                "Le code postal ne comporte pas 5 chiffres." & vbCrLf & _
                "Continuer malgré tout ?", _
                vbExclamation + vbYesNo, _
                "Code postal") <> vbYes Then

                Exit Function

            End If

        End If

    End If

    If StrComp( _
        typeCorrespondant, _
        "Spécialiste", _
        vbTextCompare) = 0 Then

        If Trim$(typesSelectionnes) = "" Then

            MsgBox _
                "Pour un spécialiste, sélectionnez au moins un type " & _
                "d'examen ou d'avis.", _
                vbExclamation, _
                "Nouveau correspondant"

            Exit Function

        End If

    End If

    NCE_ValiderSaisie = True

End Function

'===============================================================================
' CALCUL PAR Saisie_Correspondants
'===============================================================================
Private Function NCE_CalculerDonneesDepuisSaisie( _
    ByVal xlApp As Object, _
    ByVal ws As Object, _
    ByVal typeCorrespondant As String, _
    ByVal sexe As String, _
    ByVal prenom As String, _
    ByVal nom As String, _
    ByVal idStructure As String, _
    ByVal adresse1 As String, _
    ByVal adresse2 As String, _
    ByVal codePostal As String, _
    ByVal ville As String, _
    ByVal telephone As String, _
    ByVal relation As String, _
    ByVal typesExamens As String) As Object

    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")

    modFichiers.EcrireCelluleTexte ws.Range("B4"), typeCorrespondant
    modFichiers.EcrireCelluleTexte ws.Range("B5"), sexe
    modFichiers.EcrireCelluleTexte ws.Range("B6"), prenom
    modFichiers.EcrireCelluleTexte ws.Range("B7"), nom
    modFichiers.EcrireCelluleTexte ws.Range("B8"), idStructure
    modFichiers.EcrireCelluleTexte ws.Range("B9"), adresse1
    modFichiers.EcrireCelluleTexte ws.Range("B10"), adresse2
    modFichiers.EcrireCelluleTexte ws.Range("B11"), codePostal
    modFichiers.EcrireCelluleTexte ws.Range("B12"), ville
    modFichiers.EcrireCelluleTexte ws.Range("B13"), telephone
    modFichiers.EcrireCelluleTexte ws.Range("B14"), relation
    modFichiers.EcrireCelluleTexte ws.Range("B15"), typesExamens
    ws.Range("B16").Value = ""

    modBaseCorrespondants.PROD_ActualiserPlagesStructures ws.Parent
    xlApp.CalculateFull

    d.Add "TypeCorrespondant", typeCorrespondant
    d.Add "Titre", NCE_LireRange(ws.Range("E4"))
    d.Add "Civilite", NCE_LireRange(ws.Range("E5"))
    d.Add "Prenom", prenom
    d.Add "Nom", UCase$(nom)
    d.Add "NomAffichage", NCE_LireRange(ws.Range("E6"))
    d.Add "NomDestinataire", NCE_LireRange(ws.Range("E7"))
    d.Add "NomStructure", NCE_LireRange(ws.Range("E8"))
    d.Add "Adresse1", NCE_LireRange(ws.Range("E9"))
    d.Add "Adresse2", NCE_LireRange(ws.Range("E10"))
    d.Add "CodePostal", NCE_LireRange(ws.Range("E11"))
    d.Add "Ville", NCE_LireRange(ws.Range("E12"))
    d.Add "Telephone", NCE_LireRange(ws.Range("E13"))
    d.Add "FormuleAppel", NCE_LireRange(ws.Range("E14"))
    d.Add "FormulePolitesse", NCE_LireRange(ws.Range("E15"))
    d.Add "BlocStructureComplet", NCE_LireRange(ws.Range("E16"))
    d.Add "BlocDestinataireComplet", NCE_LireRange(ws.Range("E17"))
    d.Add "Actif", NCE_LireRange(ws.Range("E18"))
    d.Add "AValider", NCE_LireRange(ws.Range("E19"))
    d.Add "NombreOccurrences", NCE_LireRange(ws.Range("E20"))
    d.Add "MethodeMiseAJour", NCE_LireRange(ws.Range("E21"))
    d.Add "TypeExamenPrincipal", NCE_LireRange(ws.Range("E22"))
    d.Add "TypesExamensPossibles", NCE_LireRange(ws.Range("E23"))
    d.Add "CommentaireCreation", NCE_LireRange(ws.Range("E25"))
    d.Add "ID_Structure", NCE_LireRange(ws.Range("E26"))
    d.Add "TutoiementVouvoiement", relation

    Set NCE_CalculerDonneesDepuisSaisie = d

End Function

Private Sub NCE_ReinitialiserZoneSaisie(ByVal ws As Object)

    ws.Range("B4:B13").ClearContents
    ws.Range("B14").Value = "vous"
    ws.Range("B15:B16").ClearContents

End Sub

'===============================================================================
' ECRITURE GENERALISTE
'===============================================================================
Private Sub NCE_AjouterGeneraliste( _
    ByVal ws As Object, _
    ByVal d As Object)

    Dim colID As Long
    Dim ligne As Long
    Dim nouvelID As String

    colID = _
        NCE_TrouverColonneEntete(ws, "ID")

    If colID = 0 Then
        Err.Raise vbObjectError + 720, , "Colonne ID absente de Generalistes."
    End If

    nouvelID = _
        NCE_ProchainID(ws, colID, "GEN_", 4)

    ligne = _
        ws.Cells(ws.Rows.Count, colID).End(NCE_XL_UP).Row + 1

    NCE_Ecrire ws, ligne, "ID", nouvelID
    NCE_Ecrire ws, ligne, "NomAffichage", CStr(d("NomAffichage"))
    NCE_Ecrire ws, ligne, "Titre", CStr(d("Titre"))
    NCE_Ecrire ws, ligne, "Civilite", CStr(d("Civilite"))
    NCE_Ecrire ws, ligne, "PrenomOuInitiale", CStr(d("Prenom"))
    NCE_Ecrire ws, ligne, "Nom", CStr(d("Nom"))
    NCE_Ecrire ws, ligne, "Structure", CStr(d("NomStructure"))
    NCE_Ecrire ws, ligne, "Adresse1", CStr(d("Adresse1"))
    NCE_Ecrire ws, ligne, "Adresse2", CStr(d("Adresse2"))
    NCE_Ecrire ws, ligne, "CodePostal", CStr(d("CodePostal"))
    NCE_Ecrire ws, ligne, "Ville", CStr(d("Ville"))
    NCE_Ecrire ws, ligne, "Telephone", CStr(d("Telephone"))
    NCE_Ecrire ws, ligne, "BlocDestinataireComplet", CStr(d("BlocDestinataireComplet"))
    NCE_Ecrire ws, ligne, "FormuleAppel", CStr(d("FormuleAppel"))
    NCE_Ecrire ws, ligne, "TutoiementVouvoiement", CStr(d("TutoiementVouvoiement"))
    NCE_Ecrire ws, ligne, "NombreOccurrences", "0"
    NCE_Ecrire ws, ligne, "SourceLignes", ""
    NCE_Ecrire ws, ligne, "Actif", "Oui"
    NCE_Ecrire ws, ligne, "AValider", "Non"
    NCE_Ecrire ws, ligne, "CommentaireExtraction", CStr(d("CommentaireCreation"))
    NCE_Ecrire ws, ligne, "SourceInternet", ""
    NCE_Ecrire ws, ligne, "MethodeMiseAJour", CStr(d("MethodeMiseAJour"))
    NCE_Ecrire ws, ligne, "ID_Structure", CStr(d("ID_Structure"))

End Sub

'===============================================================================
' ECRITURE NOUVEAU SPECIALISTE
'===============================================================================
Private Sub NCE_AjouterNouveauSpecialiste( _
    ByVal wsSpe As Object, _
    ByVal wsSPT As Object, _
    ByVal d As Object)

    Dim colID As Long
    Dim ligne As Long
    Dim nouvelID As String
    Dim cleGroupe As String

    Dim types As Variant
    Dim i As Long
    Dim typeExamen As String
    Dim prioriteType As Long
    Dim prioriteFiche As Long

    colID = _
        NCE_TrouverColonneEntete(wsSpe, "ID")

    If colID = 0 Then
        Err.Raise vbObjectError + 721, , "Colonne ID absente de Specialistes."
    End If

    nouvelID = _
        NCE_ProchainID(wsSpe, colID, "SPE_", 4)

    cleGroupe = _
        "AUTO_" & nouvelID

    types = _
        Split(CStr(d("TypesExamensPossibles")), ";")

    prioriteFiche = 100

    For i = LBound(types) To UBound(types)

        typeExamen = Trim$(CStr(types(i)))

        If typeExamen <> "" Then

            prioriteType = _
                NCE_PrioritePourNouveauType( _
                    wsSPT, _
                    typeExamen)

            If prioriteType < prioriteFiche Then
                prioriteFiche = prioriteType
            End If

        End If

    Next i

    ligne = _
        wsSpe.Cells(wsSpe.Rows.Count, colID).End(NCE_XL_UP).Row + 1

    NCE_Ecrire wsSpe, ligne, "ID", nouvelID
    NCE_Ecrire wsSpe, ligne, "TypeExamen", CStr(d("TypeExamenPrincipal"))
    NCE_Ecrire wsSpe, ligne, "TypesExamensPossibles", CStr(d("TypesExamensPossibles"))
    NCE_Ecrire wsSpe, ligne, "NomDestinataire", CStr(d("NomDestinataire"))
    NCE_Ecrire wsSpe, ligne, "Titre", CStr(d("Titre"))
    NCE_Ecrire wsSpe, ligne, "Civilite", CStr(d("Civilite"))
    NCE_Ecrire wsSpe, ligne, "PrenomOuInitiale", CStr(d("Prenom"))
    NCE_Ecrire wsSpe, ligne, "Nom", CStr(d("Nom"))
    NCE_Ecrire wsSpe, ligne, "Structure", CStr(d("NomStructure"))
    NCE_Ecrire wsSpe, ligne, "Adresse1", CStr(d("Adresse1"))
    NCE_Ecrire wsSpe, ligne, "Adresse2", CStr(d("Adresse2"))
    NCE_Ecrire wsSpe, ligne, "CodePostal", CStr(d("CodePostal"))
    NCE_Ecrire wsSpe, ligne, "Ville", CStr(d("Ville"))
    NCE_Ecrire wsSpe, ligne, "Telephone", CStr(d("Telephone"))
    NCE_Ecrire wsSpe, ligne, "BlocDestinataireComplet", CStr(d("BlocDestinataireComplet"))
    NCE_Ecrire wsSpe, ligne, "FormuleAppel", CStr(d("FormuleAppel"))
    NCE_Ecrire wsSpe, ligne, "TutoiementVouvoiement", CStr(d("TutoiementVouvoiement"))
    NCE_Ecrire wsSpe, ligne, "Priorite", CStr(prioriteFiche)
    NCE_Ecrire wsSpe, ligne, "Actif", "Oui"
    NCE_Ecrire wsSpe, ligne, "AValider", "Non"
    NCE_Ecrire wsSpe, ligne, "CommentaireExtraction", CStr(d("CommentaireCreation"))
    NCE_Ecrire wsSpe, ligne, "ID_Structure", CStr(d("ID_Structure"))

    For i = LBound(types) To UBound(types)

        typeExamen = Trim$(CStr(types(i)))

        If typeExamen <> "" Then

            prioriteType = _
                NCE_PrioritePourNouveauType( _
                    wsSPT, _
                    typeExamen)

            NCE_AjouterLigneSPT _
                wsSPT, _
                nouvelID, _
                typeExamen, _
                CStr(d("NomDestinataire")), _
                CStr(d("BlocStructureComplet")), _
                CStr(d("CodePostal")), _
                CStr(d("Ville")), _
                CStr(d("BlocDestinataireComplet")), _
                CStr(d("FormuleAppel")), _
                prioriteType, _
                cleGroupe, _
                CStr(d("FormulePolitesse")), _
                CStr(d("TutoiementVouvoiement"))

        End If

    Next i

End Sub

'===============================================================================
' AJOUT DE TYPES A UN SPECIALISTE EXISTANT
'===============================================================================
Private Sub NCE_AjouterTypesSpecialisteExistant( _
    ByVal wsSpe As Object, _
    ByVal wsSPT As Object, _
    ByVal ligneSpe As Long, _
    ByVal idSpecialiste As String, _
    ByVal typesNouveaux As String)

    Dim nomDest As String
    Dim structure As String
    Dim adresse1 As String
    Dim adresse2 As String
    Dim cp As String
    Dim ville As String
    Dim telephone As String
    Dim blocDest As String
    Dim formuleAppel As String
    Dim relation As String
    Dim formulePolitesse As String
    Dim structureComplete As String
    Dim cleGroupe As String

    Dim types As Variant
    Dim i As Long
    Dim typeExamen As String
    Dim prioriteType As Long

    Dim typesPossiblesActuels As String
    Dim typesPossiblesMaj As String
    Dim commentaire As String

    nomDest = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "NomDestinataire")

    structure = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "Structure")

    adresse1 = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "Adresse1")

    adresse2 = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "Adresse2")

    cp = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "CodePostal")

    ville = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "Ville")

    telephone = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "Telephone")

    blocDest = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "BlocDestinataireComplet")

    formuleAppel = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "FormuleAppel")

    relation = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "TutoiementVouvoiement")

    formulePolitesse = _
        NCE_FormulePolitesseDepuisRelation(relation)

    structureComplete = _
        NCE_ConstruireBlocStructure( _
            structure, _
            adresse1, _
            adresse2, _
            cp, _
            ville, _
            telephone)

    cleGroupe = _
        NCE_CleGroupePourSpecialiste( _
            wsSPT, _
            idSpecialiste)

    types = _
        Split(typesNouveaux, ";")

    For i = LBound(types) To UBound(types)

        typeExamen = _
            Trim$(CStr(types(i)))

        If typeExamen <> "" Then

            prioriteType = _
                NCE_PrioritePourNouveauType( _
                    wsSPT, _
                    typeExamen)

            NCE_AjouterLigneSPT _
                wsSPT, _
                idSpecialiste, _
                typeExamen, _
                nomDest, _
                structureComplete, _
                cp, _
                ville, _
                blocDest, _
                formuleAppel, _
                prioriteType, _
                cleGroupe, _
                formulePolitesse, _
                relation

        End If

    Next i

    typesPossiblesActuels = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "TypesExamensPossibles")

    typesPossiblesMaj = _
        NCE_UnionTypes( _
            typesPossiblesActuels, _
            typesNouveaux)

    NCE_Ecrire _
        wsSpe, _
        ligneSpe, _
        "TypesExamensPossibles", _
        typesPossiblesMaj

    If Trim$(NCE_LireCelluleParEntete( _
        wsSpe, _
        ligneSpe, _
        "TypeExamen")) = "" Then

        NCE_Ecrire _
            wsSpe, _
            ligneSpe, _
            "TypeExamen", _
            NCE_PremierType(typesNouveaux)

    End If

    commentaire = _
        NCE_LireCelluleParEntete( _
            wsSpe, _
            ligneSpe, _
            "CommentaireExtraction")

    If Trim$(commentaire) <> "" Then
        commentaire = commentaire & " | "
    End If

    commentaire = _
        commentaire & _
        "Type(s) ajouté(s) via formulaire Word le " & _
        Format$(Date, "dd/mm/yyyy") & _
        " : " & _
        typesNouveaux

    NCE_Ecrire _
        wsSpe, _
        ligneSpe, _
        "CommentaireExtraction", _
        commentaire

End Sub

Private Sub NCE_AjouterLigneSPT( _
    ByVal ws As Object, _
    ByVal idSpecialiste As String, _
    ByVal typeExamen As String, _
    ByVal nomDestinataire As String, _
    ByVal blocStructure As String, _
    ByVal codePostal As String, _
    ByVal ville As String, _
    ByVal blocDestinataire As String, _
    ByVal formuleAppel As String, _
    ByVal priorite As Long, _
    ByVal cleGroupe As String, _
    ByVal formulePolitesse As String, _
    ByVal relation As String)

    Dim colID As Long
    Dim ligne As Long
    Dim nouvelID As String

    colID = _
        NCE_TrouverColonneEntete( _
            ws, _
            "ID_Ligne")

    If colID = 0 Then
        Err.Raise vbObjectError + 722, , "Colonne ID_Ligne absente de Specialistes_ParType."
    End If

    nouvelID = _
        NCE_ProchainID( _
            ws, _
            colID, _
            "SPT_", _
            4)

    ligne = _
        ws.Cells(ws.Rows.Count, colID).End(NCE_XL_UP).Row + 1

    NCE_Ecrire ws, ligne, "ID_Ligne", nouvelID
    NCE_Ecrire ws, ligne, "ID_Specialiste", idSpecialiste
    NCE_Ecrire ws, ligne, "TypeExamen", typeExamen
    NCE_Ecrire ws, ligne, "NomDestinataire", nomDestinataire
    NCE_Ecrire ws, ligne, "Structure", blocStructure
    NCE_Ecrire ws, ligne, "CodePostal", codePostal
    NCE_Ecrire ws, ligne, "Ville", ville
    NCE_Ecrire ws, ligne, "BlocDestinataireComplet", blocDestinataire
    NCE_Ecrire ws, ligne, "FormuleAppel", formuleAppel
    NCE_Ecrire ws, ligne, "Priorite", CStr(priorite)
    NCE_Ecrire ws, ligne, "Actif", "Oui"
    NCE_Ecrire ws, ligne, "AValider", "Non"
    NCE_Ecrire ws, ligne, "CleRegroupement", cleGroupe
    NCE_Ecrire ws, ligne, "FormulePolitesse", formulePolitesse
    NCE_Ecrire ws, ligne, "TutoiementVouvoiement", relation

End Sub

'===============================================================================
' DOUBLONS
'===============================================================================
Private Function NCE_TrouverLigneExacte( _
    ByVal ws As Object, _
    ByVal prenom As String, _
    ByVal nom As String, _
    ByVal enteteAffichage As String) As Long

    Dim colID As Long
    Dim colPrenom As Long
    Dim colNom As Long
    Dim colAffichage As Long
    Dim colActif As Long

    Dim derniereLigne As Long
    Dim i As Long

    Dim prenomRecherche As String
    Dim nomRecherche As String
    Dim identite1 As String
    Dim identite2 As String

    Dim prenomLigne As String
    Dim nomLigne As String
    Dim affichage As String
    Dim actif As String

    NCE_TrouverLigneExacte = 0

    prenomRecherche = _
        NCE_Normaliser(prenom)

    nomRecherche = _
        NCE_Normaliser(nom)

    If nomRecherche = "" Then Exit Function

    identite1 = _
        Trim$(prenomRecherche & " " & nomRecherche)

    identite2 = _
        Trim$(nomRecherche & " " & prenomRecherche)

    colID = NCE_TrouverColonneEntete(ws, "ID")
    colPrenom = NCE_TrouverColonneEntete(ws, "PrenomOuInitiale")
    colNom = NCE_TrouverColonneEntete(ws, "Nom")
    colAffichage = NCE_TrouverColonneEntete(ws, enteteAffichage)
    colActif = NCE_TrouverColonneEntete(ws, "Actif")

    If colID = 0 _
    Or colNom = 0 _
    Or colAffichage = 0 Then
        Exit Function
    End If

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colID).End(NCE_XL_UP).Row

    For i = 2 To derniereLigne

        If colActif > 0 Then

            actif = _
                NCE_LireCellule(ws, i, colActif)

            If Not NCE_EstOui(actif) Then
                GoTo LigneSuivante
            End If

        End If

        prenomLigne = _
            NCE_Normaliser(NCE_LireCellule(ws, i, colPrenom))

        nomLigne = _
            NCE_Normaliser(NCE_LireCellule(ws, i, colNom))

        affichage = _
            NCE_NormaliserIdentite( _
                NCE_LireCellule(ws, i, colAffichage))

        If prenomRecherche <> "" _
        And nomLigne = nomRecherche _
        And prenomLigne = prenomRecherche Then

            NCE_TrouverLigneExacte = i
            Exit Function

        End If

        If prenomRecherche <> "" _
        And (affichage = identite1 _
             Or affichage = identite2) Then

            NCE_TrouverLigneExacte = i
            Exit Function

        End If

LigneSuivante:

    Next i

End Function

Private Function NCE_ListerDoublonsProbables( _
    ByVal wsGen As Object, _
    ByVal wsSpe As Object, _
    ByVal prenom As String, _
    ByVal nom As String, _
    ByVal telephone As String, _
    ByVal idAIgnorer As String) As String

    Dim resultat As String

    resultat = _
        NCE_ListerDoublonsProbablesFeuille( _
            wsGen, _
            "Généraliste", _
            "NomAffichage", _
            prenom, _
            nom, _
            telephone, _
            idAIgnorer)

    resultat = _
        resultat & _
        NCE_ListerDoublonsProbablesFeuille( _
            wsSpe, _
            "Spécialiste", _
            "NomDestinataire", _
            prenom, _
            nom, _
            telephone, _
            idAIgnorer)

    NCE_ListerDoublonsProbables = resultat

End Function

Private Function NCE_ListerDoublonsProbablesFeuille( _
    ByVal ws As Object, _
    ByVal typeLibelle As String, _
    ByVal enteteAffichage As String, _
    ByVal prenom As String, _
    ByVal nom As String, _
    ByVal telephone As String, _
    ByVal idAIgnorer As String) As String

    Dim colID As Long
    Dim colNom As Long
    Dim colPrenom As Long
    Dim colAffichage As Long
    Dim colStructure As Long
    Dim colVille As Long
    Dim colTelephone As Long
    Dim colActif As Long

    Dim derniereLigne As Long
    Dim i As Long

    Dim idLigne As String
    Dim nomRecherche As String
    Dim prenomRecherche As String
    Dim telRecherche As String

    Dim nomLigne As String
    Dim prenomLigne As String
    Dim telLigne As String
    Dim actif As String

    Dim probable As Boolean
    Dim texte As String

    nomRecherche = NCE_Normaliser(nom)
    prenomRecherche = NCE_Normaliser(prenom)
    telRecherche = NCE_GarderChiffres(telephone)

    colID = NCE_TrouverColonneEntete(ws, "ID")
    colNom = NCE_TrouverColonneEntete(ws, "Nom")
    colPrenom = NCE_TrouverColonneEntete(ws, "PrenomOuInitiale")
    colAffichage = NCE_TrouverColonneEntete(ws, enteteAffichage)
    colStructure = NCE_TrouverColonneEntete(ws, "Structure")
    colVille = NCE_TrouverColonneEntete(ws, "Ville")
    colTelephone = NCE_TrouverColonneEntete(ws, "Telephone")
    colActif = NCE_TrouverColonneEntete(ws, "Actif")

    If colID = 0 Or colNom = 0 Then Exit Function

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colID).End(NCE_XL_UP).Row

    For i = 2 To derniereLigne

        idLigne = _
            Trim$(NCE_LireCellule(ws, i, colID))

        If idAIgnorer <> "" _
        And StrComp(idLigne, idAIgnorer, vbTextCompare) = 0 Then
            GoTo LigneSuivante
        End If

        If colActif > 0 Then

            actif = _
                NCE_LireCellule(ws, i, colActif)

            If Not NCE_EstOui(actif) Then
                GoTo LigneSuivante
            End If

        End If

        nomLigne = _
            NCE_Normaliser(NCE_LireCellule(ws, i, colNom))

        prenomLigne = _
            NCE_Normaliser(NCE_LireCellule(ws, i, colPrenom))

        telLigne = _
            NCE_GarderChiffres(NCE_LireCellule(ws, i, colTelephone))

        probable = False

        If nomRecherche <> "" _
        And nomLigne = nomRecherche _
        And prenomLigne <> prenomRecherche Then

            probable = True

        ElseIf telRecherche <> "" _
        And telLigne <> "" _
        And telRecherche = telLigne Then

            probable = True

        End If

        If probable Then

            texte = _
                texte & _
                typeLibelle & " — " & _
                idLigne & " — " & _
                NCE_LireCellule(ws, i, colAffichage)

            If colStructure > 0 _
            And Trim$(NCE_LireCellule(ws, i, colStructure)) <> "" Then

                texte = _
                    texte & " — " & _
                    NCE_LireCellule(ws, i, colStructure)

            End If

            If colVille > 0 _
            And Trim$(NCE_LireCellule(ws, i, colVille)) <> "" Then

                texte = _
                    texte & " — " & _
                    NCE_LireCellule(ws, i, colVille)

            End If

            texte = texte & vbCrLf

        End If

LigneSuivante:

    Next i

    NCE_ListerDoublonsProbablesFeuille = texte

End Function

'===============================================================================
' TYPES D'EXAMENS
'===============================================================================
Private Function NCE_TypesSelectionnes(ByVal frm As Object) As String

    Dim lst As Object
    Dim i As Long
    Dim valeur As String
    Dim resultat As String

    Set lst = frm.Controls("lstTypesExamens")

    For i = 0 To lst.ListCount - 1

        If lst.Selected(i) Then

            valeur = _
                Trim$(CStr(lst.List(i)))

            If valeur <> "" _
            And StrComp(valeur, "A_COMPLETER", vbTextCompare) <> 0 Then

                resultat = _
                    NCE_AjouterTypeUnique( _
                        resultat, _
                        valeur)

            End If

        End If

    Next i

    NCE_TypesSelectionnes = resultat

End Function

Private Function NCE_TypesNonAssocies( _
    ByVal wsSPT As Object, _
    ByVal idSpecialiste As String, _
    ByVal typesSelectionnes As String) As String

    Dim types As Variant
    Dim i As Long
    Dim typeExamen As String
    Dim resultat As String

    types = _
        Split(typesSelectionnes, ";")

    For i = LBound(types) To UBound(types)

        typeExamen = _
            Trim$(CStr(types(i)))

        If typeExamen <> "" Then

            If Not NCE_AssociationExiste( _
                wsSPT, _
                idSpecialiste, _
                typeExamen) Then

                resultat = _
                    NCE_AjouterTypeUnique( _
                        resultat, _
                        typeExamen)

            End If

        End If

    Next i

    NCE_TypesNonAssocies = resultat

End Function

Private Function NCE_AssociationExiste( _
    ByVal ws As Object, _
    ByVal idSpecialiste As String, _
    ByVal typeExamen As String) As Boolean

    Dim colID As Long
    Dim colType As Long
    Dim colActif As Long

    Dim derniereLigne As Long
    Dim i As Long

    NCE_AssociationExiste = False

    colID = _
        NCE_TrouverColonneEntete( _
            ws, _
            "ID_Specialiste")

    colType = _
        NCE_TrouverColonneEntete( _
            ws, _
            "TypeExamen")

    colActif = _
        NCE_TrouverColonneEntete( _
            ws, _
            "Actif")

    If colID = 0 Or colType = 0 Then Exit Function

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colID).End(NCE_XL_UP).Row

    For i = 2 To derniereLigne

        If StrComp( _
            Trim$(NCE_LireCellule(ws, i, colID)), _
            idSpecialiste, _
            vbTextCompare) = 0 _
        And StrComp( _
            Trim$(NCE_LireCellule(ws, i, colType)), _
            typeExamen, _
            vbTextCompare) = 0 Then

            If colActif = 0 _
            Or NCE_EstOui(NCE_LireCellule(ws, i, colActif)) Then

                NCE_AssociationExiste = True
                Exit Function

            End If

        End If

    Next i

End Function

Private Function NCE_PrioritePourNouveauType( _
    ByVal ws As Object, _
    ByVal typeExamen As String) As Long

    Dim colType As Long
    Dim colActif As Long
    Dim derniereLigne As Long
    Dim i As Long

    NCE_PrioritePourNouveauType = 1

    colType = NCE_TrouverColonneEntete(ws, "TypeExamen")
    colActif = NCE_TrouverColonneEntete(ws, "Actif")

    If colType = 0 Then Exit Function

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colType).End(NCE_XL_UP).Row

    For i = 2 To derniereLigne

        If StrComp( _
            Trim$(NCE_LireCellule(ws, i, colType)), _
            typeExamen, _
            vbTextCompare) = 0 Then

            If colActif = 0 _
            Or NCE_EstOui(NCE_LireCellule(ws, i, colActif)) Then

                NCE_PrioritePourNouveauType = 100
                Exit Function

            End If

        End If

    Next i

End Function

Private Function NCE_CleGroupePourSpecialiste( _
    ByVal ws As Object, _
    ByVal idSpecialiste As String) As String

    Dim colID As Long
    Dim colCle As Long
    Dim derniereLigne As Long
    Dim i As Long

    Dim cle As String
    Dim dict As Object
    Dim cles As Variant

    Set dict = CreateObject("Scripting.Dictionary")

    colID = NCE_TrouverColonneEntete(ws, "ID_Specialiste")
    colCle = NCE_TrouverColonneEntete(ws, "CleRegroupement")

    If colID = 0 Or colCle = 0 Then
        NCE_CleGroupePourSpecialiste = "AUTO_" & idSpecialiste
        Exit Function
    End If

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colID).End(NCE_XL_UP).Row

    For i = 2 To derniereLigne

        If StrComp( _
            Trim$(NCE_LireCellule(ws, i, colID)), _
            idSpecialiste, _
            vbTextCompare) = 0 Then

            cle = _
                UCase$(Trim$(NCE_LireCellule(ws, i, colCle)))

            If cle <> "" Then
                If Not dict.Exists(cle) Then
                    dict.Add cle, True
                End If
            End If

        End If

    Next i

    If dict.Count = 1 Then

        cles = dict.Keys

        NCE_CleGroupePourSpecialiste = _
            CStr(cles(LBound(cles)))

    Else

        NCE_CleGroupePourSpecialiste = _
            "AUTO_" & idSpecialiste

    End If

End Function

Private Function NCE_UnionTypes( _
    ByVal liste1 As String, _
    ByVal liste2 As String) As String

    Dim resultat As String
    Dim a As Variant
    Dim i As Long
    Dim t As String

    resultat = ""

    If Trim$(liste1) <> "" Then

        a = Split(liste1, ";")

        For i = LBound(a) To UBound(a)

            t = Trim$(CStr(a(i)))

            If t <> "" Then
                resultat = NCE_AjouterTypeUnique(resultat, t)
            End If

        Next i

    End If

    If Trim$(liste2) <> "" Then

        a = Split(liste2, ";")

        For i = LBound(a) To UBound(a)

            t = Trim$(CStr(a(i)))

            If t <> "" Then
                resultat = NCE_AjouterTypeUnique(resultat, t)
            End If

        Next i

    End If

    NCE_UnionTypes = resultat

End Function

Private Function NCE_AjouterTypeUnique( _
    ByVal liste As String, _
    ByVal nouveauType As String) As String

    If Trim$(nouveauType) = "" Then
        NCE_AjouterTypeUnique = liste
        Exit Function
    End If

    If Trim$(liste) = "" Then

        NCE_AjouterTypeUnique = nouveauType

    ElseIf InStr( _
        1, _
        ";" & liste & ";", _
        ";" & nouveauType & ";", _
        vbTextCompare) = 0 Then

        NCE_AjouterTypeUnique = _
            liste & ";" & nouveauType

    Else

        NCE_AjouterTypeUnique = liste

    End If

End Function

Private Function NCE_PremierType(ByVal liste As String) As String

    Dim p As Long

    p = InStr(1, liste, ";", vbBinaryCompare)

    If p > 0 Then
        NCE_PremierType = Trim$(Left$(liste, p - 1))
    Else
        NCE_PremierType = Trim$(liste)
    End If

End Function

Private Function NCE_TypesEnLignes(ByVal liste As String) As String

    Dim a As Variant
    Dim i As Long
    Dim resultat As String

    a = Split(liste, ";")

    For i = LBound(a) To UBound(a)

        If Trim$(CStr(a(i))) <> "" Then
            resultat = resultat & _
                "• " & Trim$(CStr(a(i))) & vbCrLf
        End If

    Next i

    NCE_TypesEnLignes = resultat

End Function

'===============================================================================
' RESUME / BLOCS
'===============================================================================
Private Function NCE_ResumeNouvelleFiche( _
    ByVal typeCorrespondant As String, _
    ByVal d As Object) As String

    Dim texte As String

    texte = _
        "Nouvelle fiche " & typeCorrespondant & vbCrLf & vbCrLf & _
        CStr(d("NomAffichage")) & vbCrLf

    If Trim$(CStr(d("NomStructure"))) <> "" Then
        texte = texte & CStr(d("NomStructure")) & vbCrLf
    End If

    If Trim$(CStr(d("Adresse1"))) <> "" Then
        texte = texte & CStr(d("Adresse1")) & vbCrLf
    End If

    If Trim$(CStr(d("Adresse2"))) <> "" Then
        texte = texte & CStr(d("Adresse2")) & vbCrLf
    End If

    texte = texte & _
        Trim$(CStr(d("CodePostal")) & " " & CStr(d("Ville"))) & vbCrLf

    If Trim$(CStr(d("Telephone"))) <> "" Then
        texte = texte & _
            "Tél. : " & CStr(d("Telephone")) & vbCrLf
    End If

    texte = texte & _
        vbCrLf & _
        "Formule d'appel : " & CStr(d("FormuleAppel")) & vbCrLf & _
        "Relation : " & CStr(d("TutoiementVouvoiement")) & vbCrLf

    If StrComp(typeCorrespondant, "Spécialiste", vbTextCompare) = 0 Then

        texte = texte & _
            vbCrLf & _
            "Type(s) :" & vbCrLf & _
            NCE_TypesEnLignes(CStr(d("TypesExamensPossibles")))

    End If

    NCE_ResumeNouvelleFiche = texte

End Function

Private Function NCE_ConstruireBlocStructure( _
    ByVal structure As String, _
    ByVal adresse1 As String, _
    ByVal adresse2 As String, _
    ByVal codePostal As String, _
    ByVal ville As String, _
    ByVal telephone As String) As String

    Dim texte As String

    If Trim$(structure) <> "" Then
        texte = Trim$(structure)
    End If

    If Trim$(adresse1) <> "" Then
        NCE_AjouterLigne texte, adresse1
    End If

    If Trim$(adresse2) <> "" Then
        NCE_AjouterLigne texte, adresse2
    End If

    If Trim$(codePostal & " " & ville) <> "" Then
        NCE_AjouterLigne texte, Trim$(codePostal & " " & ville)
    End If

    If Trim$(telephone) <> "" Then
        NCE_AjouterLigne texte, telephone
    End If

    NCE_ConstruireBlocStructure = texte

End Function

Private Sub NCE_AjouterLigne( _
    ByRef texte As String, _
    ByVal ligne As String)

    If Trim$(ligne) = "" Then Exit Sub

    If Trim$(texte) = "" Then
        texte = Trim$(ligne)
    Else
        texte = texte & vbLf & Trim$(ligne)
    End If

End Sub

Private Function NCE_FormulePolitesseDepuisRelation( _
    ByVal relation As String) As String

    If StrComp(Trim$(relation), "tu", vbTextCompare) = 0 Then
        NCE_FormulePolitesseDepuisRelation = "Amitiés."
    Else
        NCE_FormulePolitesseDepuisRelation = "Bien cordialement."
    End If

End Function

'===============================================================================
' SAUVEGARDE / IDS / ECRITURE PAR ENTETE
'===============================================================================
Private Function NCE_CreerSauvegardeBase() As String

    Dim chemin As String
    Dim dossier As String
    Dim nomBase As String
    Dim nomSansExtension As String
    Dim p As Long
    Dim sauvegarde As String

    chemin = CHEMIN_BASE_CORRESPONDANTS

    If Dir$(chemin) = "" Then
        Err.Raise vbObjectError + 723, , "Base correspondants introuvable."
    End If

    dossier = _
        Left$(chemin, InStrRev(chemin, "\"))

    nomBase = _
        Mid$(chemin, InStrRev(chemin, "\") + 1)

    p = InStrRev(nomBase, ".")

    If p > 0 Then
        nomSansExtension = Left$(nomBase, p - 1)
    Else
        nomSansExtension = nomBase
    End If

    sauvegarde = _
        dossier & _
        nomSansExtension & _
        "_AVANT_AJOUT_CORRESP_" & _
        Format$(Now, "yyyymmdd_hhnnss") & "_" & modFichiers.IdUnique() & _
        ".xlsx"

    FileCopy chemin, sauvegarde

    NCE_CreerSauvegardeBase = sauvegarde

End Function

Private Function NCE_ProchainID( _
    ByVal ws As Object, _
    ByVal colID As Long, _
    ByVal prefixe As String, _
    ByVal nombreChiffres As Long) As String

    Dim derniereLigne As Long
    Dim i As Long
    Dim valeur As String
    Dim suffixe As String
    Dim numero As Long
    Dim maximum As Long

    maximum = 0

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colID).End(NCE_XL_UP).Row

    For i = 2 To derniereLigne

        valeur = _
            UCase$(Trim$(NCE_LireCellule(ws, i, colID)))

        If Left$(valeur, Len(prefixe)) = UCase$(prefixe) Then

            suffixe = _
                Mid$(valeur, Len(prefixe) + 1)

            If IsNumeric(suffixe) Then

                numero = CLng(suffixe)

                If numero > maximum Then
                    maximum = numero
                End If

            End If

        End If

    Next i

    NCE_ProchainID = _
        prefixe & _
        Format$(maximum + 1, String$(nombreChiffres, "0"))

End Function

Private Sub NCE_Ecrire( _
    ByVal ws As Object, _
    ByVal ligne As Long, _
    ByVal entete As String, _
    ByVal valeur As String)

    Dim col As Long

    col = _
        NCE_TrouverColonneEntete(ws, entete)

    If col = 0 Then

        Err.Raise _
            vbObjectError + 724, _
            , _
            "Colonne absente : " & ws.Name & "!" & entete

    End If

    modFichiers.EcrireCelluleTexte ws.Cells(ligne, col), valeur

End Sub

'===============================================================================
' UTILITAIRES EXCEL
'===============================================================================
Private Function NCE_IDStructureSelectionnee( _
    ByVal frm As Object) As String

    Dim cmb As Object

    Set cmb = frm.Controls("cmbStructure")

    NCE_IDStructureSelectionnee = ""

    If cmb.ListIndex < 0 Then Exit Function

    NCE_IDStructureSelectionnee = _
        Trim$(CStr(cmb.List(cmb.ListIndex, 2)))

End Function

Private Function NCE_TrouverColonneEntete( _
    ByVal ws As Object, _
    ByVal entete As String) As Long

    Dim derniereColonne As Long
    Dim c As Long
    Dim valeur As String

    derniereColonne = _
        ws.Cells(1, ws.Columns.Count).End(NCE_XL_TO_LEFT).Column

    For c = 1 To derniereColonne

        valeur = _
            Trim$(CStr(ws.Cells(1, c).Value))

        If StrComp(valeur, entete, vbTextCompare) = 0 Then

            NCE_TrouverColonneEntete = c
            Exit Function

        End If

    Next c

End Function

Private Function NCE_LireCellule( _
    ByVal ws As Object, _
    ByVal ligne As Long, _
    ByVal colonne As Long) As String

    Dim v As Variant

    If colonne <= 0 Then Exit Function

    v = ws.Cells(ligne, colonne).Value

    If IsError(v) Then Err.Raise vbObjectError + 728, , "Erreur Excel dans la base correspondants."
    If IsNull(v) Or IsEmpty(v) Then Exit Function

    NCE_LireCellule = CStr(v)

End Function

Private Function NCE_LireCelluleParEntete( _
    ByVal ws As Object, _
    ByVal ligne As Long, _
    ByVal entete As String) As String

    Dim col As Long

    col = _
        NCE_TrouverColonneEntete(ws, entete)

    If col = 0 Then Exit Function

    NCE_LireCelluleParEntete = _
        NCE_LireCellule(ws, ligne, col)

End Function

Private Function NCE_LireRange(ByVal rng As Object) As String

    Dim v As Variant

    v = rng.Value

    If IsError(v) Then Err.Raise vbObjectError + 728, , "Erreur Excel dans la base correspondants."
    If IsNull(v) Or IsEmpty(v) Then Exit Function

    NCE_LireRange = CStr(v)

End Function

Private Function NCE_EstOui(ByVal valeur As String) As Boolean

    Dim t As String

    t = LCase$(Trim$(valeur))

    NCE_EstOui = _
        (t = "oui" _
         Or t = "true" _
         Or t = "vrai" _
         Or t = "1")

End Function

'===============================================================================
' NORMALISATION
'===============================================================================
Private Function NCE_Normaliser(ByVal texte As String) As String

    texte = LCase$(Trim$(texte))

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
    texte = Replace(texte, "ù", "u")
    texte = Replace(texte, "û", "u")
    texte = Replace(texte, "ü", "u")
    texte = Replace(texte, "ú", "u")
    texte = Replace(texte, "ÿ", "y")
    texte = Replace(texte, "œ", "oe")
    texte = Replace(texte, "æ", "ae")

    texte = Replace(texte, Chr(160), " ")
    texte = Replace(texte, vbTab, " ")
    texte = Replace(texte, vbCr, " ")
    texte = Replace(texte, vbLf, " ")
    texte = Replace(texte, "’", "'")
    texte = Replace(texte, "'", " ")
    texte = Replace(texte, "-", " ")
    texte = Replace(texte, ".", " ")
    texte = Replace(texte, ",", " ")
    texte = Replace(texte, ";", " ")
    texte = Replace(texte, ":", " ")
    texte = Replace(texte, "/", " ")
    texte = Replace(texte, "\", " ")
    texte = Replace(texte, "(", " ")
    texte = Replace(texte, ")", " ")

    Do While InStr(texte, "  ") > 0
        texte = Replace(texte, "  ", " ")
    Loop

    NCE_Normaliser = Trim$(texte)

End Function

Private Function NCE_NormaliserIdentite(ByVal texte As String) As String

    Dim t As String

    t = " " & NCE_Normaliser(texte) & " "

    t = Replace(t, " monsieur ", " ")
    t = Replace(t, " madame ", " ")
    t = Replace(t, " docteur ", " ")
    t = Replace(t, " dr ", " ")

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    NCE_NormaliserIdentite = Trim$(t)

End Function

Private Function NCE_GarderChiffres(ByVal texte As String) As String

    Dim i As Long
    Dim ch As String
    Dim resultat As String

    For i = 1 To Len(texte)

        ch = Mid$(texte, i, 1)

        If ch >= "0" And ch <= "9" Then
            resultat = resultat & ch
        End If

    Next i

    NCE_GarderChiffres = resultat

End Function
