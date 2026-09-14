Attribute VB_Name = "modSortieDragon"
Option Explicit

'===============================================================================
' MODULE : modSortieDragon
' VERSION : DOMICILE-GRAS-1Z / SORTIE Z:
'
' Enregistre le document final dans \\DS224\home\sortiedragon sous la forme :
'   NOM Prenom aammddhhmm.docx
'
' Particularité :
' certains anciens courriers mémorisent l'identité ainsi :
'   gPatient.Nom = "COMPAGNET Carole"
'   gPatient.Prenom = ""
'
' Dans ce cas, le module sépare localement le nom et le prénom uniquement
' pour construire le nom du fichier. Le mécanisme d'anonymisation n'est pas
' modifié.
'===============================================================================



Public Function SD_EnregistrerCourrierFinal( _
    ByVal doc As Document, _
    ByRef cheminEnregistre As String) As Boolean

    Dim nomPatient As String
    Dim prenomPatient As String
    Dim baseNom As String
    Dim cheminCandidat As String
    Dim compteur As Long
    Dim anciensAlerts As WdAlertLevel

    Dim numeroErreur As Long
    Dim descriptionErreur As String
    Dim etapeErreur As String
    Dim nomDocument As String

    On Error GoTo GestionErreur

    SD_EnregistrerCourrierFinal = False
    cheminEnregistre = ""
    etapeErreur = "Initialisation"

    anciensAlerts = Application.DisplayAlerts

    etapeErreur = "Vérification du document"

    If doc Is Nothing Then
        Err.Raise vbObjectError + 810, , _
            "Le document final n'est plus disponible."
    End If

    nomDocument = doc.Name

    etapeErreur = "Lecture du nom et du prénom du patient"

    SD_ObtenirNomPrenomPourFichier _
        nomPatient, _
        prenomPatient

    nomPatient = SD_NettoyerPartieNomFichier(nomPatient)
    prenomPatient = SD_NettoyerPartieNomFichier(prenomPatient)

    If nomPatient = "" Or prenomPatient = "" Then
        Err.Raise vbObjectError + 811, , _
            "Le nom ou le prénom du patient n'a pas été identifié. " & _
            "Nom=[" & nomPatient & "] Prénom=[" & prenomPatient & "]"
    End If

    etapeErreur = "Création/vérification de \\DS224\home\sortiedragon"
    SD_CreerDossierSiNecessaire SD_DOSSIER_SORTIE

    etapeErreur = "Construction du nom de fichier"

    baseNom = _
        UCase$(nomPatient) & " " & _
        SD_MajusculeInitialesPrenom(prenomPatient) & " " & _
        Format$(Now, "yymmddhhnn")

    cheminCandidat = _
        SD_DOSSIER_SORTIE & "\" & _
        baseNom & ".docx"

    compteur = 2

    etapeErreur = "Recherche d'un éventuel doublon de nom"

    Do While SD_FichierExiste(cheminCandidat)

        cheminCandidat = _
            SD_DOSSIER_SORTIE & "\" & _
            baseNom & "_" & _
            CStr(compteur) & _
            ".docx"

        compteur = compteur + 1

    Loop

    etapeErreur = "SaveAs2 vers " & cheminCandidat

    Application.DisplayAlerts = wdAlertsNone

    doc.SaveAs2 _
        FileName:=cheminCandidat, _
        FileFormat:=wdFormatXMLDocument, _
        AddToRecentFiles:=False

    Application.DisplayAlerts = anciensAlerts

    cheminEnregistre = cheminCandidat
    SD_EnregistrerCourrierFinal = True

    Exit Function

GestionErreur:

    numeroErreur = Err.Number
    descriptionErreur = Err.description

    On Error Resume Next
    Application.DisplayAlerts = anciensAlerts
    On Error GoTo 0

    cheminEnregistre = ""

    MsgBox _
        "Le courrier a été corrigé, mais son enregistrement automatique " & _
        "dans \\DS224\home\sortiedragon a échoué." & vbCrLf & vbCrLf & _
        "Étape : " & etapeErreur & vbCrLf & _
        "Erreur : " & numeroErreur & " - " & descriptionErreur & vbCrLf & _
        "Document : " & nomDocument & vbCrLf & _
        "Destination : " & cheminCandidat, _
        vbExclamation, _
        "Enregistrement du courrier"

End Function

Private Sub SD_ObtenirNomPrenomPourFichier( _
    ByRef nomPatient As String, _
    ByRef prenomPatient As String)

    Dim identite As String

    nomPatient = Trim$(gPatient.nom)
    prenomPatient = Trim$(gPatient.prenom)

    'Cas normal : les deux champs ont été correctement séparés.
    If nomPatient <> "" And prenomPatient <> "" Then Exit Sub

    'Certains anciens courriers stockent tout le bloc "NOM Prenom"
    'dans gPatient.Nom. On tente d'abord cette valeur.
    identite = Trim$(gPatient.nom)

    'Sinon, utiliser NomComplet, qui est déjà la référence locale utilisée
    'par le mécanisme d'anonymisation/restauration.
    If identite = "" Then
        identite = Trim$(gPatient.NomComplet)
    End If

    SD_SeparerIdentite _
        identite, _
        nomPatient, _
        prenomPatient

End Sub

Private Sub SD_SeparerIdentite( _
    ByVal identite As String, _
    ByRef nomPatient As String, _
    ByRef prenomPatient As String)

    Dim morceaux() As String
    Dim i As Long
    Dim premierPrenom As Long
    Dim nomConstruit As String
    Dim prenomConstruit As String

    identite = SD_NormaliserEspaces(identite)

    If identite = "" Then Exit Sub

    'Retirer une éventuelle civilité si elle s'est glissée dans le bloc.
    identite = SD_RetirerCiviliteDebut(identite)
    identite = SD_NormaliserEspaces(identite)

    If InStr(1, identite, " ", vbBinaryCompare) = 0 Then
        nomPatient = identite
        prenomPatient = ""
        Exit Sub
    End If

    morceaux = Split(identite, " ")

    'Chercher le premier mot qui n'est pas entièrement en majuscules.
    'Exemples :
    '   COMPAGNET Carole        -> COMPAGNET / Carole
    '   LE GOFF Jean-Pierre     -> LE GOFF / Jean-Pierre
    premierPrenom = 0

    For i = LBound(morceaux) To UBound(morceaux)

        If Trim$(morceaux(i)) <> "" Then

            If Not SD_EstMotMajuscules(morceaux(i)) Then
                premierPrenom = i
                Exit For
            End If

        End If

    Next i

    'Si tout est en majuscules, la séparation est intrinsèquement ambiguë.
    'Le repli le plus sûr pour les courriers usuels est :
    'dernier mot = prénom ; tous les mots précédents = nom.
    If premierPrenom = 0 Then
        premierPrenom = UBound(morceaux)
    End If

    nomConstruit = ""
    prenomConstruit = ""

    For i = LBound(morceaux) To UBound(morceaux)

        If i < premierPrenom Then

            If nomConstruit <> "" Then nomConstruit = nomConstruit & " "
            nomConstruit = nomConstruit & morceaux(i)

        Else

            If prenomConstruit <> "" Then prenomConstruit = prenomConstruit & " "
            prenomConstruit = prenomConstruit & morceaux(i)

        End If

    Next i

    nomPatient = Trim$(nomConstruit)
    prenomPatient = Trim$(prenomConstruit)

End Sub

Private Function SD_EstMotMajuscules( _
    ByVal mot As String) As Boolean

    Dim i As Long
    Dim c As String
    Dim contientLettre As Boolean

    mot = Trim$(mot)
    contientLettre = False

    For i = 1 To Len(mot)

        c = Mid$(mot, i, 1)

        If LCase$(c) <> UCase$(c) Then
            contientLettre = True
        End If

    Next i

    If Not contientLettre Then
        SD_EstMotMajuscules = True
    Else
        SD_EstMotMajuscules = (mot = UCase$(mot))
    End If

End Function

Private Function SD_RetirerCiviliteDebut( _
    ByVal texte As String) As String

    Dim t As String

    t = Trim$(texte)

    If LCase$(Left$(t, 9)) = "monsieur " Then
        t = Mid$(t, 10)
    ElseIf LCase$(Left$(t, 7)) = "madame " Then
        t = Mid$(t, 8)
    ElseIf LCase$(Left$(t, 5)) = "m. " Then
        t = Mid$(t, 4)
    ElseIf LCase$(Left$(t, 4)) = "mr " Then
        t = Mid$(t, 4)
    ElseIf LCase$(Left$(t, 4)) = "mme " Then
        t = Mid$(t, 5)
    End If

    SD_RetirerCiviliteDebut = Trim$(t)

End Function

Private Function SD_NormaliserEspaces( _
    ByVal texte As String) As String

    texte = Replace(texte, vbCr, " ")
    texte = Replace(texte, vbLf, " ")
    texte = Replace(texte, vbTab, " ")
    texte = Replace(texte, Chr$(160), " ")

    Do While InStr(texte, "  ") > 0
        texte = Replace(texte, "  ", " ")
    Loop

    SD_NormaliserEspaces = Trim$(texte)

End Function

Private Sub SD_CreerDossierSiNecessaire( _
    ByVal dossier As String)

    Dim fso As Object
    Dim dossierParent As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(dossier) Then Exit Sub

    dossierParent = fso.GetParentFolderName(dossier)

    If dossierParent = "" _
    Or Not fso.FolderExists(dossierParent) Then

        Err.Raise _
            vbObjectError + 813, _
            "SD_CreerDossierSiNecessaire", _
            "Le partage NAS parent est inaccessible : " & dossierParent

    End If

    fso.CreateFolder dossier

    If Not fso.FolderExists(dossier) Then

        Err.Raise _
            vbObjectError + 814, _
            "SD_CreerDossierSiNecessaire", _
            "Le dossier de sortie NAS n'a pas pu être créé : " & dossier

    End If

End Sub

Public Sub SD_TesterAccesDossierNAS()

    Dim fso As Object
    Dim fichierTest As Object
    Dim cheminTest As String
    Dim numeroErreur As Long
    Dim descriptionErreur As String

    On Error GoTo GestionErreur

    SD_CreerDossierSiNecessaire SD_DOSSIER_SORTIE

    Set fso = CreateObject("Scripting.FileSystemObject")

    cheminTest = _
        SD_DOSSIER_SORTIE & _
        "\_test_ecriture_" & _
        Format$(Now, "yyyymmdd_hhnnss") & _
        ".tmp"

    Set fichierTest = _
        fso.CreateTextFile( _
            cheminTest, _
            True, _
            False)

    fichierTest.WriteLine _
        "Test d'écriture ModeleCourrierChatGPT " & _
        Format$(Now, "dd/mm/yyyy hh:nn:ss")

    fichierTest.Close
    Set fichierTest = Nothing

    If Not fso.FileExists(cheminTest) Then

        Err.Raise _
            vbObjectError + 815, _
            "SD_TesterAccesDossierNAS", _
            "Le fichier de test n'a pas été créé."

    End If

    fso.DeleteFile cheminTest, True

    MsgBox _
        "Accès en écriture au NAS validé :" & _
        vbCrLf & vbCrLf & _
        SD_DOSSIER_SORTIE, _
        vbInformation, _
        "Test du dossier de sortie"

    Exit Sub

GestionErreur:

    numeroErreur = Err.Number
    descriptionErreur = Err.description

    On Error Resume Next

    If Not fichierTest Is Nothing Then
        fichierTest.Close
    End If

    If Not fso Is Nothing Then
        If cheminTest <> "" Then
            If fso.FileExists(cheminTest) Then
                fso.DeleteFile cheminTest, True
            End If
        End If
    End If

    On Error GoTo 0

    MsgBox _
        "Échec de l'accès en écriture au NAS." & _
        vbCrLf & vbCrLf & _
        "Dossier : " & SD_DOSSIER_SORTIE & _
        vbCrLf & _
        "Erreur : " & numeroErreur & " - " & descriptionErreur, _
        vbExclamation, _
        "Test du dossier de sortie"

End Sub

Private Function SD_FichierExiste( _
    ByVal chemin As String) As Boolean

    SD_FichierExiste = _
        (Len(Dir$(chemin, _
            vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)

End Function

Private Function SD_NettoyerPartieNomFichier( _
    ByVal texte As String) As String

    texte = SD_NormaliserEspaces(texte)

    texte = Replace(texte, "\", "")
    texte = Replace(texte, "/", "")
    texte = Replace(texte, ":", "")
    texte = Replace(texte, "*", "")
    texte = Replace(texte, "?", "")
    texte = Replace(texte, """", "")
    texte = Replace(texte, "<", "")
    texte = Replace(texte, ">", "")
    texte = Replace(texte, "|", "")

    Do While Len(texte) > 0 _
    And (Right$(texte, 1) = "." _
         Or Right$(texte, 1) = " ")

        texte = Left$(texte, Len(texte) - 1)

    Loop

    SD_NettoyerPartieNomFichier = Trim$(texte)

End Function

Private Function SD_MajusculeInitialesPrenom( _
    ByVal texte As String) As String

    Dim resultat As String
    Dim i As Long
    Dim caractere As String
    Dim mettreMajuscule As Boolean

    texte = LCase$(Trim$(texte))

    resultat = ""
    mettreMajuscule = True

    For i = 1 To Len(texte)

        caractere = Mid$(texte, i, 1)

        If mettreMajuscule And caractere <> " " Then
            caractere = UCase$(caractere)
            mettreMajuscule = False
        End If

        resultat = resultat & caractere

        If caractere = " " _
        Or caractere = "-" _
        Or caractere = "'" _
        Or caractere = "'" Then

            mettreMajuscule = True

        End If

    Next i

    SD_MajusculeInitialesPrenom = resultat

End Function

Private Function SD_DOSSIER_SORTIE() As String
    SD_DOSSIER_SORTIE = modConfig.CheminNasConfigure("SORTIE", "Dossier", "\\DS224\home\sortiedragon")
End Function

Public Function SD_CopierRevisionFinale(ByVal doc As Document, ByVal source As String, ByVal publicationID As String, ByRef destination As String) As Boolean
    Dim pat As Object, fso As Object, base As String
    Set pat = modIntegrationUnifie.PatientVerifie(doc)
    SD_CreerDossierSiNecessaire SD_DOSSIER_SORTIE
    base = modFichiers.NomFichierSur(UCase$(CStr(pat("Nom"))) & " " & CStr(pat("Prenom")) & " " & publicationID)
    destination = SD_DOSSIER_SORTIE & "\" & base & ".docx"
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile source, destination, True
    SD_CopierRevisionFinale = fso.FileExists(destination)
End Function
