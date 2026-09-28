Attribute VB_Name = "modEchange"
Option Explicit
' =====================================================================
' modEchange - Poste secretaire : detection des courriers valides par le
' medecin (fichiers-drapeaux dans Echange\AEnvoyer\), traitement (choix
' des actes, feuille de soins, journal) puis archivage dans Traites\.
' =====================================================================

Private mProchaine As Date
Private mScrutationActive As Boolean
Private mPlanifie As Boolean

Public Sub DemarrerScrutation()
    ArreterScrutation
    mScrutationActive = True
    VerifierEchange
    ProgrammerProchaine
End Sub

Public Sub ArreterScrutation()
    On Error Resume Next
    If mPlanifie Then
        Application.OnTime mProchaine, NomMacroScrutation(), , False
    End If
    mPlanifie = False
    mScrutationActive = False
End Sub

Private Sub ProgrammerProchaine()
    If Not mScrutationActive Or mPlanifie Then Exit Sub
    Dim secondes As Long
    secondes = CLng(modConfig.ConfigNum("ECHANGE", "ScrutationSecondes", 30))
    If secondes < 10 Then secondes = 10
    mProchaine = Now + TimeSerial(0, 0, secondes)
    Application.OnTime mProchaine, NomMacroScrutation()
    mPlanifie = True
End Sub

Private Function NomMacroScrutation() As String
    NomMacroScrutation = "'" & ThisWorkbook.Name & "'!modEchange.TickScrutation"
End Function

' Appelee par OnTime : met a jour le compteur sur la feuille Accueil
Public Sub VerifierEchange()
    On Error GoTo Erreur
    Dim n As Long, ws As Worksheet
    n = NombreEnAttente()
    Set ws = ThisWorkbook.Worksheets("Accueil")
    If n > 0 Then
        ws.Range("B4").Value = ">>> " & n & " courrier(s) valide(s) en attente de traitement <<<"
        ws.Range("B4").Font.Bold = True
        ws.Range("B4").Font.Color = RGB(192, 0, 0)
        Application.StatusBar = n & " courrier(s) du medecin en attente"
    Else
        ws.Range("B4").Value = "Aucun courrier en attente."
        ws.Range("B4").Font.Bold = False
        ws.Range("B4").Font.Color = RGB(0, 128, 0)
        Application.StatusBar = False
    End If
Sortie:
    Exit Sub
Erreur:
    Application.StatusBar = "File des courriers inaccessible : verifier le NAS."
    modLog.LogErreur "Scrutation du NAS impossible."
    Resume Sortie
End Sub

Public Function NombreEnAttente() As Long
    NombreEnAttente = CourriersEnAttente().Count
End Function
' Liste des drapeaux en attente (dictionnaires + _Chemin/_Libelle)
Public Function CourriersEnAttente() As Collection
    Dim r As Object, d As Object, col As New Collection
    Set r = modServiceNas.Appeler("publications", modServiceNas.Parametres())
    For Each d In r("items")
        d("_Chemin") = CStr(d("PublicationID")): d("ID") = CStr(d("PublicationID"))
        col.Add d
    Next d
    Set CourriersEnAttente = col
End Function
Public Sub DeplacerVersTraites(ByVal cheminDrapeau As String)
    Dim r As Object
    Set r = modServiceNas.CommandeID("ack", cheminDrapeau)
End Sub
' Les archives NAS ne sont jamais ouvertes directement dans une application.
' Consulter une copie locale controlee ; corriger via modEditionSecretariat.
Public Sub OuvrirCourrier(ByVal d As Object)
    Dim copie As String
    copie = CopieLectureArchive(d)
    If LCase$(Right$(copie, 5)) = ".docx" Then
        OuvrirCopieWordLectureSeule copie
    Else
        ThisWorkbook.FollowHyperlink copie
    End If
End Sub

Private Function CopieLectureArchive(ByVal d As Object) As String
    Dim source As String, extension As String, empreinte As String
    Dim dossier As String, copie As String, re As Object, fso As Object
    If d Is Nothing Then Err.Raise vbObjectError + 1116, , "Publication absente."
    extension = "pdf"
    If d.Exists("CheminDocx") Then
        If modFichiers.FichierExiste(CStr(d("CheminDocx"))) Then extension = "docx"
    End If
    If extension = "docx" Then
        source = CStr(d("CheminDocx"))
    ElseIf d.Exists("CheminPdf") Then
        source = CStr(d("CheminPdf"))
    End If
    If Len(source) = 0 Then Err.Raise vbObjectError + 1116, , "Fichier du courrier introuvable."
    If Not d.Exists("sha_" & extension) Then Err.Raise vbObjectError + 1116, , "Empreinte de publication absente."
    empreinte = LCase$(CStr(d("sha_" & extension)))
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9a-f]{64}$"
    If Not re.Test(empreinte) Then Err.Raise vbObjectError + 1116, , "Empreinte de publication invalide."
    If StrComp(source, modConfig.chemin("Documents") & "\" & empreinte & "." & extension, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1116, , "Fichier hors des archives du cabinet."
    End If
    If Not modFichiers.FichierExiste(source) Then Err.Raise vbObjectError + 1116, , "Archive introuvable."
    If modDonneesTransport.EmpreinteFichierSHA256(source) <> empreinte Then
        Err.Raise vbObjectError + 1116, , "Archive modifiee : ouverture interrompue."
    End If
    dossier = Environ$("LOCALAPPDATA")
    If Len(dossier) = 0 Then Err.Raise vbObjectError + 1116, , "Dossier local de consultation indisponible."
    dossier = dossier & "\CabinetCardio\LecturesArchives"
    modFichiers.EnsureDossier dossier
    copie = dossier & "\" & empreinte & "." & extension
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(copie) Then fso.CopyFile source, copie, False
    If modDonneesTransport.EmpreinteFichierSHA256(copie) <> empreinte Then
        Err.Raise vbObjectError + 1116, , "Copie de consultation modifiee : ouverture interrompue."
    End If
    SetAttr copie, GetAttr(copie) Or vbReadOnly
    CopieLectureArchive = copie
End Function

Private Sub OuvrirCopieWordLectureSeule(ByVal chemin As String)
    Dim word As Object, doc As Object, existant As Object
    Dim securite As Long, liens As Boolean, optionsCapturees As Boolean
    Dim wordCree As Boolean, docCree As Boolean, numero As Long, description As String
    On Error Resume Next
    Set word = GetObject(, "Word.Application")
    Err.Clear
    On Error GoTo Echec
    If word Is Nothing Then
        Set word = CreateObject("Word.Application")
        wordCree = True
    End If
    For Each existant In word.Documents
        If StrComp(CStr(existant.FullName), chemin, vbTextCompare) = 0 Then
            If Not existant.ReadOnly Then Err.Raise vbObjectError + 1116, , "Copie deja ouverte en modification : fermez-la avant de consulter."
            word.Visible = True
            existant.Activate
            Exit Sub
        End If
    Next existant
    securite = word.AutomationSecurity
    liens = word.Options.UpdateLinksAtOpen
    optionsCapturees = True
    word.AutomationSecurity = 3
    word.Options.UpdateLinksAtOpen = False
    Set doc = word.Documents.Open(chemin, False, True, False)
    docCree = True
    doc.ActiveWindow.View.ReadingLayout = False
    If Not doc.ReadOnly Then Err.Raise vbObjectError + 1116, , "La consultation en lecture seule a echoue."
    word.AutomationSecurity = securite
    word.Options.UpdateLinksAtOpen = liens
    optionsCapturees = False
    word.Visible = True
    doc.Activate
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If docCree Then doc.Close 0
    If optionsCapturees Then
        word.AutomationSecurity = securite
        word.Options.UpdateLinksAtOpen = liens
    End If
    If wordCree Then word.Quit 0
    On Error GoTo 0
    Err.Raise numero, "Consultation du courrier", description
End Sub

' Publie sur le NAS l'identite minimale d'un patient marque Arrive.
' Ce drapeau est la source de synchronisation du poste medecin ; il ne
' contient ni adresse, ni telephone, ni donnees comptables.
Public Function PublierArrivee(ByVal rdv As Object, ByVal pat As Object) As String
    Dim r As Object
    If CStr(rdv("PatientID")) <> CStr(pat("ID")) Then Err.Raise vbObjectError + 1115, , "Patient different du rendez-vous."
    Set r = modServiceNas.CommandeID("arrive", CStr(rdv("ID")))
    PublierArrivee = CStr(r("ConsultationID"))
End Function
' Utilisee par le bouton d'accueil ET par la grille de l'agenda.
' L'arrivee est une publication durable ; le statut est reparable en reexecutant.
Public Sub SignalerArrivee(ByVal rdv As Object)
    Dim r As Object
    Set r = modServiceNas.CommandeID("arrive", CStr(rdv("ID")))
End Sub
Public Sub RetirerArrivee(ByVal rdv As Object)
    Dim r As Object
    Set r = modServiceNas.CommandeID("cancel_arrival", CStr(rdv("ID")))
End Sub

Public Sub TickScrutation()
    mPlanifie = False
    If Not mScrutationActive Then Exit Sub
    VerifierEchange
    ProgrammerProchaine
End Sub
