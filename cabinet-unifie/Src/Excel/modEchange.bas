Attribute VB_Name = "modEchange"
Option Explicit
' =====================================================================
' modEchange - Poste secretaire : detection des courriers valides par le
' medecin (fichiers-drapeaux dans Echange\AEnvoyer\), traitement (choix
' des actes, feuille de soins, journal) puis archivage dans Traites\.
' =====================================================================

Private mProchaine As Date
Private mScrutationActive As Boolean

Public Sub DemarrerScrutation()
    ArreterScrutation
    mScrutationActive = True
    VerifierEchange
End Sub

Public Sub ArreterScrutation()
    On Error Resume Next
    If mScrutationActive Then
        Application.OnTime mProchaine, NomMacroScrutation(), , False
    End If
    mScrutationActive = False
End Sub

Private Sub ProgrammerProchaine()
    Dim secondes As Long
    secondes = CLng(modConfig.ConfigNum("ECHANGE", "ScrutationSecondes", 30))
    If secondes < 10 Then secondes = 10
    mProchaine = Now + TimeSerial(0, 0, secondes)
    Application.OnTime mProchaine, NomMacroScrutation()
End Sub

Private Function NomMacroScrutation() As String
    NomMacroScrutation = "'" & ThisWorkbook.Name & "'!modEchange.VerifierEchange"
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
    If mScrutationActive Then ProgrammerProchaine
    Exit Sub
Erreur:
    Application.StatusBar = "File des courriers inaccessible : verifier le NAS."
    modLog.LogErreur "Scrutation du NAS impossible."
    Resume Sortie
End Sub

Public Function NombreEnAttente() As Long
    NombreEnAttente = modFichiers.ListerFichiers(modConfig.Chemin("Echange") & "\AEnvoyer", ".txt").Count
End Function

' Liste des drapeaux en attente (dictionnaires + _Chemin/_Libelle)
Public Function CourriersEnAttente() As Collection
    Dim col As Collection, dossier As String, f As Variant, d As Object, fso As Object
    Set col = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    dossier = modConfig.Chemin("Echange") & "\AEnvoyer"
    For Each f In modFichiers.ListerFichiers(dossier, ".txt")
        Set d = modFichiers.LireDrapeau(CStr(f))
        d("_Chemin") = CStr(f)
        d("ID") = fso.GetFileName(CStr(f))
        If Not d.Exists("Nom") Then d("Nom") = "?"
        If Not d.Exists("Prenom") Then d("Prenom") = ""
        If Not d.Exists("TypeCourrier") Then d("TypeCourrier") = ""
        If Not d.Exists("DateValidation") Then d("DateValidation") = ""
        col.Add d
    Next f
    Set CourriersEnAttente = col
End Function

Public Sub DeplacerVersTraites(ByVal cheminDrapeau As String)
    Dim fso As Object, dest As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    modFichiers.EnsureDossier modConfig.Chemin("Echange") & "\Traites"
    dest = modConfig.Chemin("Echange") & "\Traites\" & fso.GetFileName(cheminDrapeau)
    If fso.FileExists(dest) Then
        If fso.FileExists(cheminDrapeau) Then Err.Raise vbObjectError + 910, "modEchange", "Conflit dans l historique des courriers."
        Exit Sub
    End If
    modFichiers.RenommerAtomique cheminDrapeau, dest
End Sub

Public Sub OuvrirCourrier(ByVal d As Object)
    Dim chemin As String
    If d.Exists("CheminPdf") Then chemin = d("CheminPdf")
    If d.Exists("CheminDocx") Then
        If modFichiers.FichierExiste(CStr(d("CheminDocx"))) Then chemin = d("CheminDocx")
    End If
    If Len(chemin) = 0 Or Not modFichiers.FichierExiste(chemin) Then
        MsgBox "Fichier du courrier introuvable.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    ThisWorkbook.FollowHyperlink chemin
End Sub

' Publie sur le NAS l'identite minimale d'un patient marque Arrive.
' Ce drapeau est la source de synchronisation du poste medecin ; il ne
' contient ni adresse, ni telephone, ni donnees comptables.
Public Function PublierArrivee(ByVal rdv As Object, ByVal pat As Object) As String
    Dim d As Object, dossier As String, nomBase As String
    If modTexte.DateFr(CStr(rdv("Date"))) <> Date Then Err.Raise vbObjectError + 912, "modEchange", "Seul un rendez-vous du jour peut etre marque arrive."
    If CStr(rdv("PatientID")) <> CStr(pat("ID")) Then Err.Raise vbObjectError + 913, "modEchange", "Patient et rendez-vous incoherents."
    Dim controleGdt As String
    controleGdt = modGdt.ConstruireGdt(pat)  ' valider l identite sans ecrire sur le poste secretariat
    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = pat("ID")
    d("Nom") = pat("Nom")
    d("Prenom") = pat("Prenom")
    d("DDN") = pat("DDN")
    d("Sexe") = pat("Sexe")
    d("MedTraitantID") = pat("MedTraitantID")
    d("RdvID") = rdv("ID")
    d("DateRdv") = rdv("Date")
    d("HeureRdv") = rdv("Heure")
    d("HeureArrivee") = Format$(Now, "hh:nn")
    d("DateArrivee") = Format$(Date, "dd/mm/yyyy")
    d("Statut") = "Arrive"
    d("PosteSecretariat") = Environ$("COMPUTERNAME")
    dossier = modConfig.Chemin("Echange") & "\Arrives"
    nomBase = Format$(Date, "yyyymmdd") & "_" & rdv("ID")
    If modFichiers.FichierExiste(dossier & "\" & nomBase & ".txt") Then
        PublierArrivee = dossier & "\" & nomBase & ".txt"
        Exit Function
    End If
    If modFichiers.FichierExiste(dossier & "\Pris\" & nomBase & ".txt") Or _
       modFichiers.FichierExiste(dossier & "\EnCours\" & nomBase & ".txt") Then
        Err.Raise vbObjectError + 911, "modEchange", "Ce rendez-vous est deja pris en charge par le medecin."
    End If
    PublierArrivee = modFichiers.EcrireDrapeau(dossier, nomBase, d)
    modLog.LogInfo "Arrivee publiee : " & PublierArrivee
End Function

' Utilisee par le bouton d'accueil ET par la grille de l'agenda.
' L'arrivee est une publication durable ; le statut est reparable en reexecutant.
Public Sub SignalerArrivee(ByVal rdv As Object)
    Dim p As Object, pat As Object, verrou As String, numero As Long, description As String
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        If CStr(p("ID")) = CStr(rdv("PatientID")) Then Set pat = p: Exit For
    Next p
    If pat Is Nothing Then Err.Raise vbObjectError + 914, "modEchange", "Patient absent de la base NAS."
    verrou = "arrivee_" & CStr(rdv("ID"))
    If Not modFichiers.AcquerirVerrou(verrou, 5000) Then Err.Raise vbObjectError + 915, "modEchange", "Arrivee deja en cours."
    On Error GoTo Echec
    PublierArrivee rdv, pat
    modAgenda.MarquerStatut CStr(rdv("ID")), "Arrive", Year(modTexte.DateFr(CStr(rdv("Date"))))
    modFichiers.RelacherVerrou verrou
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    modFichiers.RelacherVerrou verrou
    Err.Raise numero, "modEchange.SignalerArrivee", description
End Sub

Public Sub RetirerArrivee(ByVal rdv As Object)
    Dim dossier As String, source As String
    dossier = modConfig.Chemin("Echange") & "\Arrives"
    source = dossier & "\" & Format$(modTexte.DateFr(CStr(rdv("Date"))), "yyyymmdd") & "_" & rdv("ID") & ".txt"
    If modFichiers.FichierExiste(source) Then
        modFichiers.EnsureDossier dossier & "\Annules"
        modFichiers.RenommerAtomique source, dossier & "\Annules\" & modFichiers.IdUnique() & ".txt"
    End If
End Sub
