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
    ProgrammerProchaine
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
    On Error Resume Next
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
    If mScrutationActive Then ProgrammerProchaine
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
    If fso.FileExists(dest) Then fso.DeleteFile dest, True
    fso.MoveFile cheminDrapeau, dest
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
    nomBase = Format$(Now, "yyyymmdd-hhnnss") & "_" & rdv("ID")
    PublierArrivee = modFichiers.EcrireDrapeau(dossier, nomBase, d)
    modLog.LogInfo "Arrivee publiee : " & PublierArrivee
End Function
