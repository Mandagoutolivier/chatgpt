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
