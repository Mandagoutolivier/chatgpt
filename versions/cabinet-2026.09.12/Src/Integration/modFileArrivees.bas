Attribute VB_Name = "modFileArrivees"
Option Explicit
' Une seule fenetre non modale. Les donnees restent exclusivement sur le NAS.
Private mFenetre As ufArriveesMedecin
Private mActive As Boolean
Private mPlanifie As Boolean
Private mChargement As Boolean

Public Sub Unifie_AfficherFileArrivees()
    On Error GoTo Echec
    If mFenetre Is Nothing Then Set mFenetre = New ufArriveesMedecin
    mActive = True
    If Not mFenetre.Visible Then mFenetre.Show vbModeless
    ActualiserFileArrivees
    ProgrammerFile
    Exit Sub
Echec:
    MsgBox "Liste des arrivees : " & Err.Description, vbExclamation, "Cabinet"
End Sub

Public Sub DemarrerFileAuChargement()
    Dim profil As String
    On Error GoTo Sortie
    profil = LCase$(modConfig.Config("POSTE", "Profil", ""))
    If profil <> "domicile" And profil <> "cabinet" And profil <> "medecin" Then Exit Sub
    If modConfig.ConfigNum("FILE_ARRIVEES", "AfficherAuDemarrage", 1) = 1 Then Unifie_AfficherFileArrivees
Sortie:
End Sub

Public Sub ActualiserFileArrivees()
    If Not mActive Or mChargement Then Exit Sub
    If modPowerMicUnifie.Unifie_OperationEnCours() Then Exit Sub
    If mFenetre Is Nothing Then Exit Sub
    mChargement = True
    On Error GoTo Echec
    Dim reponse As Object, items As Collection
    Set reponse = modServiceNas.Appeler("attentes", modServiceNas.Parametres())
    Set items = modAttenteLocale.FiltrerEtTrier(modServiceNas.ItemsValides(reponse), Format$(Date, "dd/mm/yyyy"))
    mFenetre.Charger items, Format$(Now, "hh:nn:ss")
Sortie:
    mChargement = False
    Exit Sub
Echec:
    ' Une liste perimee ne doit jamais rester cliquable apres une panne.
    mFenetre.Indisponible "NAS inaccessible. Actualisez avant de choisir un patient."
    modLog.Diagnostic "file_arrivees", "lecture", Err.Number
    Resume Sortie
End Sub

Private Sub ProgrammerFile()
    If Not mActive Or mPlanifie Then Exit Sub
    mPlanifie = True
    On Error GoTo Echec
    ' Word ne dispose que d un timer OnTime ; aucun autre timer Word dans U2.
    Application.OnTime When:=Now + TimeSerial(0, 0, 15), Name:="modFileArrivees.TickFileArrivees"
    Exit Sub
Echec:
    mPlanifie = False
End Sub

Public Sub TickFileArrivees()
    mPlanifie = False
    If Not mActive Then Exit Sub
    ActualiserFileArrivees
    ProgrammerFile
End Sub

Public Sub SuspendreFileArrivees()
    mActive = False
    ' Le rappel Word deja programme ne fait rien quand mActive=False.
End Sub

Public Sub FermerFileArrivees()
    SuspendreFileArrivees
    On Error Resume Next
    If Not mFenetre Is Nothing Then Unload mFenetre
    Set mFenetre = Nothing
End Sub

Public Sub DemarrerSelection(ByVal attente As Object)
    If Not mActive Or mChargement Then Exit Sub
    If modPowerMicUnifie.Unifie_OperationEnCours() Then Exit Sub
    If attente Is Nothing Then Exit Sub
    ' Aucun choix implicite du premier patient ; uniquement l objet double-clique.
    modPowerMicUnifie.Unifie_DemarrerConsultation attente
    ActualiserFileArrivees
End Sub
