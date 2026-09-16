Attribute VB_Name = "modConfig"
Option Explicit
' =====================================================================
' modConfig - Configuration centrale du logiciel de cabinet
' Racine des donnees : lue dans %APPDATA%\CabinetCardio\chemin.txt
' (ou imposee par les tests via DefinirRacine).
' Parametres : <Racine>\Config\config.ini (UTF-8, sections [X] cle=valeur)
' =====================================================================

Private mRacine As String
Private mIni As Object          ' Scripting.Dictionary : "section|cle" -> valeur
Private mIniCharge As Boolean

' Utilise par les tests et l'installateur pour imposer la racine
Public Sub DefinirRacine(ByVal nouvelleRacine As String)
    If Right$(nouvelleRacine, 1) = "\" Then nouvelleRacine = Left$(nouvelleRacine, Len(nouvelleRacine) - 1)
    mRacine = nouvelleRacine
    mIniCharge = False
End Sub

' Enregistre la racine du poste (%APPDATA%\CabinetCardio\chemin.txt).
' Utilise par le deploiement via COM : le fichier est ainsi ecrit par
' l'application Office elle-meme (visible d'elle a coup sur).
Public Sub EcrireCheminRacine(ByVal nouvelleRacine As String)
    Dim dossier As String
    dossier = Environ$("APPDATA") & "\CabinetCardio"
    modFichiers.EnsureDossier dossier
    modFichiers.EcrireTexteUTF8 dossier & "\chemin.txt", nouvelleRacine & vbCrLf
    DefinirRacine nouvelleRacine
End Sub

Public Function racine() As String
    Dim f As String, t As String
    If Len(mRacine) > 0 Then racine = mRacine: Exit Function
    f = Environ$("APPDATA") & "\CabinetCardio\chemin.txt"
    If Not modFichiers.FichierExiste(f) Then
        Err.Raise vbObjectError + 100, "modConfig", _
            "Fichier introuvable : " & f & vbCrLf & _
            "Ce fichier doit contenir le chemin du dossier CabinetCardio (ex : \\POSTE-SECRETAIRE\CabinetCardio)."
    End If
    t = modFichiers.LireTexteUTF8(f)
    t = Replace(t, vbCr, vbLf)
    If InStr(t, vbLf) > 0 Then t = Left$(t, InStr(t, vbLf) - 1)
    t = Trim$(t)
    If Right$(t, 1) = "\" Then t = Left$(t, Len(t) - 1)
    If Left$(t, 2) <> "\\" Then Err.Raise vbObjectError + 101, "modConfig", "La racine des donnees doit etre un partage NAS UNC."
    If Len(t) < 6 Then Err.Raise vbObjectError + 102, "modConfig", "Racine NAS invalide."
    mRacine = t
    racine = mRacine
End Function

Public Function chemin(ByVal sousDossier As String) As String
    chemin = racine() & "\" & sousDossier
End Function

Public Function FichierPatients() As String
    FichierPatients = chemin("Base") & "\Patients.xlsx"
End Function

Public Function FichierAgenda(Optional ByVal annee As Long = 0) As String
    If annee = 0 Then annee = Year(Date)
    FichierAgenda = chemin("Base") & "\Agenda_" & annee & ".xlsx"
End Function

Public Function FichierJournal(Optional ByVal annee As Long = 0) As String
    If annee = 0 Then annee = Year(Date)
    FichierJournal = chemin("Actes") & "\Journal_" & annee & ".xlsx"
End Function

Public Function FichierNomenclature() As String
    FichierNomenclature = chemin("Config") & "\Nomenclature.xlsx"
End Function

' Lecture d'un parametre de config.ini
Public Function Config(ByVal section As String, ByVal cle As String, _
                       Optional ByVal defaut As String = "") As String
    Dim k As String
    ChargerIni
    k = LCase$(section) & "|" & LCase$(cle)
    If mIni.Exists(k) Then Config = mIni(k) Else Config = defaut
End Function

Public Function ConfigNum(ByVal section As String, ByVal cle As String, _
                          Optional ByVal defaut As Double = 0) As Double
    Dim v As String
    v = Config(section, cle, "")
    If Len(v) = 0 Then
        ConfigNum = defaut
    Else
        ConfigNum = Val(Replace(v, ",", "."))
    End If
End Function

Public Function ConfigBool(ByVal section As String, ByVal cle As String, _
                           Optional ByVal defaut As Boolean = False) As Boolean
    Dim v As String
    v = LCase$(Config(section, cle, ""))
    If Len(v) = 0 Then
        ConfigBool = defaut
    Else
        ConfigBool = (v = "1" Or v = "oui" Or v = "true" Or v = "vrai")
    End If
End Function

Public Sub RechargerConfig()
    mIniCharge = False
End Sub

' Parseur commun a la lecture effective et aux tests, sans acces au poste.
' Les doublons SORTIE sont refuses, meme identiques : jamais premier/dernier gagne.
Private Function NettoyerBordsIni(ByVal texte As String) As String
    Do While Len(texte) > 0
        If Left$(texte, 1) <> " " And Left$(texte, 1) <> vbTab Then Exit Do
        texte = Mid$(texte, 2)
    Loop
    Do While Len(texte) > 0
        If Right$(texte, 1) <> " " And Right$(texte, 1) <> vbTab Then Exit Do
        texte = Left$(texte, Len(texte) - 1)
    Loop
    NettoyerBordsIni = texte
End Function

Public Function AnalyserConfigurationIni(ByVal contenu As String) As Object
    Dim valeurs As Object, lignes() As String, sortieVue As Boolean
    Dim i As Long, section As String, ligne As String, p As Long, cle As String, k As String
    Set valeurs = CreateObject("Scripting.Dictionary")
    contenu = Replace(contenu, vbCrLf, vbLf)
    contenu = Replace(contenu, vbCr, vbLf)
    lignes = Split(contenu, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        ligne = NettoyerBordsIni(lignes(i))
        If Len(ligne) = 0 Then
            ' vide
        ElseIf Left$(ligne, 1) = ";" Or Left$(ligne, 1) = "#" Then
            ' commentaire
        ElseIf Left$(ligne, 1) = "[" And InStr(ligne, "]") > 1 Then
            section = LCase$(NettoyerBordsIni(Mid$(ligne, 2, InStr(ligne, "]") - 2)))
            If section = "sortie" Then
                If sortieVue Then Err.Raise vbObjectError + 106, "modConfig", "Configuration ambigue : section [SORTIE] en double."
                sortieVue = True
            End If
        Else
            p = InStr(ligne, "=")
            If p > 0 Then
                cle = LCase$(NettoyerBordsIni(Left$(ligne, p - 1)))
                k = section & "|" & cle
                If section = "sortie" And (cle = "exportactif" Or cle = "dossier" Or cle = "nomfichier") Then
                    If valeurs.Exists(k) Then Err.Raise vbObjectError + 107, "modConfig", "Configuration ambigue : cle SORTIE/" & cle & " en double."
                End If
                valeurs(k) = NettoyerBordsIni(Mid$(ligne, p + 1))
            End If
        End If
    Next i
    Set AnalyserConfigurationIni = valeurs
End Function

Private Sub ChargerIni()
    Dim cheminIni As String, contenu As String, lignes() As String
    Dim i As Long, section As String, ligne As String, p As Long
    If mIniCharge Then Exit Sub
    cheminIni = chemin("Config") & "\config.ini"
    If modFichiers.FichierExiste(cheminIni) Then contenu = modFichiers.LireTexteUTF8(cheminIni)
    Set mIni = AnalyserConfigurationIni(contenu)
    ' Ces reglages dependent du poste. Les donnees partagees restent sur le NAS.
    cheminIni = Environ$("APPDATA") & "\CabinetCardio\poste.ini"
    If modFichiers.FichierExiste(cheminIni) Then
        contenu = Replace(modFichiers.LireTexteUTF8(cheminIni), vbCrLf, vbLf)
        lignes = Split(contenu, vbLf)
        section = ""
        For i = LBound(lignes) To UBound(lignes)
            ligne = Trim$(lignes(i))
            If Left$(ligne, 1) = "[" And InStr(ligne, "]") > 1 Then
                section = LCase$(Mid$(ligne, 2, InStr(ligne, "]") - 2))
            ElseIf Left$(ligne, 1) <> ";" And Left$(ligne, 1) <> "#" Then
                p = InStr(ligne, "=")
                If p > 1 Then
                    Dim k As String
                    k = section & "|" & LCase$(Trim$(Left$(ligne, p - 1)))
                    Select Case k
                        Case "ecg|dossiergdt", "cerfa|imprimante", "cerfa|calagevalide", "cerfa|praticienpreimprime", "poste|profil"
                            mIni(k) = Trim$(Mid$(ligne, p + 1))
                    End Select
                End If
            End If
        Next i
    End If
    mIniCharge = True
End Sub

Public Function CheminNasConfigure(ByVal section As String, ByVal cle As String, ByVal defaut As String) As String
    Dim valeur As String
    valeur = Config(section, cle, defaut)
    If Len(valeur) = 0 Then Err.Raise vbObjectError + 103, "modConfig", "Chemin NAS non configure : " & section & "/" & cle
    If InStr(valeur, "..") > 0 Then Err.Raise vbObjectError + 104, "modConfig", "Chemin relatif invalide."
    If Left$(valeur, 2) = "\\" Then
        CheminNasConfigure = valeur
    Else
        If InStr(valeur, ":") > 0 Or Left$(valeur, 1) = "\" Then Err.Raise vbObjectError + 105, "modConfig", "Une base ou un modele partage doit rester sur le NAS."
        CheminNasConfigure = Chemin(valeur)
    End If
End Function
