Attribute VB_Name = "modAttenteLocale"
Option Explicit

' Cache SQLite local, regenerable depuis les drapeaux d'arrivee du NAS.
' La base maitresse reste exclusivement sur le Synology.

Private Function DossierLocal() As String
    DossierLocal = Environ$("LOCALAPPDATA") & "\CabinetCardio"
    modFichiers.EnsureDossier DossierLocal
End Function

Private Function FichierDb() As String
    FichierDb = DossierLocal() & "\attente_medecin.sqlite"
End Function

Private Function SqliteExe() As String
    SqliteExe = Environ$("APPDATA") & "\CabinetCardio\Tools\sqlite3.exe"
    If Not modFichiers.FichierExiste(SqliteExe) Then Err.Raise vbObjectError + 970, _
        "modAttenteLocale", "sqlite3.exe introuvable : " & SqliteExe
End Function

Public Sub SynchroniserAttentes()
    Dim dossier As String, f As Variant, d As Object, sql As String, auj As String
    dossier = modConfig.Chemin("Echange") & "\Arrives"
    If Not modFichiers.DossierExiste(dossier) Then Err.Raise vbObjectError + 972, "modAttenteLocale", "File NAS inaccessible. Aucune selection locale autorisee."
    auj = Format$(Date, "dd/mm/yyyy")
    sql = "PRAGMA journal_mode=WAL;" & vbLf & _
          "CREATE TABLE IF NOT EXISTS attentes_v2 (" & _
          "patient_id TEXT NOT NULL, nom TEXT NOT NULL, prenom TEXT NOT NULL," & _
          "ddn TEXT NOT NULL, sexe TEXT, med_traitant_id TEXT, rdv_id TEXT PRIMARY KEY," & _
          "date_rdv TEXT, heure_rdv TEXT, heure_arrivee TEXT, statut TEXT," & _
          "source_nas TEXT NOT NULL, synchronise_le TEXT NOT NULL);" & vbLf & _
          "BEGIN IMMEDIATE;" & vbLf & "DELETE FROM attentes_v2;" & vbLf
    For Each f In modFichiers.ListerFichiers(dossier, ".txt")
        Set d = modFichiers.LireDrapeau(CStr(f))
        If Valeur(d, "DateArrivee") = auj And Valeur(d, "Statut") = "Arrive" Then
            If Len(Valeur(d, "PatientID")) = 0 Or Len(Valeur(d, "RdvID")) = 0 Then Err.Raise vbObjectError + 973, "modAttenteLocale", "Arrivee sans identifiant."
            sql = sql & "INSERT INTO attentes_v2 VALUES(" & _
                Q(Valeur(d, "PatientID")) & "," & Q(Valeur(d, "Nom")) & "," & _
                Q(Valeur(d, "Prenom")) & "," & Q(Valeur(d, "DDN")) & "," & _
                Q(Valeur(d, "Sexe")) & "," & Q(Valeur(d, "MedTraitantID")) & "," & _
                Q(Valeur(d, "RdvID")) & "," & Q(Valeur(d, "DateRdv")) & "," & _
                Q(Valeur(d, "HeureRdv")) & "," & Q(Valeur(d, "HeureArrivee")) & "," & _
                Q("Arrive") & "," & Q(CStr(f)) & "," & Q(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & ");" & vbLf
        End If
    Next f
    sql = sql & "COMMIT;" & vbLf & "DROP TABLE IF EXISTS attentes;" & vbLf
    ExecuterSql sql, False
End Sub

Public Function ChoisirAttente() As Object
    Dim sortie As String, lignes() As String, champs() As String, i As Long
    Dim col As New Collection, d As Object, f As ufListe
    SynchroniserAttentes
    sortie = ExecuterSql("SELECT patient_id,nom,prenom,ddn,sexe,med_traitant_id,rdv_id,date_rdv,heure_rdv,heure_arrivee,source_nas FROM attentes_v2 ORDER BY heure_arrivee,heure_rdv;", True)
    If Len(Trim$(sortie)) = 0 Then MsgBox "Aucun patient arrive en attente aujourd hui.", vbInformation, "Cabinet": Exit Function
    lignes = Split(Replace(sortie, vbCr, ""), vbLf)
    For i = LBound(lignes) To UBound(lignes)
        If Len(lignes(i)) > 0 Then
            champs = Split(lignes(i), Chr$(31))
            If UBound(champs) >= 10 Then
                Set d = CreateObject("Scripting.Dictionary")
                d("PatientID") = champs(0): d("Nom") = champs(1): d("Prenom") = champs(2)
                d("DDN") = champs(3): d("Sexe") = champs(4): d("MedTraitantID") = champs(5)
                d("RdvID") = champs(6): d("DateRdv") = champs(7): d("HeureRdv") = champs(8)
                d("HeureArrivee") = champs(9): d("SourceNas") = champs(10)
                d("Patient") = d("Nom") & " " & d("Prenom")
                col.Add d
            End If
        End If
    Next i
    If col.Count = 0 Then Exit Function
    Set f = New ufListe
    f.Configurer "Patients arrives", col, Array("HeureArrivee", "Patient", "DDN"), "50 pt;200 pt;70 pt"
    f.Show vbModal
    If Not f.Annule Then Set ChoisirAttente = f.Resultat
    Unload f
End Function

Public Sub ConsommerAttente(ByVal attente As Object)
    Dim source As String, dest As String, dossier As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    dossier = modConfig.Chemin("Echange") & "\Arrives"
    source = CStr(attente("SourceNas"))
    If StrComp(fso.GetParentFolderName(source), dossier, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 974, "modAttenteLocale", "Chemin d arrivee inattendu."
    modFichiers.EnsureDossier dossier & "\EnCours"
    dest = dossier & "\EnCours\" & fso.GetFileName(source)
    ' Renommage sans remplacement : un seul poste peut reclamer cette arrivee.
    modFichiers.RenommerAtomique source, dest
    attente("ReservationNas") = dest
    ' La file NAS fait foi. Un echec du cache sera repare au rafraichissement suivant.
    On Error Resume Next
    ExecuterSql "DELETE FROM attentes_v2 WHERE rdv_id=" & Q(CStr(attente("RdvID"))) & ";", False
    On Error GoTo 0
End Sub
Private Function ExecuterSql(ByVal sql As String, ByVal capturer As Boolean) As String
    Dim fSql As String, fOut As String, fErr As String, cmd As String, sh As Object, rc As Long, qte As String
    Dim numero As Long, description As String, exe As String, db As String
    exe = SqliteExe(): db = FichierDb()
    VerifierCheminCommande exe: VerifierCheminCommande db
    qte = Chr$(34)
    fSql = DossierLocal() & "\" & modFichiers.IdUnique() & ".sql"
    fOut = fSql & ".out": fErr = fSql & ".err"
    On Error GoTo Echec
    modFichiers.EcrireTexteUTF8 fSql, ".bail on" & vbLf & ".timeout 5000" & vbLf & sql
    cmd = "cmd.exe /d /v:off /s /c " & qte & qte & exe & qte & _
          " -batch -bail -separator " & qte & Chr$(31) & qte & " " & qte & db & qte & _
          " < " & qte & fSql & qte & " > " & qte & fOut & qte & " 2> " & qte & fErr & qte & qte
    Set sh = CreateObject("WScript.Shell")
    rc = sh.Run(cmd, 0, True)
    If rc <> 0 Then Err.Raise vbObjectError + 971, "modAttenteLocale", "Erreur SQLite (code " & rc & "). La selection est interrompue."
    If capturer Then ExecuterSql = modFichiers.LireTexteUTF8(fOut)
Sortie:
    modFichiers.SupprimerTemporaire fSql
    modFichiers.SupprimerTemporaire fOut
    modFichiers.SupprimerTemporaire fErr
    If numero <> 0 Then Err.Raise numero, "modAttenteLocale.ExecuterSql", description
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    Resume Sortie
End Function
Private Function Q(ByVal valeur As String) As String
    Dim i As Long
    For i = 0 To 31
        If InStr(valeur, Chr$(i)) > 0 Then Err.Raise vbObjectError + 975, "modAttenteLocale", "Caractere de controle interdit dans une arrivee."
    Next i
    Q = "'" & Replace(valeur, "'", "''") & "'"
End Function
Private Function Valeur(ByVal d As Object, ByVal cle As String) As String
    If d.Exists(cle) Then Valeur = CStr(d(cle)) Else Valeur = ""
End Function

Private Sub VerifierCheminCommande(ByVal chemin As String)
    Dim c As Variant
    For Each c In Array("%", "!", "&", "|", "<", ">", "^", Chr$(34), vbCr, vbLf)
        If InStr(chemin, CStr(c)) > 0 Then Err.Raise vbObjectError + 976, "modAttenteLocale", "Chemin Windows incompatible avec l appel SQLite."
    Next c
    If Left$(chemin, 2) = "\\" Then Err.Raise vbObjectError + 977, "modAttenteLocale", "SQLite doit etre local au poste."
End Sub

Public Sub LibererReservation(ByVal attente As Object)
    If attente Is Nothing Then Exit Sub
    If Not attente.Exists("ReservationNas") Then Exit Sub
    If modFichiers.FichierExiste(CStr(attente("ReservationNas"))) Then
        modFichiers.RenommerAtomique CStr(attente("ReservationNas")), CStr(attente("SourceNas"))
    End If
End Sub

Public Sub EnregistrerBrouillon(ByVal attente As Object, ByVal cheminBrouillon As String)
    Dim d As Object, cle As Variant, texte As String, tmp As String, cible As String
    Dim numero As Long, description As String
    cible = CStr(attente("ReservationNas"))
    Set d = modFichiers.LireDrapeau(cible)
    d("CheminBrouillon") = cheminBrouillon
    For Each cle In d.Keys
        texte = texte & cle & "=" & Replace(Replace(CStr(d(cle)), vbCr, " "), vbLf, " ") & vbCrLf
    Next cle
    tmp = cible & "." & modFichiers.IdUnique() & ".tmp"
    On Error GoTo Echec
    modFichiers.EcrireTexteUTF8 tmp, texte
    modFichiers.RenommerAtomique tmp, cible, True
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    modFichiers.SupprimerTemporaire tmp
    Err.Raise numero, "modAttenteLocale.EnregistrerBrouillon", description
End Sub
