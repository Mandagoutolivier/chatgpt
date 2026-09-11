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
    auj = Format$(Date, "dd/mm/yyyy")
    sql = "PRAGMA journal_mode=WAL;" & vbLf & _
          "CREATE TABLE IF NOT EXISTS attentes (" & _
          "patient_id TEXT PRIMARY KEY, nom TEXT NOT NULL, prenom TEXT NOT NULL," & _
          "ddn TEXT NOT NULL, sexe TEXT, med_traitant_id TEXT, rdv_id TEXT," & _
          "date_rdv TEXT, heure_rdv TEXT, heure_arrivee TEXT, statut TEXT," & _
          "source_nas TEXT NOT NULL, synchronise_le TEXT NOT NULL);" & vbLf & _
          "DELETE FROM attentes;" & vbLf & "BEGIN;" & vbLf
    For Each f In modFichiers.ListerFichiers(dossier, ".txt")
        Set d = modFichiers.LireDrapeau(CStr(f))
        If Valeur(d, "DateArrivee") = auj And Valeur(d, "Statut") = "Arrive" Then
            sql = sql & "INSERT OR REPLACE INTO attentes VALUES(" & _
                Q(Valeur(d, "PatientID")) & "," & Q(Valeur(d, "Nom")) & "," & _
                Q(Valeur(d, "Prenom")) & "," & Q(Valeur(d, "DDN")) & "," & _
                Q(Valeur(d, "Sexe")) & "," & Q(Valeur(d, "MedTraitantID")) & "," & _
                Q(Valeur(d, "RdvID")) & "," & Q(Valeur(d, "DateRdv")) & "," & _
                Q(Valeur(d, "HeureRdv")) & "," & Q(Valeur(d, "HeureArrivee")) & "," & _
                Q("Arrive") & "," & Q(CStr(f)) & "," & Q(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & ");" & vbLf
        End If
    Next f
    sql = sql & "COMMIT;" & vbLf
    ExecuterSql sql, False
End Sub

Public Function ChoisirAttente() As Object
    Dim sortie As String, lignes() As String, champs() As String, i As Long
    Dim col As New Collection, d As Object, f As ufListe
    SynchroniserAttentes
    sortie = ExecuterSql("SELECT patient_id,nom,prenom,ddn,sexe,med_traitant_id,rdv_id,date_rdv,heure_rdv,heure_arrivee,source_nas FROM attentes ORDER BY heure_arrivee,heure_rdv;", True)
    If Len(Trim$(sortie)) = 0 Then Exit Function
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
    If col.Count = 1 Then Set ChoisirAttente = col(1): Exit Function
    Set f = New ufListe
    f.Configurer "Patients arrives", col, Array("HeureArrivee", "Patient", "DDN"), "50 pt;200 pt;70 pt"
    f.Show vbModal
    If Not f.Annule Then Set ChoisirAttente = f.Resultat
    Unload f
End Function

Public Sub ConsommerAttente(ByVal attente As Object)
    On Error Resume Next
    Dim fso As Object, source As String, dest As String
    source = attente("SourceNas")
    Set fso = CreateObject("Scripting.FileSystemObject")
    dest = fso.GetParentFolderName(source) & "\Pris"
    modFichiers.EnsureDossier dest
    fso.MoveFile source, dest & "\" & fso.GetFileName(source)
    ExecuterSql "DELETE FROM attentes WHERE patient_id=" & Q(attente("PatientID")) & ";", False
End Sub

Private Function ExecuterSql(ByVal sql As String, ByVal capturer As Boolean) As String
    Dim fSql As String, fOut As String, cmd As String, sh As Object, rc As Long, qte As String
    qte = Chr$(34)
    fSql = DossierLocal() & "\commande.sql": fOut = DossierLocal() & "\sortie.txt"
    modFichiers.EcrireTexteUTF8 fSql, sql
    cmd = "cmd.exe /d /s /c " & qte & qte & SqliteExe() & qte & _
          " -batch -separator " & qte & Chr$(31) & qte & " " & qte & FichierDb() & qte & _
          " < " & qte & fSql & qte
    If capturer Then cmd = cmd & " > " & qte & fOut & qte
    cmd = cmd & qte
    Set sh = CreateObject("WScript.Shell")
    rc = sh.Run(cmd, 0, True)
    If rc <> 0 Then Err.Raise vbObjectError + 971, "modAttenteLocale", "Erreur SQLite (code " & rc & ")."
    If capturer And modFichiers.FichierExiste(fOut) Then ExecuterSql = modFichiers.LireTexteUTF8(fOut)
End Function

Private Function Q(ByVal valeur As String) As String
    Q = "'" & Replace(valeur, "'", "''") & "'"
End Function

Private Function Valeur(ByVal d As Object, ByVal cle As String) As String
    If d.Exists(cle) Then Valeur = CStr(d(cle)) Else Valeur = ""
End Function
