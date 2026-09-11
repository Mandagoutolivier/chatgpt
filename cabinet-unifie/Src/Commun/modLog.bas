Attribute VB_Name = "modLog"
Option Explicit
' =====================================================================
' modLog - Journalisation (Logs\journal_AAAAMMJJ.log) et resultats de
' tests (Logs\tests.log). Fichiers ecrits en Unicode (UTF-16 LE).
' Ne doit JAMAIS faire echouer l'appelant : On Error Resume Next.
' =====================================================================

Private mEtape As String   ' derniere etape tracee (affichee dans les messages d'erreur)

Public Sub LogInfo(ByVal msg As String)
    Ecrire "INFO", msg
End Sub

Public Sub LogErreur(ByVal msg As String)
    Ecrire "ERREUR", msg
End Sub

' Trace d'etape : memorisee pour les messages d'erreur, journalisee
Public Sub Etape(ByVal libelle As String)
    mEtape = libelle
    Ecrire "ETAPE", libelle
End Sub

Public Function DerniereEtape() As String
    DerniereEtape = mEtape
End Function

' Journal principal dans <Racine>\Logs ; si la racine est inaccessible
' (configuration du poste), journal de secours dans %TEMP%\CabinetCardio.
Private Sub Ecrire(ByVal niveau As String, ByVal msg As String)
    On Error Resume Next
    Dim fso As Object, ts As Object, dossier As String, erreurRacine As String
    dossier = modConfig.chemin("Logs")
    If Err.Number <> 0 Then
        erreurRacine = Err.Description
        Err.Clear
        dossier = Environ$("TEMP") & "\CabinetCardio"
    End If
    modFichiers.EnsureDossier dossier
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(dossier & "\journal_" & Format$(Date, "yyyymmdd") & ".log", 8, True, -1)
    If Len(erreurRacine) > 0 Then
        ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [ERREUR] racine inaccessible : " & erreurRacine
    End If
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & niveau & "] [" & Environ$("COMPUTERNAME") & "] " & msg
    ts.Close
End Sub

' --- Tests -----------------------------------------------------------

Public Sub TestDebut(ByVal suite As String)
    On Error Resume Next
    Dim fso As Object, ts As Object
    modFichiers.EnsureDossier modConfig.chemin("Logs")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(modConfig.chemin("Logs") & "\tests.log", 8, True, -1)
    ts.WriteLine "=== " & suite & " === " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ts.Close
End Sub

Public Sub TestResultat(ByVal nom As String, ByVal ok As Boolean, Optional ByVal detail As String = "")
    On Error Resume Next
    Dim fso As Object, ts As Object, statut As String
    If ok Then statut = "PASS" Else statut = "FAIL"
    modFichiers.EnsureDossier modConfig.chemin("Logs")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(modConfig.chemin("Logs") & "\tests.log", 8, True, -1)
    ts.WriteLine statut & " | " & nom & IIf(Len(detail) > 0, " | " & detail, "")
    ts.Close
    Debug.Print statut & " | " & nom & " | " & detail
End Sub

' Assertion pratique : consigne et renvoie ok pour chainage
Public Function Verifier(ByVal nom As String, ByVal condition As Boolean, _
                         Optional ByVal detail As String = "") As Boolean
    TestResultat nom, condition, detail
    Verifier = condition
End Function
