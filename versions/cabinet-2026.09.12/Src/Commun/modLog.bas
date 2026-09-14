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
    Diagnostic "etape", "debut"
End Sub

Public Function DerniereEtape() As String
    DerniereEtape = mEtape
End Function

' Journal principal dans <Racine>\Logs ; si la racine est inaccessible
' (configuration du poste), journal de secours dans %TEMP%\CabinetCardio.
Private Sub Ecrire(ByVal niveau As String, ByVal msg As String)
    ' Les anciens textes libres peuvent contenir des identites : jamais copies.
    Diagnostic "legacy", niveau, Err.Number
End Sub

Public Sub Diagnostic(ByVal etape As String, ByVal categorie As String, Optional ByVal numero As Long = 0, Optional ByVal commande As String = "")
    On Error Resume Next
    Dim dossier As String, chemin As String, fso As Object, ts As Object, f As Object
    dossier = Environ$("LOCALAPPDATA") & "\CabinetCardio\Diagnostics"
    modFichiers.EnsureDossier dossier
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(dossier).Files
        If Left$(f.Name, 3) = "u1-" And DateDiff("d", f.DateLastModified, Now) > 14 Then f.Delete
    Next f
    chemin = dossier & "\u1-" & Format$(Date, "yyyymmdd") & ".log"
    If fso.FileExists(chemin) Then
        If fso.GetFile(chemin).Size > 1048576 Then
            If fso.FileExists(chemin & ".1") Then fso.DeleteFile chemin & ".1"
            fso.MoveFile chemin, chemin & ".1"
        End If
    End If
    Set ts = fso.OpenTextFile(chemin, 8, True, -1)
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & "|" & CodeSur(etape) & "|" & CodeSur(categorie) & "|" & CStr(numero) & "|" & CodeSur(commande)
    ts.Close
End Sub

Private Function CodeSur(ByVal valeur As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[A-Za-z0-9_.-]{0,100}$"
    If re.Test(valeur) Then CodeSur = valeur Else CodeSur = "non_code"
End Function

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
