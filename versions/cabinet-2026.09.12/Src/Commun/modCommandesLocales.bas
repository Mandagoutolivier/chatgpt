Attribute VB_Name = "modCommandesLocales"
Option Explicit

Public Function CleCommandeLocale(ByVal serveur As String, ByVal compte As String, ByVal operation As String, ByVal params As Object) As String
    Dim contexte As Object
    If Len(Trim$(compte)) = 0 Then Err.Raise vbObjectError + 1108, , "Compte de reprise absent."
    Set contexte = modServiceNas.Parametres()
    contexte("serveur") = serveur: contexte("compte") = compte: contexte("operation") = operation
    Set contexte("params") = params
    CleCommandeLocale = modDonneesTransport.EmpreinteSHA256(modDonneesTransport.SerialiserJson(contexte))
End Function

Private Sub MigrerIdentifiantLocal(ByVal fso As Object, ByVal ancien As String, ByVal nouveau As String)
    Dim ts As Object, idAncien As String, idNouveau As String, numero As Long
    If ancien = nouveau Or Not fso.FileExists(ancien) Then Exit Sub
    Set ts = fso.OpenTextFile(ancien, 1, False, 0): idAncien = Trim$(ts.ReadAll): ts.Close
    If Not fso.FileExists(nouveau) Then
        On Error Resume Next
        fso.MoveFile ancien, nouveau
        numero = Err.Number: Err.Clear
        On Error GoTo 0
        If numero = 0 Then Exit Sub
        If Not fso.FileExists(nouveau) Then Err.Raise vbObjectError + 1108, , "Migration de commande impossible ; fichiers conserves."
    End If
    Set ts = fso.OpenTextFile(nouveau, 1, False, 0): idNouveau = Trim$(ts.ReadAll): ts.Close
    If idAncien <> idNouveau Then Err.Raise vbObjectError + 1108, , "Deux identifiants de reprise differents : diagnostic requis ; fichiers conserves."
    If fso.FileExists(ancien) Then fso.DeleteFile ancien
End Sub

Public Function ObtenirCommandeLocale(ByVal cle As String, ByRef chemin As String, ByRef reprise As Boolean, Optional ByVal ancienneCle As String = "", Optional ByVal dossierRecette As String = "") As String
    Dim dossier As String, fso As Object, ts As Object, temporaire As String, id As String, re As Object
    dossier = Environ$("LOCALAPPDATA") & "\CabinetCardio\Commandes"
    If Len(dossierRecette) > 0 Then dossier = dossierRecette
    modFichiers.EnsureDossier dossier
    Set fso = CreateObject("Scripting.FileSystemObject")
    chemin = dossier & "\" & cle & ".pending"
    If Len(ancienneCle) > 0 Then MigrerIdentifiantLocal fso, dossier & "\" & ancienneCle & ".pending", chemin
    reprise = fso.FileExists(chemin)
    If Not reprise Then
        If fso.GetFolder(dossier).Files.Count >= 1000 Then Err.Raise vbObjectError + 1108, , "Trop de commandes en attente. Diagnostic requis avant de continuer."
        id = modFichiers.IdUnique(): temporaire = chemin & "." & id & ".tmp"
        Set ts = fso.CreateTextFile(temporaire, False, False): ts.Write id: ts.Close
        ' Rename atomique sans ecrasement. Deux instances ne peuvent creer
        ' deux identifiants pour la meme commande ; la perdante doit reprendre.
        On Error Resume Next
        fso.MoveFile temporaire, chemin
        If Err.Number <> 0 Then
            Err.Clear: fso.DeleteFile temporaire: reprise = True
        End If
        On Error GoTo 0
    End If
    Set ts = fso.OpenTextFile(chemin, 1, False, 0): id = Trim$(ts.ReadAll): ts.Close
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[A-Za-z0-9_-]{16,100}$"
    If Not re.Test(id) Then Err.Raise vbObjectError + 1108, , "Identifiant durable invalide. Aucune commande envoyee."
    ObtenirCommandeLocale = id
End Function
