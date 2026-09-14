Attribute VB_Name = "modEtatCourrier"
Option Explicit
Private mDossierTests As String
#If VBA7 Then
Private Type BLOB
    cbData As Long
    pbData As LongPtr
End Type
Private Declare PtrSafe Function CryptProtectData Lib "crypt32.dll" (ByRef source As BLOB, ByVal description As LongPtr, ByVal entropy As LongPtr, ByVal reserve As LongPtr, ByVal prompt As LongPtr, ByVal flags As Long, ByRef target As BLOB) As Long
Private Declare PtrSafe Function CryptUnprotectData Lib "crypt32.dll" (ByRef source As BLOB, ByVal description As LongPtr, ByVal entropy As LongPtr, ByVal reserve As LongPtr, ByVal prompt As LongPtr, ByVal flags As Long, ByRef target As BLOB) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef target As Any, ByVal source As LongPtr, ByVal length As LongPtr)
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal source As LongPtr, ByVal target As LongPtr, ByVal flags As Long) As Long
#Else
Private Type BLOB
    cbData As Long
    pbData As Long
End Type
Private Declare Function CryptProtectData Lib "crypt32.dll" (ByRef source As BLOB, ByVal description As Long, ByVal entropy As Long, ByVal reserve As Long, ByVal prompt As Long, ByVal flags As Long, ByRef target As BLOB) As Long
Private Declare Function CryptUnprotectData Lib "crypt32.dll" (ByRef source As BLOB, ByVal description As Long, ByVal entropy As Long, ByVal reserve As Long, ByVal prompt As Long, ByVal flags As Long, ByRef target As BLOB) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef target As Any, ByVal source As Long, ByVal length As Long)
Private Declare Function LocalFree Lib "kernel32" (ByVal memory As Long) As Long
Private Declare Function MoveFileExW Lib "kernel32" (ByVal source As Long, ByVal target As Long, ByVal flags As Long) As Long
#End If

Private Function CheminEtat(ByVal id As String, ByVal suffixe As String) As String
    Dim re As Object, dossier As String
    If suffixe <> "-source" And suffixe <> "-etat" And suffixe <> "-publication" Then Err.Raise vbObjectError + 1160
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[A-Za-z0-9_-]{16,100}$"
    If Not re.Test(id) Then Err.Raise vbObjectError + 1160, , "Identifiant de reprise invalide."
    dossier = Environ$("LOCALAPPDATA") & "\CabinetCardio\Reprises"
    If Len(mDossierTests) > 0 Then dossier = mDossierTests
    modFichiers.EnsureDossier dossier
    CheminEtat = dossier & "\" & id & suffixe & ".dpapi"
End Function

Private Function Transformer(ByVal donnees As Variant, ByVal chiffrer As Boolean) As Variant
    Dim entree As BLOB, sortie As BLOB, resultat() As Byte, ok As Long, numero As Long
    On Error GoTo Echec
    entree.cbData = UBound(donnees) + 1: entree.pbData = VarPtr(donnees(0))
    If chiffrer Then
        ok = CryptProtectData(entree, 0, 0, 0, 0, 1, sortie)
    Else
        ok = CryptUnprotectData(entree, 0, 0, 0, 0, 1, sortie)
    End If
    If ok = 0 Or sortie.cbData <= 0 Or sortie.cbData > 5000000 Then Err.Raise vbObjectError + 1161
    ReDim resultat(0 To sortie.cbData - 1)
    CopyMemory resultat(0), sortie.pbData, sortie.cbData
    Transformer = resultat
Sortie:
    If sortie.pbData <> 0 Then LocalFree sortie.pbData
    On Error GoTo 0
    If numero <> 0 Then Err.Raise vbObjectError + 1161, , "Etat de reprise illisible ou protection Windows indisponible. Conservez le brouillon et le meme compte Windows."
    Exit Function
Echec:
    numero = Err.Number: Resume Sortie
End Function

Public Sub EcrireProtege(ByVal id As String, ByVal suffixe As String, ByVal texte As String, Optional ByVal immuable As Boolean = False)
    Dim chemin As String, tmp As String, flux As Object, fso As Object, octets As Variant
    If Len(texte) > 2000000 Then Err.Raise vbObjectError + 1162, , "Etat de reprise trop volumineux."
    chemin = CheminEtat(id, suffixe)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If immuable And fso.FileExists(chemin) Then Err.Raise vbObjectError + 1162, , "La source originale du cycle est immuable."
    If fso.GetFolder(fso.GetParentFolderName(chemin)).Size > 268435456 Then Err.Raise vbObjectError + 1162, , "Espace de reprise borne atteint. Archivez les cycles termines avant de poursuivre."
    octets = Transformer(modServiceNas.OctetsUTF8(texte), True)
    tmp = chemin & "." & modFichiers.IdUnique() & ".tmp"
    Set flux = CreateObject("ADODB.Stream"): flux.Type = 1: flux.Open
    flux.Write octets: flux.SaveToFile tmp, 1: flux.Close
    If MoveFileExW(StrPtr(tmp), StrPtr(chemin), IIf(immuable, 8, 9)) = 0 Then
        modFichiers.SupprimerTemporaire tmp
        Err.Raise vbObjectError + 1162, , "Enregistrement atomique de reprise impossible."
    End If
End Sub

Public Function LireProtege(ByVal id As String, ByVal suffixe As String) As String
    Dim flux As Object, chemin As String, octets As Variant
    chemin = CheminEtat(id, suffixe)
    If Not modFichiers.FichierExiste(chemin) Then Err.Raise vbObjectError + 1163, , "Etat de reprise absent sur ce poste. Recuperer le cycle d origine avant de relancer une correction."
    Set flux = CreateObject("ADODB.Stream"): flux.Type = 1: flux.Open: flux.LoadFromFile chemin
    If flux.Size > 5000000 Then flux.Close: Err.Raise vbObjectError + 1162
    octets = flux.Read: flux.Close
    LireProtege = modServiceNas.TexteUTF8(Transformer(octets, False))
End Function

Public Function Charger(ByVal doc As Document, ByVal source As String) As Object
    Dim id As String, etat As Object, pat As Object
    Set pat = modIntegrationUnifie.PatientVerifie(doc)
    id = Trim$(modIntegrationUnifie.VariableDoc(doc, "CycleU1"))
    If Len(id) = 0 Then
        id = modFichiers.IdUnique()
        Set etat = modServiceNas.Parametres()
        etat("PatientID") = CStr(pat("ID")): etat("ConsultationID") = modIntegrationUnifie.VariableDoc(doc, "ConsultationID")
        etat("source_hash") = modServiceNas.SHA256(source): etat("appel") = "": etat("corps") = ""
        etat("reponse") = "": etat("corps_insere") = False: etat("annexes_ok") = False
        EcrireProtege id, "-source", source, True
        EcrireProtege id, "-etat", modServiceNas.JsonValeur(etat)
        modIntegrationUnifie.FixerVariable doc, "CycleU1", id
        doc.Save
    Else
        Set etat = modJson.JsonParse(LireProtege(id, "-etat"))
        If CStr(etat("PatientID")) <> CStr(pat("ID")) Or CStr(etat("ConsultationID")) <> modIntegrationUnifie.VariableDoc(doc, "ConsultationID") Then Err.Raise vbObjectError + 1164, , "Ce cycle appartient a une autre consultation."
    End If
    Set Charger = etat
End Function

Public Sub Sauver(ByVal doc As Document, ByVal etat As Object)
    EcrireProtege Trim$(modIntegrationUnifie.VariableDoc(doc, "CycleU1")), "-etat", modServiceNas.JsonValeur(etat)
End Sub

Public Sub AvantAppel(ByVal doc As Document, ByVal etat As Object, ByVal phase As String)
    If Len(CStr(etat("appel"))) > 0 Then
        If MsgBox("Le resultat du precedent appel de correction est inconnu. Une nouvelle requete peut etre facturee par le fournisseur." & vbCrLf & "Confirmez-vous la relance de cette seule etape ?", vbYesNo + vbExclamation, "Reprise de la correction") <> vbYes Then Err.Raise vbObjectError + 1165, , "Reprise suspendue ; source et resultat deja recu conserves."
    End If
    etat("appel") = phase: Sauver doc, etat
End Sub

Public Function PublicationPrete(ByVal id As String) As Boolean
    If Not modFichiers.FichierExiste(CheminEtat(id, "-publication")) Then Exit Function
    PublicationPrete = (LireProtege(id, "-publication") = "preparee")
End Function

Public Sub DefinirDossierTests(ByVal dossier As String)
    mDossierTests = dossier
End Sub
