Attribute VB_Name = "modEcg"
Option Explicit
' =====================================================================
' modEcg - Commande "Envoyer a l'ECG" (Ctrl+Alt+G / voix) :
' ecrit l'identite du patient (du courrier actif, sinon via le
' selecteur) dans le fichier GDT surveille par Resting12Lead.
' Sur le poste ECG, il ne reste qu'a cliquer "Nouveau Patient" :
' la fiche arrive pre-remplie.
' =====================================================================

Public Sub EnvoyerECG()
    On Error GoTo Erreur
    Dim pat As Object, chemin As String
    Set pat = modIntegrationUnifie.PatientVerifie(ActiveDocument)
    If pat Is Nothing Then Exit Sub
    chemin = modGdt.EcrireGdtPatient(pat)
    MsgBox "Identite envoyee a l'ECG : " & pat("Prenom") & " " & pat("Nom") & vbCrLf & _
           "Dans Resting12Lead, cliquez simplement 'Nouveau Patient'.", _
           vbInformation, "Cabinet - ECG"
    Exit Sub
Erreur:
    Dim descErr As String
    descErr = Err.Description
    modLog.LogErreur "EnvoyerECG : " & descErr
    MsgBox "Envoi ECG impossible : " & descErr, vbExclamation, "Cabinet - ECG"
End Sub
