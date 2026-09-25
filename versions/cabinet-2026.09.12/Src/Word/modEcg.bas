Attribute VB_Name = "modEcg"
Option Explicit
' =====================================================================
' modEcg - Commande "Envoyer a l'ECG" (Ctrl+Alt+G / voix) :
' ecrit l'identite du patient (du courrier actif, sinon via le
' selecteur) dans le fichier GDT surveille par Resting12Lead.
' Sur le poste ECG, F2 / Nouveau Patient importe l identite et la DDN.
' L age et le sexe restent a controler dans la fiche avant utilisation.
' =====================================================================

Public Sub EnvoyerECG()
    On Error GoTo Erreur
    Dim pat As Object, chemin As String
    Set pat = modIntegrationUnifie.PatientVerifie(ActiveDocument)
    If pat Is Nothing Then Exit Sub
    chemin = modGdt.EcrireGdtPatient(pat)
    MsgBox "Identite envoyee a l'ECG : " & pat("Prenom") & " " & pat("Nom") & vbCrLf & _
           "Dans Resting12Lead : F2 / Nouveau Patient, puis controlez l'ID et la DDN." & vbCrLf & _
           "Pour recalculer l'age : changez le jour dans le calendrier, puis retablissez exactement la DDN." & vbCrLf & _
           "Ce geste n'est pas automatise. Verifiez l'age obtenu et renseignez le sexe.", _
           vbInformation, "Cabinet - ECG"
    Exit Sub
Erreur:
    Dim descErr As String
    descErr = Err.Description
    modLog.LogErreur "EnvoyerECG : " & descErr
    MsgBox "Envoi ECG impossible : " & descErr, vbExclamation, "Cabinet - ECG"
End Sub
