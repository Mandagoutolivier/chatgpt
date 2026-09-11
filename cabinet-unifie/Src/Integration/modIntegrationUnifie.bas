Attribute VB_Name = "modIntegrationUnifie"
Option Explicit

' Appelee par le moteur PROD(6) apres le SaveAs2 final et avant fermeture.
Public Sub TransmettreSecretariat(ByVal doc As Document, ByVal cheminDocx As String)
    On Error GoTo Erreur
    Dim pat As Object, d As Object, consultationID As String
    Set pat = modClaude.PatientDuDocument(doc)
    If pat Is Nothing Then Err.Raise vbObjectError + 950, "modIntegrationUnifie", _
        "Le courrier corrige n'est pas rattache a un patient."
    consultationID = VariableDoc(doc, "ConsultationID")
    If Len(consultationID) = 0 Then consultationID = modFichiers.IdUnique()
    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = pat("ID")
    d("Nom") = pat("Nom")
    d("Prenom") = pat("Prenom")
    d("DDN") = pat("DDN")
    d("TypeCourrier") = "consultation"
    d("ConsultationID") = consultationID
    d("SeanceID") = consultationID
    d("RdvID") = VariableDoc(doc, "RdvID")
    d("CheminDocx") = cheminDocx
    d("DateValidation") = Format$(Now, "dd/mm/yyyy hh:nn")
    d("Poste") = Environ$("COMPUTERNAME")
    modFichiers.EcrireDrapeau modConfig.Chemin("Echange") & "\AEnvoyer", _
        modFichiers.IdUnique() & "_" & pat("ID"), d
    modLog.LogInfo "Courrier PROD(6) transmis au secretariat : " & cheminDocx
    Exit Sub
Erreur:
    modLog.LogErreur "Transmission secretariat : " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function VariableDoc(ByVal doc As Document, ByVal nom As String) As String
    On Error Resume Next
    VariableDoc = doc.Variables(nom).Value
    Err.Clear
End Function
