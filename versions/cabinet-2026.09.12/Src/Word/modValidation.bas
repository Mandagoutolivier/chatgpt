Attribute VB_Name = "modValidation"
Option Explicit
' =====================================================================
' modValidation - Validation du courrier par le medecin (Ctrl+Alt+V) :
'  1. enregistre le docx + un PDF dans le dossier du patient
'  2. depose un fichier-drapeau dans Echange\AEnvoyer\ ; le poste
'     secretaire le detecte, imprime, choisit l'acte, alimente le
'     journal comptable et la feuille de soins.
' =====================================================================

Public Sub ValiderCourrier()
    modPowerMicUnifie.Unifie_D_Finaliser
End Sub
' Validation d'un document rattache a un patient : docx + PDF dans le
' dossier du patient, drapeau pour le secretariat. Renvoie le type de
' courrier. silencieux=True : aucun message (lettres derivees automatiques).
Public Function ValiderDocument(ByVal doc As Document, ByVal silencieux As Boolean) As String
    Dim pat As Object, cor As Object
    Dim dossier As String, base As String, cheminDocx As String, cheminPdf As String
    Dim d As Object, typeCourrier As String, consultationID As String, dateActe As String, publicationID As String, dateValidation As String

    If Trim$(modIntegrationUnifie.VariableDoc(doc, "RelectureValidee")) <> "1" Then Err.Raise vbObjectError + 1127, , "Relecture medicale requise."
    Set pat = modIntegrationUnifie.PatientVerifie(doc)
    If pat Is Nothing Then Err.Raise vbObjectError + 520, "modValidation", _
        "Document sans patient rattache : appuyez sur F6 (ou Ctrl+Alt+P) pour choisir le patient, puis validez."
    Set cor = modApiConfiguration.CorrespondantDuDocument(doc)
    If cor Is Nothing Then Err.Raise vbObjectError + 1128, , "Destinataire principal absent."
    If UCase$(Trim$(CStr(cor("ID")))) = "A_COMPLETER" Then Err.Raise vbObjectError + 1128, , "Destinataire principal a completer."

    typeCourrier = VariableDoc(doc, "TypeCourrier")
    If Len(typeCourrier) = 0 Then typeCourrier = "courrier"

    ' Identifiant stable du document : attribue a la premiere validation et
    ' conserve dans le document. Une revalidation apres correction produit
    ' une nouvelle VERSION du meme acte, jamais un acte supplementaire.
    consultationID = VariableDoc(doc, "ConsultationID")
    If Len(Trim$(consultationID)) = 0 Then Err.Raise vbObjectError + 1129, , "Consultation NAS obligatoire."
    dateActe = VariableDoc(doc, "DateActe")
    If Not modTexte.DateFrValide(dateActe) Then Err.Raise vbObjectError + 1129, , "Date de consultation obligatoire."

    publicationID = Trim$(modIntegrationUnifie.VariableDoc(doc, "PublicationID"))
    If Len(publicationID) = 0 Then
        publicationID = modFichiers.IdUnique()
        modIntegrationUnifie.FixerVariable doc, "PublicationID", publicationID
        modIntegrationUnifie.FixerVariable doc, "DateValidation", Format$(Now, "dd/mm/yyyy hh:nn:ss")
    End If
    dateValidation = modIntegrationUnifie.VariableDoc(doc, "DateValidation")
    dossier = modPatient.DossierPatient(pat)
    base = Format$(modTexte.DateFr(dateActe), "yyyy-mm-dd") & " " & _
           modFichiers.NomFichierSur(typeCourrier) & " " & consultationID
    cheminDocx = dossier & "\" & base & "_" & publicationID & ".docx"
    cheminPdf = Left$(cheminDocx, Len(cheminDocx) - 4) & "pdf"

    If Not modEtatCourrier.PublicationPrete(publicationID) Then
        doc.SaveAs2 cheminDocx, 12
        doc.ExportAsFixedFormat cheminPdf, 17
        Dim sortie As String
        If Not SD_CopierRevisionFinale(doc, cheminDocx, publicationID, sortie) Then Err.Raise vbObjectError + 1130, , "Copie de sortie interrompue."
        modIntegrationUnifie.FixerVariable doc, "PublicationPreparee", publicationID
        ' Ne pas resauver les binaires prepares : PDF et DOCX viennent de la
        ' meme revision. Le drapeau local sert seulement a reprendre l envoi.
        modEtatCourrier.EcrireProtege publicationID, "-publication", "preparee"
    End If

    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = pat("ID")
    d("Nom") = pat("Nom")
    d("Prenom") = pat("Prenom")
    d("DDN") = modTexte.DdnPatient(pat)
    d("TypeCourrier") = typeCourrier
    d("ConsultationID") = consultationID
    d("SeanceID") = consultationID
    d("DateActe") = dateActe
    d("RdvID") = VariableDoc(doc, "RdvID")
    d("AnneeAgenda") = VariableDoc(doc, "AnneeAgenda")
    If Not cor Is Nothing Then d("DestinataireID") = cor("ID")
    d("CheminDocx") = cheminDocx
    d("CheminPdf") = cheminPdf
    d("DateValidation") = dateValidation
    d("PublicationID") = publicationID
    d("Poste") = Environ$("COMPUTERNAME")
    Dim k As Variant, resultat As Object
    For Each k In Array("Nom", "Prenom", "DDN", "Sexe")
        d("Patient_" & k) = CStr(pat(k))
    Next k
    d("Relu") = True
    Set resultat = modServiceNas.Appeler("publish", d)


    ' La copie secondaire est facultative ([SORTIE] ExportActif=1).
    ' Si activee, son contenu a ete verifie avant publication ; son nom
    ' porte uniquement l identifiant de publication.

    modLog.LogInfo "Courrier valide : " & cheminDocx & " (consultation " & consultationID & ")"
    ValiderDocument = typeCourrier
End Function


' UNE TOUCHE (bouton D du PowerMic) : corrige le courrier, ajoute les
' lettres de demande a la suite, enregistre (dossier patient + dossier de
' sortie) et transmet au secretariat.
Public Sub FinaliserCourrier()
    modPowerMicUnifie.Unifie_D_Finaliser
End Sub
' Chemin libre : base.ext, puis "base v2.ext", "base v3.ext"... Les
' corrections successives sont conservees au lieu de s'ecraser.

Private Function VariableDoc(ByVal doc As Document, ByVal nom As String) As String
    On Error Resume Next
    VariableDoc = doc.Variables(nom).Value
    Err.Clear
End Function

Private Sub DefinirVariableDoc(ByVal doc As Document, ByVal nom As String, ByVal valeur As String)
    modIntegrationUnifie.FixerVariable doc, nom, valeur
End Sub
