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

    Set pat = modIntegrationUnifie.PatientVerifie(doc)
    If pat Is Nothing Then Err.Raise vbObjectError + 520, "modValidation", _
        "Document sans patient rattache : appuyez sur F6 (ou Ctrl+Alt+P) pour choisir le patient, puis validez."
    Set cor = modClaude.CorrespondantDuDocument(doc)
    If cor Is Nothing Then
        ' destinataire dicte (saisie rapide) : reconnu dans la base si possible
        If Len(modCourrier.ReconnaitreDestinataire(doc)) > 0 Then Set cor = modClaude.CorrespondantDuDocument(doc)
        modCourrier.MettreEnFormeDestinataire doc
    End If

    typeCourrier = VariableDoc(doc, "TypeCourrier")
    If Len(typeCourrier) = 0 Then typeCourrier = "courrier"

    ' Identifiant stable du document : attribue a la premiere validation et
    ' conserve dans le document. Une revalidation apres correction produit
    ' une nouvelle VERSION du meme acte, jamais un acte supplementaire.
    consultationID = VariableDoc(doc, "ConsultationID")
    If Len(consultationID) = 0 Then
        consultationID = modFichiers.IdUnique()
        DefinirVariableDoc doc, "ConsultationID", consultationID
    End If

    ' Date REELLE de l'acte : celle de la consultation si elle est portee par
    ' le document, sinon celle de la premiere validation (figee ensuite).
    dateActe = VariableDoc(doc, "DateActe")
    If Not modTexte.DateFrValide(dateActe) Then
        dateActe = Format$(Date, "dd/mm/yyyy")
        DefinirVariableDoc doc, "DateActe", dateActe
    End If

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

    doc.SaveAs2 cheminDocx, 12                      ' wdFormatXMLDocument
    doc.ExportAsFixedFormat cheminPdf, 17           ' wdExportFormatPDF

    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = pat("ID")
    d("Nom") = pat("Nom")
    d("Prenom") = pat("Prenom")
    d("DDN") = modTexte.DdnPatient(pat)
    d("NIR") = pat("NIR")
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
    modFichiers.EcrireDrapeau modConfig.chemin("Echange") & "\AEnvoyer", _
                              consultationID & "_" & publicationID, d

    ' copie dans le dossier de sortie du cabinet ([SORTIE] Dossier, ex :
    ' \\DS224\home\sortiedragon) sous "NOM Prenom aammjjhhmm.docx", comme
    ' l'ancien modele : le secretariat y retrouve le fichier complet.
    ' La copie sortiedragon a deja ete verifiee par PROD avant publication.

    modLog.LogInfo "Courrier valide : " & cheminDocx & " (consultation " & consultationID & ")"
    ValiderDocument = typeCourrier
End Function

Private Sub CopierVersSortie(ByVal cheminDocx As String, ByVal pat As Object, ByVal typeCourrier As String)
    On Error Resume Next
    Dim dossier As String, dest As String
    dossier = modConfig.Config("SORTIE", "Dossier", "")
    If Len(dossier) = 0 Then Exit Sub
    If Left$(LCase$(typeCourrier), 7) = "demande" Then Exit Sub    ' les demandes sont dans le fichier principal
    modFichiers.EnsureDossier dossier
    dest = dossier & "\" & modFichiers.NomFichierSur(UCase$(pat("Nom")) & " " & pat("Prenom") & " " & Format$(Now, "yymmddhhnn")) & ".docx"
    FileCopy cheminDocx, dest
    If Err.Number <> 0 Then
        modLog.LogErreur "Copie vers la sortie impossible (" & dest & ") : " & Err.Description
    Else
        modLog.LogInfo "Copie de sortie : " & dest
    End If
End Sub

' UNE TOUCHE (bouton D du PowerMic) : corrige le courrier, ajoute les
' lettres de demande a la suite, enregistre (dossier patient + dossier de
' sortie) et transmet au secretariat.
Public Sub FinaliserCourrier()
    modPowerMicUnifie.Unifie_D_Finaliser
End Sub
' Chemin libre : base.ext, puis "base v2.ext", "base v3.ext"... Les
' corrections successives sont conservees au lieu de s'ecraser.
Private Function NomVersionne(ByVal dossier As String, ByVal base As String, _
                              ByVal ext As String) As String
    Dim chemin As String, n As Long
    chemin = dossier & "\" & base & "." & ext
    n = 1
    Do While Len(Dir$(chemin)) > 0
        n = n + 1
        chemin = dossier & "\" & base & " v" & n & "." & ext
        If n > 200 Then Exit Do
    Loop
    NomVersionne = chemin
End Function

Private Function VariableDoc(ByVal doc As Document, ByVal nom As String) As String
    On Error Resume Next
    VariableDoc = doc.Variables(nom).Value
    Err.Clear
End Function

Private Sub DefinirVariableDoc(ByVal doc As Document, ByVal nom As String, ByVal valeur As String)
    modIntegrationUnifie.FixerVariable doc, nom, valeur
End Sub
