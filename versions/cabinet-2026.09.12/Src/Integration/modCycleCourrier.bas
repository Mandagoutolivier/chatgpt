Attribute VB_Name = "modCycleCourrier"
Option Explicit
' Proprietaire unique du document source pendant la correction.
Private mDocumentSource As Document
Private mTraitementEnCours As Boolean

Public Function DocumentSource() As Document
    Set DocumentSource = mDocumentSource
End Function

Public Function CycleEnCours() As Boolean
    CycleEnCours = mTraitementEnCours
End Function

Public Sub ExecuterCycleCourrier()
    Dim docPrincipal As Document, rngCorps As Range, premier As Range, etat As Object, reponse As Object
    Dim original As String, source As String, corpsCorrige As String, numero As Long, description As String
    Dim ancienEcran As Boolean, hashCourant As String, verrou As Integer, cor As Object
    If mTraitementEnCours Then Exit Sub
    mTraitementEnCours = True
    ancienEcran = Application.ScreenUpdating
    On Error GoTo Echec
    If Documents.Count = 0 Then Err.Raise vbObjectError + 960, , "Aucun courrier ouvert."
    Set docPrincipal = ActiveDocument
    modGlobals.ReinitialiserContexteTraitement
    modProdRapide.PR_ReinitialiserSuiviMultipage
    Set mDocumentSource = docPrincipal
    modIntegrationUnifie.InitialiserPatientProd docPrincipal
    Set cor = modControleCourrier.AssurerDestinataire(docPrincipal)
    If cor Is Nothing Then GoTo Sortie
    modEtatCourrier.PrendreVerrou docPrincipal, verrou
    modControleCourrier.InvaliderRelecture docPrincipal
    modIntegrationUnifie.SauvegarderBrouillon docPrincipal
    If Not modAnonymisation.LocaliserCorpsCourrier(docPrincipal, rngCorps, premier) Then Err.Raise vbObjectError + 960, , "Corps du courrier introuvable."
    Set gPlageOriginale = rngCorps.Duplicate
    original = rngCorps.Text
    Set etat = modEtatCourrier.Charger(docPrincipal, original)
    hashCourant = modServiceNas.SHA256(original)
    If CBool(etat("corps_insere")) Then
        If etat.Exists("document_hash") Then
            If hashCourant <> CStr(etat("document_hash")) Then
                If MsgBox("Le corps a ete modifie depuis le dernier cycle. Conserver ce cycle et commencer une nouvelle correction du texte actuel ?", vbYesNo + vbQuestion, "Texte modifie") <> vbYes Then GoTo Sortie
                modIntegrationUnifie.FixerVariable docPrincipal, "CycleU1", ""
                Set etat = modEtatCourrier.Charger(docPrincipal, original)
            End If
        End If
    ElseIf hashCourant <> CStr(etat("source_hash")) Then
        If MsgBox("Le texte a change pendant cette reprise. Conserver le cycle precedent et commencer une nouvelle correction du texte actuel ?", vbYesNo + vbQuestion, "Nouveau cycle explicite") <> vbYes Then GoTo Sortie
        modIntegrationUnifie.FixerVariable docPrincipal, "CycleU1", ""
        Set etat = modEtatCourrier.Charger(docPrincipal, original)
    End If
    If etat.Exists("destinataire_id") Then
        If CStr(etat("destinataire_id")) <> CStr(cor("ID")) Then
            If MsgBox("Le destinataire a change. Conserver le cycle precedent et recommencer la correction du texte actuel avec ce destinataire ?", vbYesNo + vbQuestion, "Nouveau destinataire") <> vbYes Then GoTo Sortie
            modIntegrationUnifie.FixerVariable docPrincipal, "CycleU1", ""
            Set etat = modEtatCourrier.Charger(docPrincipal, original)
        End If
    End If
    etat("destinataire_id") = CStr(cor("ID"))
    modEtatCourrier.Sauver docPrincipal, etat
    original = modEtatCourrier.LireProtege(Trim$(modIntegrationUnifie.VariableDoc(docPrincipal, "CycleU1")), "-source")
    source = Replace(original, gPatient.NomComplet, gPatient.civilite & " " & MARQUEUR_PATIENT, 1, -1, vbTextCompare)
    If InStr(1, source, MARQUEUR_PATIENT, vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 960, , "Inserez l identite du patient avec C avant de finaliser."
    gTexteAnonymise = source
    If Len(CStr(etat("corps"))) = 0 Then
        modEtatCourrier.AvantAppel docPrincipal, etat, "corps"
        Set reponse = modOpenAI_v22_corrige.AppelerOpenAIStructure(source, modPrompts.ConstruirePromptReecritureMedicale(""))
        corpsCorrige = CStr(reponse("corps_courrier"))
        If InStr(1, corpsCorrige, MARQUEUR_PATIENT, vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 960, , "Marqueur patient absent du courrier corrige."
        corpsCorrige = modMiseEnPageFinale.MPF_NormaliserRetoursTexte(corpsCorrige)
        corpsCorrige = modProdRapide.PR_CorrigerReferentPatientOrientationCCN(corpsCorrige)
        etat("corps") = corpsCorrige: etat("appel") = ""
        modEtatCourrier.Sauver docPrincipal, etat
    End If
    corpsCorrige = CStr(etat("corps"))
    If Len(CStr(etat("reponse"))) = 0 Then
        If modPrompts.TexteContientDemandeExamenEligible(source) Then
            modEtatCourrier.AvantAppel docPrincipal, etat, "annexes"
            Set reponse = modOpenAI_v22_corrige.AppelerOpenAIStructure(corpsCorrige, modPrompts.ConstruirePromptDemandeExamenSeule(""))
            If reponse("demandes").Count = 0 Then Err.Raise vbObjectError + 436, , "Demande detectee dans la dictee mais aucune annexe recue."
        Else
            Set reponse = modServiceNas.Parametres(): Set reponse("demandes") = New Collection
        End If
        reponse("corps_courrier") = corpsCorrige
        etat("reponse") = modServiceNas.JsonValeur(reponse): etat("appel") = ""
        modEtatCourrier.Sauver docPrincipal, etat
    End If
    gReponseAPICabinetTest = CStr(etat("reponse"))
    gCorpsCorrigeAnonymiseCabinetTest = corpsCorrige
    Application.ScreenUpdating = False
    If Not CBool(etat("corps_insere")) Then
        If Not modInsertion.RemplacerCorpsOriginalParTexte(corpsCorrige) Then Err.Raise vbObjectError + 960, , "Insertion du courrier interrompue."
        modMiseEnPageFinale.MPF_AppliquerMiseEnPageCourrier docPrincipal
        If Not modAnonymisation.LocaliserCorpsCourrier(docPrincipal, rngCorps, premier) Then Err.Raise vbObjectError + 960, , "Corps introuvable apres insertion."
        etat("corps_insere") = True: etat("document_hash") = modServiceNas.SHA256(rngCorps.Text)
        docPrincipal.Save: modEtatCourrier.Sauver docPrincipal, etat
    End If
    gCorrectionCabinetTestValidee = True
    ' Reconstituer les pages a partir des resultats durables. Aucune nouvelle
    ' requete IA et aucune reutilisation d'une selection d'un autre document.
    modProdRapide.PR_SupprimerDemandesDejaAjoutees docPrincipal
    modProdRapide.PR_ReinitialiserSuiviMultipage
    modProdRapide.PR_AjouterDemandesAuDocumentPrincipal
    etat("reponse") = gReponseAPICabinetTest
    etat("annexes_ok") = gDemandesMultipagesAjouteesProdRapide
    modEtatCourrier.Sauver docPrincipal, etat
    docPrincipal.Save
    If Not gDemandesMultipagesAjouteesProdRapide Then Err.Raise vbObjectError + 436, , "Annexes incompletes. D reprend cette etape avec les textes deja recus."
    modMiseEnPageFinale.MPF_SecuriserToutesSignatures docPrincipal
    modGras.AppliquerGrasDocumentComplet docPrincipal
    If Not modAnonymisation.LocaliserCorpsCourrier(docPrincipal, rngCorps, premier) Then Err.Raise vbObjectError + 960, , "Corps introuvable apres mise en forme."
    etat("document_hash") = modServiceNas.SHA256(rngCorps.Text)
    modEtatCourrier.Sauver docPrincipal, etat
    modControleCourrier.PreparerRelecture docPrincipal, source, corpsCorrige
    docPrincipal.Save
Sortie:
    On Error Resume Next
    Set mDocumentSource = Nothing: Set gPlageOriginale = Nothing
    gTexteAnonymise = "": gTexteCorrige = "": gReponseAPICabinetTest = ""
    gCorpsCorrigeAnonymiseCabinetTest = "": gCorrectionCabinetTestValidee = False
    Application.ScreenUpdating = ancienEcran: Application.StatusBar = vbNullString
    modEtatCourrier.LibererVerrou verrou
    mTraitementEnCours = False
    On Error GoTo 0
    If numero <> 0 Then Err.Raise numero, "Correction", description
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    modLog.Diagnostic "correction", "echec", numero
    Resume Sortie
End Sub
