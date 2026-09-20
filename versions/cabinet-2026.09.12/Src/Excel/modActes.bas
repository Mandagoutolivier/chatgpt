Attribute VB_Name = "modActes"
Option Explicit
' =====================================================================
' modActes - Nomenclature des actes (Config\Nomenclature.xlsx, editable
' par la secretaire) et enregistrement d'une seance au journal.
' =====================================================================

Public Function Nomenclature() As Collection
    Dim tous As Collection, actifs As Collection, a As Object
    Set tous = modBaseIO.LireTableX(modConfig.FichierNomenclature(), "ACTES", "Code")
    Set actifs = New Collection
    For Each a In tous
        If a("Actif") <> "0" Then actifs.Add a
    Next a
    Set Nomenclature = actifs
End Function

Public Function ActeParCode(ByVal code As String) As Object
    Dim a As Object
    For Each a In Nomenclature()
        If a("Code") = code Then Set ActeParCode = a: Exit Function
    Next a
    Set ActeParCode = Nothing
End Function

' Enregistre une seance au journal comptable : une ligne par acte (les
' actes associes - ex ECG avec la consultation - generent leur ligne).
' tiersPayant accepte un booleen historique ou un dictionnaire CodeActe -> Boolean.
Public Function EnregistrerSeance(ByVal infos As Object, ByVal ActesChoisis As Collection, _
                                  ByVal modePaiement As String, ByVal tiersPayant As Variant, _
                                  ByVal fdsImprimee As Boolean, Optional ByRef dejaEnregistree As Boolean = False) As String
    Dim seanceID As String, lignes As Collection
    seanceID = ValeurOuVide(infos, "ConsultationID")
    If Len(seanceID) = 0 Then seanceID = ValeurOuVide(infos, "SeanceID")
    If Len(seanceID) = 0 Then Err.Raise vbObjectError + 640, "modActes", "Courrier sans identifiant de consultation : rattachez-le avant facturation."
    Set lignes = ConstruireLignesFacturation(infos, ActesChoisis, modePaiement, tiersPayant, fdsImprimee)
    Dim annee As Long, resultat As Object
    annee = Year(modTexte.DateFr(CStr(infos("DateActe"))))
    modBaseIO.CreerClasseurSiAbsent modConfig.FichierJournal(annee), "JOURNAL", modJournal.EntetesJournal()
    Set resultat = modBaseIO.EnregistrerActes(lignes, seanceID, ValeurOuVide(infos, "PublicationID"))
    dejaEnregistree = Not CBool(resultat("ajoute"))
    If CBool(resultat("selection_differente")) Then Err.Raise vbObjectError + 648, , "Les actes coches different des actes deja enregistres. Consultez la seance avant de poursuivre ; aucune modification comptable effectuee."
    EnregistrerSeance = seanceID
End Function

Public Function ConstruireLignesFacturation(ByVal infos As Object, ByVal ActesChoisis As Collection, _
                                             ByVal modePaiement As String, ByVal tiersPayant As Variant, _
                                             ByVal fdsImprimee As Boolean) As Collection
    Dim seanceID As String, lignes As New Collection, a As Object, codes As Object
    seanceID = ValeurOuVide(infos, "ConsultationID")
    If Len(seanceID) = 0 Then seanceID = ValeurOuVide(infos, "SeanceID")
    If Len(seanceID) = 0 Then Err.Raise vbObjectError + 640, "modActes", "Courrier sans identifiant de consultation : rattachez-le avant facturation."
    If ActesChoisis.Count = 0 Then Err.Raise vbObjectError + 641, "modActes", "Aucun acte choisi."
    If Not modTexte.DateFrValide(ValeurOuVide(infos, "DateActe")) Then Err.Raise vbObjectError + 642, "modActes", "Date de l acte absente ou invalide."
    Set codes = CreateObject("Scripting.Dictionary"): codes.CompareMode = 1
    For Each a In ActesChoisis
        Dim organisme As Boolean
        organisme = TiersPayantPourCode(tiersPayant, CStr(a("Code")))
        VerifierActeUnique codes, CStr(a("Code")), CStr(a("Tarif"))
        lignes.Add LigneJournal(seanceID, infos, a("Code"), a("Tarif"), ChoixLibelleCerfa(a, CStr(a("Code"))), modePaiement, _
                                organisme, fdsImprimee)
        If Len(a("CodeAssocie")) > 0 Then
            VerifierActeUnique codes, CStr(a("CodeAssocie")), CStr(a("TarifAssocie"))
            lignes.Add LigneJournal(seanceID, infos, a("CodeAssocie"), a("TarifAssocie"), CStr(a("CodeAssocie")), modePaiement, _
                                    organisme, fdsImprimee)
        End If
    Next a
    Set ConstruireLignesFacturation = lignes
End Function

Public Function TiersPayantPourCode(ByVal repartition As Variant, ByVal code As String) As Boolean
    If VarType(repartition) = vbBoolean Then
        TiersPayantPourCode = CBool(repartition)
        Exit Function
    End If
    If Not IsObject(repartition) Then Err.Raise vbObjectError + 650, "modActes", "Repartition du tiers payant invalide."
    If TypeName(repartition) <> "Dictionary" Then Err.Raise vbObjectError + 650, "modActes", "Repartition du tiers payant invalide."
    If Not repartition.Exists(code) Then Err.Raise vbObjectError + 650, "modActes", "Payeur absent pour l acte " & code & "."
    If VarType(repartition(code)) <> vbBoolean Then Err.Raise vbObjectError + 650, "modActes", "Payeur invalide pour l acte " & code & "."
    TiersPayantPourCode = CBool(repartition(code))
End Function

Private Function LigneJournal(ByVal seanceID As String, ByVal infos As Object, _
                              ByVal code As String, ByVal tarif As String, _
                              ByVal codeCerfa As String, _
                              ByVal modePaiement As String, ByVal tiersPayant As Boolean, _
                              ByVal fdsImprimee As Boolean) As Object
    Dim d As Object, paye As Boolean
    paye = Not tiersPayant And (LCase$(modePaiement) <> "impaye" And LCase$(modePaiement) <> "impayé")
    Set d = CreateObject("Scripting.Dictionary")
    d("Date") = Format$(modTexte.DateFr(ValeurOuVide(infos, "DateActe")), "dd""/""mm""/""yyyy")
    d("SeanceID") = seanceID
    d("PatientID") = ValeurOuVide(infos, "PatientID")
    d("Nom") = ValeurOuVide(infos, "Nom")
    d("Prenom") = ValeurOuVide(infos, "Prenom")
    d("DDN") = ValeurOuVide(infos, "DDN")
    d("NIR") = ValeurOuVide(infos, "NIR")
    d("CodeActe") = code
    d("CodeCerfa") = codeCerfa
    d("Montant") = Replace(tarif, ",", ".")
    d("ModePaiement") = modePaiement
    d("TiersPayant") = IIf(tiersPayant, "O", "N")
    d("Paye") = IIf(paye, "O", "N")
    d("DateEncaissement") = IIf(paye, Format$(Date, "dd""/""mm""/""yyyy"), "")
    d("FeuilleSoinsImprimee") = "N"  ' mis a O uniquement apres retour de PrintOut
    Set LigneJournal = d
End Function

Private Function ValeurOuVide(ByVal dict As Object, ByVal cle As String) As String
    If dict.Exists(cle) Then ValeurOuVide = dict(cle) Else ValeurOuVide = ""
End Function

' Lignes pour la feuille de soins : un element par acte, associes compris
Public Function LignesPourImpression(ByVal ActesChoisis As Collection) As Collection
    Dim col As New Collection, a As Object, d As Object
    For Each a In ActesChoisis
        Set d = CreateObject("Scripting.Dictionary")
        d("CodeActe") = ChoixLibelleCerfa(a, a("Code"))
        d("Montant") = a("Tarif")
        col.Add d
        If Len(a("CodeAssocie")) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            d("CodeActe") = a("CodeAssocie")
            d("Montant") = a("TarifAssocie")
            col.Add d
        End If
    Next a
    Set LignesPourImpression = col
End Function

Private Function ChoixLibelleCerfa(ByVal a As Object, ByVal defaut As String) As String
    If a.Exists("LibelleCerfa") Then
        If Len(a("LibelleCerfa")) > 0 Then
            ChoixLibelleCerfa = a("LibelleCerfa")
            Exit Function
        End If
    End If
    ChoixLibelleCerfa = defaut
End Function

' Total d'une selection d'actes (associes compris)
Public Function TotalActes(ByVal ActesChoisis As Collection) As Double
    Dim a As Object, total As Double
    For Each a In ActesChoisis
        total = total + Val(Replace(a("Tarif"), ",", "."))
        If Len(a("CodeAssocie")) > 0 Then total = total + Val(Replace(a("TarifAssocie"), ",", "."))
    Next a
    TotalActes = total
End Function

Private Sub VerifierActeUnique(ByVal codes As Object, ByVal code As String, ByVal tarif As String)
    Dim re As Object
    If Len(Trim$(code)) = 0 Then Err.Raise vbObjectError + 643, "modActes", "Code acte vide."
    If codes.Exists(code) Then Err.Raise vbObjectError + 644, "modActes", "Acte selectionne deux fois (direct ou associe) : " & code
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9]+([.,][0-9]{1,2})?$"
    If Not re.Test(Trim$(tarif)) Then Err.Raise vbObjectError + 645, "modActes", "Tarif invalide pour " & code
    codes.Add code, True
End Sub

Public Function ExisteSeanceEnregistree(ByVal seanceID As String) As Boolean
    Dim p As Object, r As Object, items As Collection
    If Len(Trim$(seanceID)) = 0 Then Exit Function
    Set p = modServiceNas.Parametres(): p("id") = seanceID: p("limit") = 1
    Set r = modServiceNas.Appeler("journal.read", p)
    Set items = modServiceNas.ItemsValides(r)
    ExisteSeanceEnregistree = (items.Count > 0)
End Function

Public Function DonneesImpressionFigees(ByVal sauvegardees As Collection, ByVal patientID As String, ByRef lignesImpression As Collection) As Object
    Dim identite As Object, ligne As Object, sortie As Object, cle As Variant
    Dim premiere As Boolean, champs As Variant, valeur As String, assureDistinct As Boolean
    If Len(Trim$(patientID)) = 0 Then Err.Raise vbObjectError + 646, , "Patient de la seance absent."
    Set identite = CreateObject("Scripting.Dictionary"): identite.CompareMode = 1
    Set lignesImpression = New Collection
    champs = Array("Nom", "Prenom", "DDN", "NIR", "AssureNom", "AssurePrenom", "AssureDDN", "AssureNIR")
    premiere = True
    For Each ligne In sauvegardees
        If Not ligne.Exists("PatientID") Then Err.Raise vbObjectError + 646, , "Patient de la seance absent."
        If CStr(ligne("PatientID")) <> patientID Then Err.Raise vbObjectError + 646, , "Seance rattachee a un autre patient."
        If Not ligne.Exists("Date") Then Err.Raise vbObjectError + 649, , "Date figee absente."
        If Not ligne.Exists("CodeActe") Then Err.Raise vbObjectError + 649, , "Acte fige absent."
        If Not ligne.Exists("Montant") Then Err.Raise vbObjectError + 649, , "Montant fige absent."
        If premiere Then
            identite("PatientID") = patientID: identite("DateActe") = CStr(ligne("Date"))
            For Each cle In champs
                If ligne.Exists(CStr(cle)) Then identite(CStr(cle)) = CStr(ligne(CStr(cle))) Else identite(CStr(cle)) = ""
            Next cle
            premiere = False
        Else
            If CStr(ligne("Date")) <> CStr(identite("DateActe")) Then Err.Raise vbObjectError + 649, , "Dates figees incoherentes."
            For Each cle In champs
                valeur = "": If ligne.Exists(CStr(cle)) Then valeur = CStr(ligne(CStr(cle)))
                If valeur <> CStr(identite(CStr(cle))) Then Err.Raise vbObjectError + 649, , "Identites figees incoherentes."
            Next cle
        End If
        Set sortie = CreateObject("Scripting.Dictionary")
        sortie("CodeActe") = CStr(ligne("CodeActe"))
        If ligne.Exists("CodeCerfa") Then
            If Len(Trim$(CStr(ligne("CodeCerfa")))) > 0 Then sortie("CodeActe") = CStr(ligne("CodeCerfa"))
        End If
        sortie("Montant") = CStr(ligne("Montant"))
        lignesImpression.Add sortie
    Next ligne
    If premiere Then Err.Raise vbObjectError + 649, , "Seance figee sans ligne."
    If Len(Trim$(CStr(identite("Nom")))) = 0 Then Err.Raise vbObjectError + 649, , "Nom fige absent."
    If Len(Trim$(CStr(identite("Prenom")))) = 0 Then Err.Raise vbObjectError + 649, , "Prenom fige absent."
    If Len(Trim$(CStr(identite("DDN")))) = 0 Then Err.Raise vbObjectError + 649, , "Naissance figee absente."
    ' Toute donnee assuree, meme une DDN isolee, interdit le repli sur le patient.
    assureDistinct = Len(Trim$(CStr(identite("AssureNom")) & CStr(identite("AssurePrenom")) & CStr(identite("AssureDDN")) & CStr(identite("AssureNIR")))) > 0
    If assureDistinct Then
        If Len(Trim$(CStr(identite("AssureNom")))) = 0 Or Len(Trim$(CStr(identite("AssurePrenom")))) = 0 Then Err.Raise vbObjectError + 649, , "Identite de l assure figee incomplete."
        If Not modTexte.DateFrValide(CStr(identite("AssureDDN"))) Then Err.Raise vbObjectError + 649, , "Naissance de l assure figee invalide."
    End If
    Set DonneesImpressionFigees = identite
End Function

' Une impression demandee ne permet jamais de clore le courrier tant que
' le papier et sa confirmation serveur ne sont pas tous deux acquis.
Public Function TraitementPeutEtreCloture(ByVal impressionDemandee As Boolean, ByVal papierConfirme As Boolean) As Boolean
    TraitementPeutEtreCloture = (Not impressionDemandee) Or papierConfirme
End Function

' Reprise apres un echec d'impression : utilise exclusivement l'identite et
' les montants figes lors de la premiere facturation. False = annulation ou
' papier non confirme ; le courrier doit rester en attente, sans ack.
' impressionEnvoyee distingue un nouvel envoi de la confirmation d'un papier anterieur.
Public Function ImprimerSeanceEnregistree(ByVal infos As Object, ByVal seanceID As String, _
                                         Optional ByRef impressionEnvoyee As Boolean = False) As Boolean
    Dim saved As Object, p As Object, tentative As Object, r As Object
    Dim copie As Object, lignes As Collection, sauvegardees As Collection, confirme As Boolean
    impressionEnvoyee = False
    Set saved = modServiceNas.CommandeID("billing.get", seanceID)
    Set sauvegardees = saved("lignes")
    Set copie = DonneesImpressionFigees(sauvegardees, ValeurOuVide(infos, "PatientID"), lignes)
    If CStr(saved("impression_etat")) = "inconnue" Then
        Dim decision As VbMsgBoxResult
        decision = MsgBox("Verifiez la feuille deja demandee pour " & CStr(copie("Nom")) & " " & CStr(copie("Prenom")) & " du " & CStr(copie("DateActe")) & "." & vbCrLf & "Est-elle sortie correctement ?" & vbCrLf & "Oui : confirmer cette feuille. Non : proposer une reimpression. Annuler : conserver en attente.", vbYesNoCancel + vbQuestion, "Resultat papier a verifier")
        If decision = vbCancel Then Exit Function
        If decision = vbYes Then
            Set p = modServiceNas.Parametres(): p("id") = seanceID
            p("tentative") = CStr(saved("tentative")): p("confirmee") = True
            Set r = modServiceNas.Appeler("printed", p)
            ImprimerSeanceEnregistree = True
            Exit Function
        End If
    End If
    modCerfaPrint.VerifierAvantReimpression copie, lignes
    confirme = CStr(saved("impression_etat")) <> "actes_enregistres"
    If confirme Then
        If MsgBox("Une impression a deja ete demandee. Verifiez la sortie papier avant de poursuivre." & vbCrLf & "Confirmez-vous une nouvelle impression des actes enregistres ?", vbYesNo + vbExclamation, "Reimpression explicite") <> vbYes Then Exit Function
    End If
    Set p = modServiceNas.Parametres(): p("id") = seanceID: p("reimpression_confirmee") = confirme
    Set tentative = modServiceNas.Appeler("print.request", p)
    modCerfaPrint.ImprimerFeuille copie, lignes
    impressionEnvoyee = True
    If MsgBox("La feuille est-elle sortie correctement sur papier ?" & vbCrLf & "Non conserve un resultat inconnu et le courrier en attente.", vbYesNo + vbQuestion, "Verifier la feuille imprimee") <> vbYes Then Exit Function
    Set p = modServiceNas.Parametres(): p("id") = seanceID
    p("tentative") = CStr(tentative("tentative")): p("confirmee") = True
    Set r = modServiceNas.Appeler("printed", p)
    ImprimerSeanceEnregistree = True
End Function
