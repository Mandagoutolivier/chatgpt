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
' infos : dictionnaire du drapeau (PatientID, Nom, Prenom, DDN, NIR).
' Renvoie le SeanceID.
Public Function EnregistrerSeance(ByVal infos As Object, ByVal ActesChoisis As Collection, _
                                  ByVal modePaiement As String, ByVal tiersPayant As Boolean, _
                                  ByVal fdsImprimee As Boolean, Optional ByRef dejaEnregistree As Boolean = False) As String
    Dim seanceID As String, lignes As Collection, a As Object
    seanceID = ValeurOuVide(infos, "ConsultationID")
    If Len(seanceID) = 0 Then seanceID = ValeurOuVide(infos, "SeanceID")
    If Len(seanceID) = 0 Then Err.Raise vbObjectError + 640, "modActes", "Courrier sans identifiant de consultation : rattachez-le avant facturation."
    If ActesChoisis.Count = 0 Then Err.Raise vbObjectError + 641, "modActes", "Aucun acte choisi."
    If Not modTexte.DateFrValide(ValeurOuVide(infos, "DateActe")) Then Err.Raise vbObjectError + 642, "modActes", "Date de l acte absente ou invalide."
    Dim codes As Object
    Set codes = CreateObject("Scripting.Dictionary")
    codes.CompareMode = 1
    Set lignes = New Collection
    For Each a In ActesChoisis
        VerifierActeUnique codes, CStr(a("Code")), CStr(a("Tarif"))
        lignes.Add LigneJournal(seanceID, infos, a("Code"), a("Tarif"), modePaiement, tiersPayant, fdsImprimee)
        If Len(a("CodeAssocie")) > 0 Then
            VerifierActeUnique codes, CStr(a("CodeAssocie")), CStr(a("TarifAssocie"))
            lignes.Add LigneJournal(seanceID, infos, a("CodeAssocie"), a("TarifAssocie"), _
                                    modePaiement, tiersPayant, fdsImprimee)
        End If
    Next a
    Dim annee As Long
    annee = Year(modTexte.DateFr(CStr(infos("DateActe"))))
    modBaseIO.CreerClasseurSiAbsent modConfig.FichierJournal(annee), "JOURNAL", modJournal.EntetesJournal()
    dejaEnregistree = Not modBaseIO.AjouterSeanceUnique(modConfig.FichierJournal(annee), lignes, seanceID)
    EnregistrerSeance = seanceID
End Function

Private Function LigneJournal(ByVal seanceID As String, ByVal infos As Object, _
                              ByVal code As String, ByVal tarif As String, _
                              ByVal modePaiement As String, ByVal tiersPayant As Boolean, _
                              ByVal fdsImprimee As Boolean) As Object
    Dim d As Object, paye As Boolean
    paye = Not tiersPayant And (LCase$(modePaiement) <> "impaye" And LCase$(modePaiement) <> "impayé")
    Set d = CreateObject("Scripting.Dictionary")
    d("Date") = Format$(modTexte.DateFr(ValeurOuVide(infos, "DateActe")), "dd/mm/yyyy")
    d("SeanceID") = seanceID
    d("PatientID") = ValeurOuVide(infos, "PatientID")
    d("Nom") = ValeurOuVide(infos, "Nom")
    d("Prenom") = ValeurOuVide(infos, "Prenom")
    d("DDN") = ValeurOuVide(infos, "DDN")
    d("NIR") = ValeurOuVide(infos, "NIR")
    d("CodeActe") = code
    d("Montant") = Replace(tarif, ",", ".")
    d("ModePaiement") = modePaiement
    d("TiersPayant") = IIf(tiersPayant, "O", "N")
    d("Paye") = IIf(paye, "O", "N")
    d("DateEncaissement") = IIf(paye, Format$(Date, "dd/mm/yyyy"), "")
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

' Reprise apres un echec d'impression : utilise les montants deja enregistres.
Public Sub ImprimerSeanceEnregistree(ByVal infos As Object, ByVal seanceID As String)
    Dim journal As String, ligne As Object, lignes As New Collection, copie As Object, cle As Variant, impression As Object
    journal = modConfig.FichierJournal(Year(modTexte.DateFr(CStr(infos("DateActe")))))
    Set copie = CreateObject("Scripting.Dictionary")
    For Each cle In infos.Keys
        copie(cle) = infos(cle)
    Next cle
    For Each ligne In modBaseIO.LireTableX(journal, "JOURNAL", "SeanceID")
        If CStr(ligne("SeanceID")) = seanceID Then
            If CStr(ligne("PatientID")) <> CStr(infos("PatientID")) Then Err.Raise vbObjectError + 646, , "Consultation rattachee a un autre patient dans le journal."
            copie("DateActe") = ligne("Date")
            copie("Nom") = ligne("Nom"): copie("Prenom") = ligne("Prenom")
            copie("DDN") = ligne("DDN"): copie("NIR") = ligne("NIR")
            Set impression = CreateObject("Scripting.Dictionary")
            impression("CodeActe") = ligne("CodeActe")
            impression("Montant") = ligne("Montant")
            lignes.Add impression
        End If
    Next ligne
    If lignes.Count = 0 Then Err.Raise vbObjectError + 647, , "Consultation absente du journal."
    modCerfaPrint.ImprimerFeuille copie, lignes
    modBaseIO.MarquerFeuilleImprimee journal, seanceID
End Sub
