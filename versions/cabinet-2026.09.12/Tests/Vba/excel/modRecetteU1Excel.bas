Attribute VB_Name = "modRecetteU1Excel"
Option Explicit
Private Const NOMBRE_ATTENDU_EXCEL As Long = 25

Public Function Executer() As String
    Dim col As Collection, cas As Variant, numero As Long, reussis As Long, description As String
    On Error GoTo Echec
    Set col = modCerfaPrint.AnalyserPositions("# TEST FICTIF" & vbLf & "ASSURE_NOM;20,5;30;100;10")
    If col.Count <> 1 Or CDbl(col(1)(1)) <> 20.5 Then Err.Raise 5, , "Positions CERFA valides"
    reussis = 1
    For Each cas In Array("ASSURE_NOM;abc;30;100;10", "ASSURE_NOM;20;30;-1;10", "ASSURE_NOM;200;30;100;10", "ASSURE_NOM;20;295;100;10", "ASSURE_NOM;20;30;100;0", "ASSURE_NOM;20;30;100", "ASSURE_NOM;20;30;100;10" & vbLf & "assure_nom;20;30;100;10", "# vide")
        numero = 0
        On Error Resume Next
        Set col = modCerfaPrint.AnalyserPositions(CStr(cas))
        numero = Err.Number: Err.Clear
        On Error GoTo Echec
        If numero <> vbObjectError + 710 Then Err.Raise 5, , "Position CERFA invalide acceptee"
        reussis = reussis + 1
    Next cas
    If modServiceNas.SHA256("abc") <> "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad" Then Err.Raise 5, , "SHA256 Excel"
    reussis = reussis + 1
    If modUI.MontantReglement("12,30") <> 12.3@ Or modUI.MontantReglement("12.3") <> 12.3@ Or modUI.MontantReglement("12") <> 12@ Then Err.Raise 5, , "Montants decimaux exacts"
    reussis = reussis + 1
    For Each cas In Array("12abc", "1e2", "-1", "12.345", "1,2.3", "", "10000000")
        Dim somme As Currency
        numero = 0
        On Error Resume Next
        somme = modUI.MontantReglement(CStr(cas))
        numero = Err.Number: Err.Clear
        On Error GoTo Echec
        If numero <> vbObjectError + 661 Then Err.Raise 5, , "Montant incorrect accepte"
        reussis = reussis + 1
    Next cas
    Dim sauvegardees As New Collection, ligneFigee As Object, ligneFigee2 As Object
    Dim identiteFigee As Object, lignesImpression As Collection
    Set ligneFigee = modServiceNas.Parametres()
    ligneFigee("PatientID") = "P1": ligneFigee("Date") = "12/09/2026"
    ligneFigee("Nom") = "ANCIEN": ligneFigee("Prenom") = "FICTIF": ligneFigee("DDN") = "01/01/1980"
    ligneFigee("NIR") = "280010000000169": ligneFigee("CodeActe") = "TEST": ligneFigee("CodeCerfa") = "TEST CERFA": ligneFigee("Montant") = "1.00"
    sauvegardees.Add ligneFigee
    Set identiteFigee = modActes.DonneesImpressionFigees(sauvegardees, "P1", lignesImpression)
    If CStr(identiteFigee("Nom")) <> "ANCIEN" Or CStr(identiteFigee("NIR")) <> "280010000000169" Or lignesImpression.Count <> 1 Or CStr(lignesImpression(1)("CodeActe")) <> "TEST CERFA" Then Err.Raise 5, , "Instantane fige non conserve"
    ligneFigee("NIR") = "": ligneFigee("AssureNom") = "ASSURE": ligneFigee("AssurePrenom") = "FICTIF"
    ligneFigee("AssureDDN") = "02/02/1970": ligneFigee("AssureNIR") = "170020000000157"
    Set identiteFigee = modActes.DonneesImpressionFigees(sauvegardees, "P1", lignesImpression)
    If CStr(identiteFigee("NIR")) <> "" Or CStr(identiteFigee("AssureNIR")) <> "170020000000157" Then Err.Raise 5, , "Assure distinct fige refuse"
    reussis = reussis + 1
    numero = 0
    On Error Resume Next
    Set identiteFigee = modActes.DonneesImpressionFigees(sauvegardees, "P2", lignesImpression)
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    If numero <> vbObjectError + 646 Then Err.Raise 5, , "Patient different accepte pour la reimpression"
    reussis = reussis + 1
    Set ligneFigee2 = modServiceNas.Parametres()
    For Each cas In ligneFigee.Keys: ligneFigee2(CStr(cas)) = ligneFigee(CStr(cas)): Next cas
    ligneFigee2("Nom") = "NOUVEAU": sauvegardees.Add ligneFigee2
    numero = 0
    On Error Resume Next
    Set identiteFigee = modActes.DonneesImpressionFigees(sauvegardees, "P1", lignesImpression)
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    If numero <> vbObjectError + 649 Then Err.Raise 5, , "Identites figees incoherentes acceptees"
    reussis = reussis + 1

    Dim infos As Object, acte As Object, acte2 As Object, actes As New Collection, repartition As Object
    Dim lignesFacturation As Collection
    Set infos = modServiceNas.Parametres()
    infos("ConsultationID") = "S1": infos("PatientID") = "P1": infos("DateActe") = "12/09/2026"
    infos("Nom") = "FICTIF": infos("Prenom") = "TEST": infos("DDN") = "01/01/1980": infos("NIR") = "280010000000169"
    Set acte = modServiceNas.Parametres()
    acte("Code") = "CS": acte("Tarif") = "25.00": acte("CodeAssocie") = "ECG": acte("TarifAssocie") = "14.26": acte("LibelleCerfa") = "CS CERFA"
    actes.Add acte
    Set lignesFacturation = modActes.ConstruireLignesFacturation(infos, actes, "CB", False, False)
    If lignesFacturation.Count <> 2 Or CStr(lignesFacturation(1)("TiersPayant")) <> "N" Or CStr(lignesFacturation(2)("TiersPayant")) <> "N" Or CStr(lignesFacturation(1)("Paye")) <> "O" Or CStr(lignesFacturation(2)("Paye")) <> "O" Then Err.Raise 5, , "Compatibilite paiement Patient"
    If CStr(lignesFacturation(1)("CodeCerfa")) <> "CS CERFA" Or CStr(lignesFacturation(2)("CodeCerfa")) <> "ECG" Then Err.Raise 5, , "Libelles CERFA non figes"
    reussis = reussis + 1
    Set lignesFacturation = modActes.ConstruireLignesFacturation(infos, actes, "CB", True, False)
    If CStr(lignesFacturation(1)("TiersPayant")) <> "O" Or CStr(lignesFacturation(2)("TiersPayant")) <> "O" Or CStr(lignesFacturation(1)("Paye")) <> "N" Or CStr(lignesFacturation(2)("Paye")) <> "N" Then Err.Raise 5, , "Compatibilite tiers payant complet"
    reussis = reussis + 1
    Set acte2 = modServiceNas.Parametres()
    acte2("Code") = "TIERS": acte2("Tarif") = "7.20": acte2("CodeAssocie") = "TIERSA": acte2("TarifAssocie") = "0.80"
    actes.Add acte2
    Set repartition = modServiceNas.Parametres(): repartition("CS") = False: repartition("TIERS") = True
    Set lignesFacturation = modActes.ConstruireLignesFacturation(infos, actes, "CB", repartition, False)
    If lignesFacturation.Count <> 4 Then Err.Raise 5, , "Nombre de lignes mixtes incorrect"
    If CStr(lignesFacturation(1)("TiersPayant")) <> "N" Or CStr(lignesFacturation(2)("TiersPayant")) <> "N" Or CStr(lignesFacturation(3)("TiersPayant")) <> "O" Or CStr(lignesFacturation(4)("TiersPayant")) <> "O" Then Err.Raise 5, , "Repartition mixte incorrecte"
    If CStr(lignesFacturation(1)("Paye")) <> "O" Or CStr(lignesFacturation(2)("Paye")) <> "O" Or CStr(lignesFacturation(3)("Paye")) <> "N" Or CStr(lignesFacturation(4)("Paye")) <> "N" Then Err.Raise 5, , "Statut de paiement mixte incorrect"
    reussis = reussis + 1
    repartition.Remove "TIERS": numero = 0
    On Error Resume Next
    Set lignesFacturation = modActes.ConstruireLignesFacturation(infos, actes, "CB", repartition, False)
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    If numero <> vbObjectError + 650 Then Err.Raise 5, , "Payeur manquant accepte"
    reussis = reussis + 1
    If reussis <> NOMBRE_ATTENDU_EXCEL Then Err.Raise 5, , "Nombre de controles Excel inattendu"
    Executer = "{""reussis"":" & CStr(reussis) & ",""attendus"":" & CStr(NOMBRE_ATTENDU_EXCEL) & ",""echec"":false}"
    Exit Function
Echec:
    description = Err.Description
    Executer = "{""reussis"":" & CStr(reussis) & ",""attendus"":" & CStr(NOMBRE_ATTENDU_EXCEL) & ",""echec"":true,""description"":" & modServiceNas.JsonValeur(description) & "}"
End Function
