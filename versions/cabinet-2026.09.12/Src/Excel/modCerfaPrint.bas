Attribute VB_Name = "modCerfaPrint"
Option Explicit
' =====================================================================
' modCerfaPrint - Impression calee sur la feuille de soins pre-imprimee
' (cerfa S3110). Le document Word est construit A LA VOLEE : une zone de
' texte flottante par champ, positions en mm dans Config\cerfa_positions.txt,
' decalage global d'imprimante dans Config\cerfa_offsets.txt (dx;dy en mm).
' =====================================================================

Private Const MM_EN_POINTS As Double = 2.834645

' infos : dictionnaire (Nom, Prenom, DDN, NIR)
' actes : Collection de dictionnaires (CodeActe/Code + Montant/Tarif)
' versPdf : si renseigne, exporte en PDF au lieu d'imprimer (tests/controle)
Public Sub ImprimerFeuille(ByVal infos As Object, ByVal actes As Collection, _
                           Optional ByVal versPdf As String = "", Optional ByVal verifierSeulement As Boolean = False)
    Dim valeurs As Object, a As Object, i As Long, total As Double
    Dim code As String, montant As String

    If actes.Count > 4 Then Err.Raise vbObjectError + 701, "modCerfaPrint", "Plus de quatre lignes : une seconde feuille de soins est necessaire. Aucune impression lancee."
    If Not modTexte.DateFrValide(ValeurOuVide(infos, "DateActe")) Then Err.Raise vbObjectError + 702, "modCerfaPrint", "Date de l acte absente ou invalide."
    Set valeurs = CreateObject("Scripting.Dictionary")
    valeurs.CompareMode = 1
    Dim assureDistinct As Boolean, nir As String, rpps As String, am As String
    ' Une DDN assuree isolee est une fiche incomplete, jamais le patient assure.
    assureDistinct = Len(Trim$(ValeurOuVide(infos, "AssureNom") & ValeurOuVide(infos, "AssurePrenom") & ValeurOuVide(infos, "AssureDDN") & ValeurOuVide(infos, "AssureNIR"))) > 0
    valeurs("PATIENT_NOM") = Trim$(ValeurOuVide(infos, "Nom") & " " & ValeurOuVide(infos, "Prenom"))
    valeurs("PATIENT_DDN") = ValeurOuVide(infos, "DDN")
    If assureDistinct Then
        If Len(Trim$(ValeurOuVide(infos, "AssureNom"))) = 0 Or Len(Trim$(ValeurOuVide(infos, "AssurePrenom"))) = 0 Or Not modTexte.DateFrValide(ValeurOuVide(infos, "AssureDDN")) Or Len(Trim$(ValeurOuVide(infos, "AssureNIR"))) = 0 Then Err.Raise vbObjectError + 703, , "Fiche de l assure incomplete ou invalide."
        valeurs("ASSURE_NOM") = Trim$(ValeurOuVide(infos, "AssureNom") & " " & ValeurOuVide(infos, "AssurePrenom"))
        valeurs("ASSURE_DDN") = ValeurOuVide(infos, "AssureDDN")
        nir = ValeurOuVide(infos, "AssureNIR")
        ExigerPositions Array("PATIENT_NOM", "PATIENT_DDN")
    Else
        valeurs("ASSURE_NOM") = valeurs("PATIENT_NOM")
        valeurs("ASSURE_DDN") = valeurs("PATIENT_DDN")
        nir = ValeurOuVide(infos, "NIR")
    End If
    Dim p As Object, r As Object
    Set p = modServiceNas.Parametres(): p("nir") = nir
    Set r = modServiceNas.Appeler("nir.validate", p)
    valeurs("ASSURE_NIR") = CStr(r("nir"))
    rpps = modConfig.Config("MEDECIN", "RPPS", "")
    am = modConfig.Config("MEDECIN", "NumeroAM", "")
    If Not ChiffresExactement(rpps, 11) Or Not ChiffresExactement(am, 9) Then Err.Raise vbObjectError + 704, , "Renseignez RPPS (11 chiffres) et numero AM (9 chiffres) dans la configuration."
    valeurs("MEDECIN_RPPS") = rpps: valeurs("MEDECIN_AM") = am
    If modConfig.Config("CERFA", "PraticienPreimprime", "0") <> "1" Then ExigerPositions Array("MEDECIN_RPPS", "MEDECIN_AM")
    If modConfig.Config("CERFA", "CalageValide", "0") <> "1" And Len(versPdf) = 0 Then Err.Raise vbObjectError + 705, , "Validez le calage papier avant impression (CERFA CalageValide=1 dans poste.ini)."
    i = 0
    For Each a In actes
        i = i + 1
        If i > 4 Then Exit For
        If a.Exists("CodeActe") Then code = a("CodeActe") Else code = ValeurOuVide(a, "Code")
        If a.Exists("Montant") Then montant = a("Montant") Else montant = ValeurOuVide(a, "Tarif")
        valeurs("DATE" & i) = Format$(modTexte.DateFr(ValeurOuVide(infos, "DateActe")), "dd/mm/yyyy")
        valeurs("CODE" & i) = code
        valeurs("MONTANT" & i) = Format$(Val(Replace(montant, ",", ".")), "0.00")
        total = total + Val(Replace(montant, ",", "."))
    Next a
    valeurs("TOTAL") = Format$(total, "0.00")

    Dim champ As Variant
    For Each champ In valeurs.Keys
        If Len(CStr(valeurs(champ))) > 0 Then
            If CStr(champ) <> "PATIENT_NOM" And CStr(champ) <> "PATIENT_DDN" Then
                If Not ((CStr(champ) = "MEDECIN_RPPS" Or CStr(champ) = "MEDECIN_AM") And modConfig.Config("CERFA", "PraticienPreimprime", "0") = "1") Then ExigerPositions Array(CStr(champ))
            End If
        End If
    Next champ
    Dim dxControle As Double, dyControle As Double, positionControle As Variant
    LireOffsets dxControle, dyControle
    For Each positionControle In LirePositions()
        If valeurs.Exists(CStr(positionControle(0))) Then
            If CDbl(positionControle(1)) + dxControle < 0 Or CDbl(positionControle(2)) + dyControle < 0 Or CDbl(positionControle(1)) + dxControle + CDbl(positionControle(3)) > 210 Or CDbl(positionControle(2)) + dyControle + 7 > 297 Then Err.Raise vbObjectError + 710, , "Calage CERFA hors de la page."
        End If
    Next positionControle
    If verifierSeulement Then
        Dim imprimante As String, service As Object, item As Object, trouve As Boolean
        imprimante = modConfig.Config("CERFA", "Imprimante", "")
        If Len(Trim$(imprimante)) = 0 Then Err.Raise vbObjectError + 708, , "Choisissez une imprimante CERFA dans poste.ini."
        Set service = GetObject("winmgmts:\\.\root\cimv2")
        For Each item In service.ExecQuery("SELECT Name FROM Win32_Printer")
            If StrComp(CStr(item.Name), imprimante, vbTextCompare) = 0 Then trouve = True
        Next item
        If Not trouve Then Err.Raise vbObjectError + 708, , "Imprimante CERFA configuree introuvable."
        Exit Sub
    End If
    ImprimerDocumentCale valeurs, versPdf
End Sub

' Construit et imprime (ou exporte) le document cale
Private Sub ImprimerDocumentCale(ByVal valeurs As Object, ByVal versPdf As String)
    Dim word As Object, doc As Object, shp As Object
    Dim positions As Collection, p As Variant, dx As Double, dy As Double

    Set positions = LirePositions()
    LireOffsets dx, dy

    Set word = CreateObject("Word.Application")
    word.visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3
    On Error GoTo Nettoyage
    Set doc = word.Documents.Add()
    doc.PageSetup.PageWidth = 210 * MM_EN_POINTS
    doc.PageSetup.PageHeight = 297 * MM_EN_POINTS
    doc.PageSetup.TopMargin = 0
    doc.PageSetup.BottomMargin = 0
    doc.PageSetup.LeftMargin = 0
    doc.PageSetup.RightMargin = 0

    For Each p In positions
        If valeurs.Exists(CStr(p(0))) Then
            If Len(CStr(valeurs(p(0)))) > 0 Then
                Set shp = doc.Shapes.AddTextbox(1, _
                    (CDbl(p(1)) + dx) * MM_EN_POINTS, (CDbl(p(2)) + dy) * MM_EN_POINTS, _
                    CDbl(p(3)) * MM_EN_POINTS, 18)
                shp.RelativeHorizontalPosition = 1
                shp.RelativeVerticalPosition = 1
                shp.TextFrame.TextRange.Text = CStr(valeurs(p(0)))
                shp.TextFrame.TextRange.Font.Name = "Arial"
                shp.TextFrame.TextRange.Font.Size = CDbl(p(4))
                shp.TextFrame.MarginLeft = 0
                shp.TextFrame.MarginTop = 0
                shp.Line.visible = 0        ' msoFalse
                shp.Fill.visible = 0
            End If
        End If
    Next p

    If Len(versPdf) > 0 Then
        doc.ExportAsFixedFormat versPdf, 17
    Else
        Dim imprimante As String
        imprimante = modConfig.Config("CERFA", "Imprimante", "")
        If Len(imprimante) > 0 Then word.ActivePrinter = imprimante
        doc.PrintOut Background:=False
    End If
    doc.Close 0
    word.Quit
    Exit Sub
Nettoyage:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    word.Quit
    On Error GoTo 0
    Err.Raise numErr, "modCerfaPrint", descErr
End Sub

' Impression de croix de reperes puis saisie des ecarts mesures
Public Sub CalageCerfa()
    Dim valeurs As Object, reponse As String, dx As Double, dy As Double
    If MsgBox("Placez une feuille de soins SACRIFIEE dans l'imprimante." & vbCrLf & _
              "Des croix de repere vont s'imprimer aux 4 coins (a 20 mm des bords)." & vbCrLf & vbCrLf & _
              "Imprimer maintenant ?", vbOKCancel + vbInformation, "Calage CERFA") <> vbOK Then Exit Sub

    Dim word As Object, doc As Object, shp As Object, coords As Variant, c As Variant
    Set word = CreateObject("Word.Application")
    word.visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3
    On Error GoTo Nettoyage
    Set doc = word.Documents.Add()
    doc.PageSetup.PageWidth = 210 * MM_EN_POINTS
    doc.PageSetup.PageHeight = 297 * MM_EN_POINTS
    doc.PageSetup.TopMargin = 0: doc.PageSetup.BottomMargin = 0
    doc.PageSetup.LeftMargin = 0: doc.PageSetup.RightMargin = 0
    coords = Array(Array(20, 20), Array(190, 20), Array(20, 277), Array(190, 277))
    For Each c In coords
        Set shp = doc.Shapes.AddTextbox(1, (CDbl(c(0)) - 2) * MM_EN_POINTS, _
                                        (CDbl(c(1)) - 4) * MM_EN_POINTS, 12 * MM_EN_POINTS, 16)
        shp.RelativeHorizontalPosition = 1
        shp.RelativeVerticalPosition = 1
        shp.TextFrame.TextRange.Text = "+"
        shp.TextFrame.TextRange.Font.Size = 14
        shp.TextFrame.MarginLeft = 0: shp.TextFrame.MarginTop = 0
        shp.Line.visible = 0: shp.Fill.visible = 0
    Next c
    Dim imprimante As String
    imprimante = modConfig.Config("CERFA", "Imprimante", "")
    If Len(imprimante) > 0 Then word.ActivePrinter = imprimante
    doc.PrintOut Background:=False
    doc.Close 0
    word.Quit
    On Error GoTo 0

    reponse = InputBox("Mesurez sur la feuille imprimee :" & vbCrLf & _
        "- ecart HORIZONTAL en mm entre la croix haut-gauche et 20 mm du bord gauche" & vbCrLf & _
        "  (positif si la croix est trop a droite)", "Calage CERFA - decalage X", "0")
    If Len(reponse) = 0 Then Exit Sub
    dx = -Val(Replace(reponse, ",", "."))
    reponse = InputBox("- ecart VERTICAL en mm entre la croix haut-gauche et 20 mm du bord haut" & vbCrLf & _
        "  (positif si la croix est trop basse)", "Calage CERFA - decalage Y", "0")
    If Len(reponse) = 0 Then Exit Sub
    dy = -Val(Replace(reponse, ",", "."))
    modFichiers.EcrireTexteUTF8 Environ$("APPDATA") & "\CabinetCardio\cerfa_offsets.txt", _
        Replace(CStr(dx), ",", ".") & ";" & Replace(CStr(dy), ",", ".")
    MsgBox "Calage enregistre (dx=" & dx & " mm, dy=" & dy & " mm)." & vbCrLf & _
           "Refaites une impression d'essai pour verifier.", vbInformation, "Calage CERFA"
    Exit Sub
Nettoyage:
    Dim descErr As String
    descErr = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    word.Quit
    On Error GoTo 0
    MsgBox "Erreur d'impression : " & descErr, vbCritical, "Calage CERFA"
End Sub

Private Function LirePositions() As Collection
    Dim chemin As String
    chemin = modConfig.chemin("Config") & "\cerfa_positions.txt"
    If Not modFichiers.FichierExiste(chemin) Then Err.Raise vbObjectError + 700, , "Fichier de positions CERFA introuvable."
    Set LirePositions = AnalyserPositions(modFichiers.LireTexteUTF8(chemin))
End Function

Public Function AnalyserPositions(ByVal contenu As String) As Collection
    Dim lignes As Variant, parties As Variant, col As New Collection, vus As Object
    Dim i As Long, cle As String, x As Double, y As Double, largeur As Double, police As Double, re As Object
    Set vus = CreateObject("Scripting.Dictionary"): vus.CompareMode = 1
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[A-Z][A-Z0-9_]*$"
    lignes = Split(Replace(contenu, vbCr, ""), vbLf)
    For i = LBound(lignes) To UBound(lignes)
        If Len(Trim$(lignes(i))) > 0 And Left$(Trim$(lignes(i)), 1) <> "#" Then
            parties = Split(lignes(i), ";")
            If UBound(parties) <> 4 Then Err.Raise vbObjectError + 710, , "Position CERFA : cinq champs requis."
            cle = UCase$(Trim$(parties(0)))
            If Not re.Test(cle) Or vus.Exists(cle) Then Err.Raise vbObjectError + 710, , "Position CERFA vide, invalide ou dupliquee."
            x = NombrePosition(CStr(parties(1))): y = NombrePosition(CStr(parties(2)))
            largeur = NombrePosition(CStr(parties(3))): police = NombrePosition(CStr(parties(4)))
            If x < 0 Or y < 0 Or largeur <= 0 Or x + largeur > 210 Or y + 7 > 297 Or police < 4 Or police > 72 Then Err.Raise vbObjectError + 710, , "Position CERFA hors page ou dimension invalide."
            col.Add Array(cle, x, y, largeur, police): vus.Add cle, True
        End If
    Next i
    If col.Count = 0 Then Err.Raise vbObjectError + 710, , "Aucune position CERFA configuree."
    Set AnalyserPositions = col
End Function

Private Function NombrePosition(ByVal texte As String) As Double
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^-?[0-9]+([.,][0-9]+)?$"
    texte = Trim$(texte)
    If Not re.Test(texte) Then Err.Raise vbObjectError + 710, , "Coordonnee CERFA non numerique."
    NombrePosition = Val(Replace(texte, ",", "."))
End Function

Private Sub LireOffsets(ByRef dx As Double, ByRef dy As Double)
    Dim chemin As String, contenu As String, parties() As String
    dx = 0: dy = 0
    chemin = Environ$("APPDATA") & "\CabinetCardio\cerfa_offsets.txt"
    If Not modFichiers.FichierExiste(chemin) Then Exit Sub
    contenu = Trim$(Replace(Replace(modFichiers.LireTexteUTF8(chemin), vbCr, ""), vbLf, ""))
    parties = Split(contenu, ";")
    If UBound(parties) <> 1 Then Err.Raise vbObjectError + 710, , "Deux decalages CERFA sont requis."
    dx = NombrePosition(CStr(parties(0)))
    dy = NombrePosition(CStr(parties(1)))
End Sub

Private Function ValeurOuVide(ByVal dict As Object, ByVal cle As String) As String
    If dict.Exists(cle) Then ValeurOuVide = dict(cle) Else ValeurOuVide = ""
End Function

Private Function ChiffresExactement(ByVal valeur As String, ByVal longueur As Long) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9]{" & CStr(longueur) & "}$"
    ChiffresExactement = re.Test(valeur)
End Function

Private Sub ExigerPositions(ByVal champs As Variant)
    Dim presents As Object, position As Variant, cle As Variant
    Set presents = CreateObject("Scripting.Dictionary")
    For Each position In LirePositions()
        presents(CStr(position(0))) = True
    Next position
    For Each cle In champs
        If Not presents.Exists(CStr(cle)) Then Err.Raise vbObjectError + 706, , "Position CERFA non configuree : " & CStr(cle) & ". Calage a effectuer sur votre formulaire."
    Next cle
End Sub

Public Sub VerifierAvantFacturation(ByVal infos As Object, ByVal actes As Collection)
    If actes.Count = 0 Then Err.Raise vbObjectError + 709, , "Aucune ligne pour la feuille de soins."
    ' La publication ne transporte plus la fiche administrative complete.
    ' Verifier les donnees courantes avant la premiere facturation ; les
    ' reimpressions utilisent ensuite l instantane fige de la seance.
    Dim pat As Object, copie As Object, cle As Variant
    Set pat = modServiceNas.LireID("PATIENTS", CStr(infos("PatientID")))
    Set copie = CreateObject("Scripting.Dictionary")
    For Each cle In infos.Keys: copie(CStr(cle)) = infos(cle): Next cle
    For Each cle In Array("Nom", "Prenom", "DDN")
        If CStr(pat(cle)) <> CStr(infos(cle)) Then Err.Raise vbObjectError + 711, , "Identite modifiee depuis la publication : faire verifier le courrier."
    Next cle
    For Each cle In Array("NIR", "AssureNom", "AssurePrenom", "AssureDDN", "AssureNIR")
        copie(CStr(cle)) = CStr(pat(cle))
    Next cle
    ImprimerFeuille copie, actes, "", True
End Sub

Public Sub VerifierAvantReimpression(ByVal infos As Object, ByVal actes As Collection)
    Dim cle As Variant
    If actes.Count = 0 Then Err.Raise vbObjectError + 709, , "Aucune ligne pour la feuille de soins."
    For Each cle In Array("PatientID", "Nom", "Prenom", "DDN", "NIR")
        If Not infos.Exists(CStr(cle)) Then Err.Raise vbObjectError + 712, , "Identite figee incomplete : " & CStr(cle)
    Next cle
    If Len(Trim$(CStr(infos("PatientID")))) = 0 Then Err.Raise vbObjectError + 712, , "Patient fige absent."
    If Len(Trim$(CStr(infos("Nom")))) = 0 Then Err.Raise vbObjectError + 712, , "Nom fige absent."
    If Len(Trim$(CStr(infos("Prenom")))) = 0 Then Err.Raise vbObjectError + 712, , "Prenom fige absent."
    If Not modTexte.DateFrValide(CStr(infos("DDN"))) Then Err.Raise vbObjectError + 712, , "Naissance figee invalide."
    ImprimerFeuille infos, actes, "", True
End Sub
