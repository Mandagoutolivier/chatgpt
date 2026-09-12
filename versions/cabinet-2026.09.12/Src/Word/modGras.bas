Attribute VB_Name = "modGras"
Option Explicit
' =====================================================================
' modGras - Mise en gras automatique du corps des courriers :
'  - medicaments (Config\Gras_Medicaments.xlsx, feuille MEDICAMENTS :
'    Terme | Variantes (separees par |) | Actif) -> GRAS + MAJUSCULES
'  - expressions (Config\Gras_Expressions.xlsx, feuille EXPRESSIONS :
'    Terme | Actif) -> GRAS, casse du texte conservee
' Recherche mot entier, insensible a la casse, limitee au signet CORPS.
' Appelee automatiquement apres la correction IA et la lettre derivee,
' et a la demande (Ctrl+Alt+B / ruban Cabinet > Mettre en gras).
' Les deux classeurs sont modifiables par la secretaire ou le medecin
' (ruban Cabinet > Medicaments... / Expressions...).
' =====================================================================

Private mMeds As Collection
Private mExps As Collection
Private mCharge As Date
Private Const CACHE_S As Long = 300

Public Function FichierMedicaments() As String
    FichierMedicaments = modConfig.chemin("Config") & "\Gras_Medicaments.xlsx"
End Function

Public Function FichierExpressions() As String
    FichierExpressions = modConfig.chemin("Config") & "\Gras_Expressions.xlsx"
End Function

Public Sub InvaliderCache()
    mCharge = 0
End Sub

Private Sub Charger(ByVal forcer As Boolean)
    If Not forcer And Not mMeds Is Nothing Then
        If DateDiff("s", mCharge, Now) < CACHE_S Then Exit Sub
    End If
    Set mMeds = LireListe(FichierMedicaments(), "MEDICAMENTS")
    Set mExps = LireListe(FichierExpressions(), "EXPRESSIONS")
    mCharge = Now
    modLog.LogInfo "Gras : " & mMeds.Count & " medicaments, " & mExps.Count & " expressions charges"
End Sub

' Lit une feuille (Terme, Variantes, Actif) : renvoie la liste des termes actifs
Private Function LireListe(ByVal fichier As String, ByVal feuille As String) As Collection
    Dim res As New Collection, lignes As Collection, l As Object, parts As Variant, p As Variant
    Set LireListe = res
    Set lignes = modServiceNas.LireTable(feuille)
    For Each l In lignes
        If EstActif(l) Then
            res.Add Trim$(l("Terme"))
            If l.Exists("Variantes") Then
                parts = Split(l("Variantes"), "|")
                For Each p In parts
                    If Len(Trim$(p)) > 0 Then res.Add Trim$(p)
                Next p
            End If
        End If
    Next l
End Function

Private Function EstActif(ByVal l As Object) As Boolean
    EstActif = True
    If l.Exists("Actif") Then
        Select Case LCase$(Trim$(l("Actif")))
            Case "0", "non", "n", "false": EstActif = False
        End Select
    End If
End Function

Public Function NbMedicaments() As Long
    Charger False
    NbMedicaments = mMeds.Count
End Function

Public Function NbExpressions() As Long
    Charger False
    NbExpressions = mExps.Count
End Function

' Applique le gras sur le corps du document ; renvoie le nombre d'occurrences.
Public Function AppliquerGras(ByVal doc As Document) As Long
    Dim t As Variant, n As Long
    If Not doc.Bookmarks.Exists("CORPS") Then Exit Function
    Charger False
    For Each t In mExps
        n = n + Marquer(doc, CStr(t), False)
    Next t
    For Each t In mMeds
        n = n + Marquer(doc, CStr(t), True)
    Next t
    AppliquerGras = n
End Function

' Marque toutes les occurrences (mot entier, sans casse) d'un terme dans CORPS.
Private Function Marquer(ByVal doc As Document, ByVal terme As String, ByVal majuscules As Boolean) As Long
    Dim variantes As Variant, v As Variant
    ' apostrophe droite / typographique : les deux formes sont cherchees
    If InStr(terme, "'") > 0 Then
        variantes = Array(terme, Replace(terme, "'", ChrW(8217)))
    Else
        variantes = Array(terme)
    End If
    For Each v In variantes
        Marquer = Marquer + MarquerForme(doc, CStr(v), majuscules)
    Next v
End Function

Private Function MarquerForme(ByVal doc As Document, ByVal terme As String, ByVal majuscules As Boolean, Optional ByVal zone As Range = Nothing) As Long
    Dim rng As Range, fin As Long, n As Long
    If Len(Trim$(terme)) = 0 Then Exit Function
    If zone Is Nothing Then Set zone = doc.Bookmarks("CORPS").Range
    fin = zone.End
    Set rng = zone.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = terme
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With
    Do While rng.Find.Execute
        If rng.End > fin Then Exit Do
        rng.Font.Bold = True
        If majuscules Then rng.Case = wdUpperCase
        n = n + 1
        rng.Collapse wdCollapseEnd
        rng.End = fin
        If rng.Start >= fin Then Exit Do
    Loop
    MarquerForme = n
End Function

' ---- commandes utilisateur -------------------------------------------------

Public Sub MettreEnGras()
    On Error GoTo Erreur
    Dim n As Long, doc As Document
    Set doc = ActiveDocument
    If Not doc.Bookmarks.Exists("CORPS") Then
        MsgBox "Ce document n'est pas un courrier du cabinet (signet CORPS absent).", vbExclamation, "Cabinet"
        Exit Sub
    End If
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Mise en gras"
    n = AppliquerGras(doc)
    ur.EndCustomRecord
    Application.StatusBar = "Mise en gras : " & n & " occurrence(s) (" & NbMedicaments() & " medicaments, " & _
                            NbExpressions() & " expressions dans les dictionnaires)."
    Exit Sub
Erreur:
    modLog.LogErreur "MettreEnGras : " & Err.Description
    MsgBox "Mise en gras impossible : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub OuvrirDictionnaireMedicaments()
    OuvrirDictionnaire FichierMedicaments()
End Sub

Public Sub OuvrirDictionnaireExpressions()
    OuvrirDictionnaire FichierExpressions()
End Sub

' Ouvre le classeur dans Excel (visible) pour modification ; le cache sera
' recharge au prochain passage en gras.
Private Sub OuvrirDictionnaire(ByVal fichier As String)
    Dim texte As String, p As Object, r As Object
    texte = Trim$(InputBox("Terme a ajouter au dictionnaire commun :", "Dictionnaire du cabinet"))
    If Len(texte) = 0 Then Exit Sub
    Set p = modServiceNas.Parametres()
    If InStr(1, fichier, "Medicaments", vbTextCompare) > 0 Then p("genre") = "MEDICAMENTS" Else p("genre") = "EXPRESSIONS"
    p("texte") = texte
    Set r = modServiceNas.Appeler("dictionary.add", p)
    InvaliderCache
    MsgBox "Terme ajoute sur le NAS.", vbInformation, "Cabinet"
End Sub
Public Function AppliquerGrasDocumentComplet(ByVal doc As Document) As Long
    Dim p As Paragraph, zones As New Collection, zone As Range, debut As Long, ouvert As Boolean, t As Variant
    Charger False
    For Each p In doc.Paragraphs
        If EstFormuleAppel(NettoyerTexteParagraphe(p.Range.Text)) And Not ouvert Then
            debut = p.Range.End
            ouvert = True
        ElseIf ouvert And EstSignature(NettoyerTexteParagraphe(p.Range.Text)) Then
            If p.Range.Start > debut Then zones.Add doc.Range(debut, p.Range.Start)
            ouvert = False
        End If
    Next p
    If zones.Count = 0 Then Err.Raise vbObjectError + 541, "modGras", "Corps des lettres introuvable pour la mise en gras."
    For Each zone In zones
        For Each t In mExps
            AppliquerGrasDocumentComplet = AppliquerGrasDocumentComplet + MarquerForme(doc, CStr(t), False, zone)
            If InStr(CStr(t), "'") > 0 Then Call MarquerForme(doc, Replace(CStr(t), "'", ChrW(8217)), False, zone)
        Next t
        For Each t In mMeds
            AppliquerGrasDocumentComplet = AppliquerGrasDocumentComplet + MarquerForme(doc, CStr(t), True, zone)
        Next t
    Next zone
End Function
