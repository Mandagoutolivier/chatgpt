Attribute VB_Name = "modAnonymise"
Option Explicit
' =====================================================================
' modAnonymise - Point de surete RGPD n.1.
' Anonymisation DIRIGEE : le courrier est rattache a un patient et un
' correspondant CONNUS de la base ; leurs identites sont remplacees par
' des balises {{...}} avant tout envoi a l'API, puis reinjectees.
'  - recherche insensible a la casse ET aux accents (pliage 1:1)
'  - variantes multiples de la date de naissance
'  - scan residuel BLOQUANT avant envoi (identites, NIR, telephones)
'  - verification des balises au retour avant reinjection
' =====================================================================

' Pliage 1 caractere -> 1 caractere (les positions restent alignees)
Private Function Plier1(ByVal s As String) As String
    Dim src As String, dst As String, i As Long, p As Long, c As String, res As String
    s = LCase$(s)
    src = ChrW$(224) & ChrW$(226) & ChrW$(228) & ChrW$(225) & ChrW$(227) & ChrW$(229) & _
          ChrW$(233) & ChrW$(232) & ChrW$(234) & ChrW$(235) & _
          ChrW$(237) & ChrW$(236) & ChrW$(238) & ChrW$(239) & _
          ChrW$(243) & ChrW$(242) & ChrW$(244) & ChrW$(246) & ChrW$(245) & _
          ChrW$(250) & ChrW$(249) & ChrW$(251) & ChrW$(252) & _
          ChrW$(253) & ChrW$(255) & ChrW$(231) & ChrW$(241) & ChrW$(230) & ChrW$(339)
    dst = "aaaaaaeeeeiiiiooooouuuuyycnao"
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        p = InStr(src, c)
        If p > 0 Then res = res & Mid$(dst, p, 1) Else res = res & c
    Next i
    Plier1 = res
End Function

Private Function EstLettreOuChiffre(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    Select Case AscW(Plier1(c))
        Case 48 To 57, 97 To 122
            EstLettreOuChiffre = True
        Case Else
            EstLettreOuChiffre = (c = "-")
    End Select
End Function

' ---------------------------------------------------------------------
' Contexte d'anonymisation : dictionnaire "retour" (balise -> valeur
' canonique) + liste ordonnee "vers" (valeur -> balise, du plus long au
' plus court).
' ---------------------------------------------------------------------
Public Function Construire(ByVal pat As Object, ByVal cor As Object) As Object
    Dim ctx As Object, vers As Object, retour As Object
    Set ctx = CreateObject("Scripting.Dictionary")
    Set vers = CreateObject("Scripting.Dictionary")
    vers.CompareMode = 0
    Set retour = CreateObject("Scripting.Dictionary")

    If Not pat Is Nothing Then
        AjouterValeur vers, retour, pat("Nom"), "{{PAT_NOM}}"
        AjouterValeur vers, retour, pat("NomNaissance"), "{{PAT_NOM_NAISSANCE}}"
        AjouterValeur vers, retour, pat("Prenom"), "{{PAT_PRENOM}}"
        AjouterDDN vers, retour, pat("DDN"), "{{PAT_DDN}}"
        AjouterValeur vers, retour, pat("NIR"), "{{PAT_NIR}}"
        AjouterValeur vers, retour, Replace(pat("NIR"), " ", ""), "{{PAT_NIR}}"
        AjouterValeur vers, retour, pat("Adresse1"), "{{PAT_ADRESSE}}"
        AjouterValeur vers, retour, pat("Adresse2"), "{{PAT_ADRESSE2}}"
        AjouterValeur vers, retour, pat("Tel"), "{{PAT_TEL}}"
        AjouterValeur vers, retour, pat("Mobile"), "{{PAT_MOBILE}}"
        AjouterValeur vers, retour, pat("Email"), "{{PAT_EMAIL}}"
    End If
    If Not cor Is Nothing Then
        AjouterValeur vers, retour, cor("Nom"), "{{DEST_NOM}}"
        AjouterValeur vers, retour, cor("Prenom"), "{{DEST_PRENOM}}"
    End If

    Set ctx("vers") = vers
    Set ctx("retour") = retour
    Set Construire = ctx
End Function

' Ajoute un correspondant supplementaire au contexte (ex : destinataire
' d'une lettre derivee), avec son propre prefixe de balise.
Public Sub AjouterCorrespondant(ByVal ctx As Object, ByVal cor As Object, ByVal prefixe As String)
    If cor Is Nothing Then Exit Sub
    AjouterValeur ctx("vers"), ctx("retour"), cor("Nom"), "{{" & prefixe & "_NOM}}"
    AjouterValeur ctx("vers"), ctx("retour"), cor("Prenom"), "{{" & prefixe & "_PRENOM}}"
End Sub

Private Sub AjouterValeur(ByVal vers As Object, ByVal retour As Object, _
                          ByVal valeur As String, ByVal balise As String)
    valeur = Trim$(valeur)
    If Len(valeur) < 2 Then Exit Sub          ' trop court = trop de faux positifs
    If Not vers.Exists(valeur) Then vers(valeur) = balise
    If Not retour.Exists(balise) Then retour(balise) = valeur
End Sub

' Variantes de la date de naissance jj/mm/aaaa
Private Sub AjouterDDN(ByVal vers As Object, ByVal retour As Object, _
                       ByVal ddn As String, ByVal balise As String)
    Dim parties() As String, j As Long, m As Long, a As Long
    Dim moisNoms As Variant, jTxt As String
    ddn = Trim$(ddn)
    If Len(ddn) = 0 Then Exit Sub
    AjouterValeur vers, retour, ddn, balise
    parties = Split(ddn, "/")
    If UBound(parties) <> 2 Then Exit Sub
    j = Val(parties(0)): m = Val(parties(1)): a = Val(parties(2))
    If j = 0 Or m = 0 Or m > 12 Or a = 0 Then Exit Sub
    moisNoms = Array("janvier", "février", "mars", "avril", "mai", "juin", "juillet", _
                     "août", "septembre", "octobre", "novembre", "décembre")
    If j = 1 Then jTxt = "1er" Else jTxt = CStr(j)
    AjouterValeur vers, retour, Format$(j, "00") & "/" & Format$(m, "00") & "/" & a, balise
    AjouterValeur vers, retour, j & "/" & m & "/" & a, balise
    AjouterValeur vers, retour, Format$(j, "00") & "-" & Format$(m, "00") & "-" & a, balise
    AjouterValeur vers, retour, Format$(j, "00") & "." & Format$(m, "00") & "." & a, balise
    AjouterValeur vers, retour, jTxt & " " & moisNoms(m - 1) & " " & a, balise
    AjouterValeur vers, retour, j & " " & moisNoms(m - 1) & " " & a, balise
End Sub

' ---------------------------------------------------------------------
' Anonymisation : remplace chaque valeur (du plus long au plus court),
' en mots entiers, sans tenir compte de la casse ni des accents.
' ---------------------------------------------------------------------
Public Function Anonymiser(ByVal texte As String, ByVal ctx As Object) As String
    Dim vers As Object, cles() As String, i As Long, n As Long, tmp As String
    Set vers = ctx("vers")
    n = vers.Count
    If n = 0 Then Anonymiser = texte: Exit Function
    ReDim cles(0 To n - 1)
    Dim k As Variant, idx As Long
    idx = 0
    For Each k In vers.Keys
        cles(idx) = k
        idx = idx + 1
    Next k
    ' tri par longueur decroissante (tri par insertion, n petit)
    Dim a As Long, b As Long
    For a = 1 To n - 1
        tmp = cles(a)
        b = a - 1
        Do While b >= 0
            If Len(cles(b)) >= Len(tmp) Then Exit Do
            cles(b + 1) = cles(b)
            b = b - 1
        Loop
        cles(b + 1) = tmp
    Next a
    For i = 0 To n - 1
        texte = RemplacerPlie(texte, cles(i), CStr(vers(cles(i))))
    Next i
    Anonymiser = texte
End Function

' Remplacement en mots entiers, insensible casse/accents, via texte plie
Private Function RemplacerPlie(ByVal texte As String, ByVal valeur As String, _
                               ByVal balise As String) As String
    Dim tPlie As String, vPlie As String, pos As Long, depart As Long
    Dim avant As String, apres As String
    vPlie = Plier1(valeur)
    If Len(vPlie) = 0 Then RemplacerPlie = texte: Exit Function
    tPlie = Plier1(texte)
    depart = 1
    Do
        pos = InStr(depart, tPlie, vPlie, vbBinaryCompare)
        If pos = 0 Then Exit Do
        If pos > 1 Then avant = Mid$(texte, pos - 1, 1) Else avant = ""
        If pos + Len(valeur) <= Len(texte) Then apres = Mid$(texte, pos + Len(valeur), 1) Else apres = ""
        If Not EstLettreOuChiffre(avant) And Not EstLettreOuChiffre(apres) Then
            texte = Left$(texte, pos - 1) & balise & Mid$(texte, pos + Len(valeur))
            tPlie = Left$(tPlie, pos - 1) & balise & Mid$(tPlie, pos + Len(valeur))
            depart = pos + Len(balise)
        Else
            depart = pos + 1
        End If
    Loop
    RemplacerPlie = texte
End Function

' ---------------------------------------------------------------------
' Scan residuel BLOQUANT : renvoie "" si le texte anonymise est sur,
' sinon la liste des problemes detectes.
' ---------------------------------------------------------------------
Public Function ScanResiduel(ByVal texteAnonyme As String, ByVal ctx As Object) As String
    Dim problemes As String, vers As Object, k As Variant
    Dim tPlie As String, re As Object

    ' retirer d'abord les balises posees (le texte d'une balise ne doit
    ' pas declencher de faux positif sur un nom court)
    Dim retour As Object
    Set retour = ctx("retour")
    For Each k In retour.Keys
        texteAnonyme = Replace(texteAnonyme, CStr(k), " ")
    Next k

    Set vers = ctx("vers")
    tPlie = Plier1(texteAnonyme)
    For Each k In vers.Keys
        If Len(k) >= 3 Then
            If InStr(tPlie, Plier1(CStr(k))) > 0 Then
                problemes = problemes & "- identite encore presente : " & Left$(k, 3) & "..." & vbCrLf
            End If
        End If
    Next k

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[12]\s?\d{2}\s?(0[1-9]|1[0-2])\s?\d{2}\s?\d{3}\s?\d{3}(\s?\d{2})?"
    If re.Test(texteAnonyme) Then
        problemes = problemes & "- numero de securite sociale (NIR) detecte" & vbCrLf
    End If

    re.Pattern = "0[1-9]([ .-]?\d{2}){4}"
    If re.Test(texteAnonyme) Then
        problemes = problemes & "- numero de telephone detecte" & vbCrLf
    End If

    ScanResiduel = problemes
End Function

' Verifie que la reponse de l'API ne contient que des balises connues ET
' qu'aucune balise presente dans le texte envoye n'a disparu (une balise
' perdue = une identite qui ne sera pas reinjectee, donc un courrier faux).
' Renvoie "" si tout est correct.
Public Function VerifierBalisesRetour(ByVal texteRetour As String, ByVal ctx As Object, _
                                      Optional ByVal texteEnvoye As String = "") As String
    Dim re As Object, matchs As Object, m As Object, retour As Object, pb As String
    Dim vues As Object
    Set retour = ctx("retour")
    Set vues = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\{\{[^}]{1,40}\}\}"
    Set matchs = re.Execute(texteRetour)
    For Each m In matchs
        vues(m.Value) = True
        If Not retour.Exists(m.Value) Then
            pb = pb & "- balise inconnue ou alteree : " & m.Value & vbCrLf
        End If
    Next m

    If Len(texteEnvoye) > 0 Then
        Set matchs = re.Execute(texteEnvoye)
        For Each m In matchs
            If retour.Exists(m.Value) And Not vues.Exists(m.Value) Then
                pb = pb & "- balise disparue de la reponse : " & m.Value & vbCrLf
                vues(m.Value) = True      ' ne signaler qu'une fois
            End If
        Next m
    End If

    VerifierBalisesRetour = pb
End Function

' Reinjection des identites dans le texte corrige
Public Function Reinjecter(ByVal texte As String, ByVal ctx As Object) As String
    Dim retour As Object, k As Variant
    Set retour = ctx("retour")
    For Each k In retour.Keys
        texte = Replace(texte, CStr(k), CStr(retour(k)))
    Next k
    Reinjecter = texte
End Function
