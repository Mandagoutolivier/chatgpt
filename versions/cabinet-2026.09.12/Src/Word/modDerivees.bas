Attribute VB_Name = "modDerivees"
Option Explicit
' =====================================================================
' modDerivees - Lettres derivees du courrier dicte : DEMANDE d'examen,
' d'avis specialise ou d'hospitalisation, adressee a un confrere.
'
' Une lettre derivee n'est PAS un resume du compte rendu : c'est une
' demande courte dont le canevas est un PROFIL (Config\profils\CODE.ini,
' voir modDemandes et LETTRES_DERIVEES.md). Le corps est redige par l'API
' a partir du corps ANONYMISE du courrier source et des consignes du
' profil ; l'en-tete, l'appel et la formule finale viennent du modele et
' de la fiche correspondant (modCourrier) : l'API ne redige QUE le corps.
'
' Deux points d'entree :
'  - LettreDerivee (Ctrl+Alt+D) : choix manuel du profil, de la modalite,
'    du motif et du destinataire ;
'  - modDemandes.GenererDemandesAutomatiques : a la validation du courrier
'    principal, une lettre par demande reperee, destinataire de priorite 1.
'
' Le destinataire est pre-rempli depuis le classeur des correspondants
' specialistes ([DERIVEES] FichierSpecialistes, feuilles Specialistes +
' Specialistes_ParType), filtre sur TYPE_BASE du profil et trie par
' priorite ; a defaut, la liste generale des correspondants.
' =====================================================================

' Commande principale (Ctrl+Alt+D / commande vocale)
Public Sub LettreDerivee()
    MsgBox "Les lettres complementaires sont preparees avec la touche D (Unifie_D_Finaliser).", vbInformation, "Cabinet"
End Sub
' Profils proposes : tous ceux de Config\profils, ou la liste
' config.ini [DERIVEES] Types=CODE;CODE;... si elle est renseignee

' Modalite du profil (libelle). "" si le profil n'en a pas, vbNullChar si annulation.

' ---------------------------------------------------------------------
' Destinataire
' ---------------------------------------------------------------------

' Destinataire sans intervention (generation automatique) : specialiste de
' priorite 1 pour TYPE_BASE ; a defaut une fiche "a completer" pour que la
' lettre existe quand meme et que le secretariat complete l'adresse.
Public Function DestinataireAutomatique(ByVal p As Object) As Object
    Dim items As Collection, d As Object
    Set items = CorrespondantsPourTypes(modDemandes.ValeurProfil(p, "IDENTITE", "TYPE_BASE"))
    If items.Count > 0 Then
        Set DestinataireAutomatique = items(1)
        Exit Function
    End If
    modLog.LogInfo "Aucun specialiste pour " & modDemandes.ValeurProfil(p, "IDENTITE", "TYPE_BASE") & " : destinataire a completer"
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("ID") = "A_COMPLETER"
    d("NomDestinataire") = "Destinataire à compléter"
    d("Titre") = "": d("Prenom") = "": d("Nom") = ""
    d("Specialite") = "": d("Adresse1") = "": d("Adresse2") = "": d("CP") = "": d("Ville") = ""
    d("Tel") = "": d("Email") = "": d("Structure") = ""
    d("BlocDestinataire") = "DESTINATAIRE À COMPLÉTER" & vbCr & "(" & modDemandes.LibelleProfil(p) & ")"
    d("Tutoiement") = "vous"
    d("FormuleAppel") = modCourrier.AppelParDefaut(False)
    d("FormulePolitesse") = modCourrier.PolitesseParDefaut(False)
    d("Priorite") = 999
    d("Actif") = "1"
    Set DestinataireAutomatique = d
End Function

' Specialistes actifs pour un ou plusieurs codes TypeExamen (virgules),
' tries par priorite croissante, sous la forme attendue par modCourrier
' (Titre, Prenom, Nom, Specialite, Adresse1/2, CP, Ville, FormuleAppel,
' FormulePolitesse, Tutoiement, BlocDestinataire, ID).
Public Function CorrespondantsPourTypes(ByVal typesExamen As String) As Collection
    Dim col As New Collection, cor As Object
    For Each cor In modBase.Correspondants()
        If InStr(1, CStr(cor("TypesExamen")), typesExamen, vbTextCompare) > 0 Then col.Add cor
    Next cor
    Set CorrespondantsPourTypes = col
End Function
Private Function FichierSpecialistes() As String
    Dim rel As String, chemin As String
    rel = modConfig.Config("DERIVEES", "FichierSpecialistes", "Base\Correspondants_Specialistes.xlsx")
    If Len(rel) = 0 Then Exit Function
    If InStr(rel, ":") > 0 Or Left$(rel, 2) = "\\" Then chemin = rel Else chemin = modConfig.racine() & "\" & rel
    If modFichiers.FichierExiste(chemin) Then
        FichierSpecialistes = chemin
    Else
        modLog.LogInfo "Classeur des specialistes absent : " & chemin & " (liste generale utilisee)"
    End If
End Function

Private Function EstOui(ByVal v As String) As Boolean
    v = LCase$(Trim$(v))
    EstOui = (v = "oui" Or v = "1" Or v = "vrai" Or v = "true" Or Len(v) = 0)
End Function

' Fiche correspondant au format modCourrier, a partir d'une ligne
' Specialistes (s, peut etre Nothing) et de sa ligne ParType (l).
Private Function CorrespondantDepuisSpecialiste(ByVal s As Object, ByVal l As Object) As Object
    Dim d As Object, tu As String, prenom As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("ID") = "SPT:" & l("ID_Ligne")
    d("NomDestinataire") = l("NomDestinataire")
    tu = LCase$(Trim$(Champ(l, "TutoiementVouvoiement")))
    If Len(tu) = 0 And Not s Is Nothing Then tu = LCase$(Trim$(Champ(s, "TutoiementVouvoiement")))
    If tu <> "tu" Then tu = "vous"
    d("Tutoiement") = tu
    If Not s Is Nothing Then
        d("Titre") = Champ(s, "Titre")
        prenom = Champ(s, "PrenomOuInitiale")
        If prenom = "." Then prenom = ""
        d("Prenom") = prenom
        d("Nom") = Champ(s, "Nom")
        d("Specialite") = Champ(s, "Structure")
        d("Adresse1") = Champ(s, "Adresse1")
        d("Adresse2") = Champ(s, "Adresse2")
        d("CP") = Champ(s, "CodePostal")
        d("Ville") = Champ(s, "Ville")
        d("Tel") = Champ(s, "Telephone")
        d("Structure") = Champ(s, "Structure")
        d("BlocDestinataire") = Replace(Champ(s, "BlocDestinataireComplet"), vbLf, vbCr)
    Else
        d("Titre") = "": d("Prenom") = "": d("Nom") = l("NomDestinataire")
        d("Specialite") = Champ(l, "Structure"): d("Adresse1") = "": d("Adresse2") = ""
        d("CP") = Champ(l, "CodePostal"): d("Ville") = Champ(l, "Ville"): d("Tel") = ""
        d("Structure") = Champ(l, "Structure")
        d("BlocDestinataire") = Replace(Champ(l, "BlocDestinataireComplet"), vbLf, vbCr)
    End If
    If Len(d("BlocDestinataire")) = 0 Then d("BlocDestinataire") = Replace(Champ(l, "BlocDestinataireComplet"), vbLf, vbCr)
    d("Email") = ""
    ' formules : celles de la ligne ParType, puis de la fiche, puis defauts du cabinet
    d("FormuleAppel") = Champ(l, "FormuleAppel")
    If Len(d("FormuleAppel")) = 0 And Not s Is Nothing Then d("FormuleAppel") = Champ(s, "FormuleAppel")
    If Len(d("FormuleAppel")) = 0 Then d("FormuleAppel") = modCourrier.AppelParDefaut(tu = "tu")
    d("FormulePolitesse") = Champ(l, "FormulePolitesse")
    If Len(d("FormulePolitesse")) = 0 Then d("FormulePolitesse") = modCourrier.PolitesseParDefaut(tu = "tu")
    d("Priorite") = Val(Champ(l, "Priorite"))
    d("Actif") = "1"
    Set CorrespondantDepuisSpecialiste = d
End Function

Private Function Champ(ByVal dict As Object, ByVal nom As String) As String
    If dict Is Nothing Then Exit Function
    If dict.Exists(nom) Then Champ = Trim$(CStr(dict(nom)))
End Function

Private Sub InsererParPriorite(ByVal col As Collection, ByVal d As Object)
    Dim i As Long
    For i = 1 To col.Count
        If col(i)("Priorite") > d("Priorite") Then
            col.Add d, , i
            Exit Sub
        End If
    Next i
    col.Add d
End Sub

' ---------------------------------------------------------------------
' Prompt systeme de la demande : gabarit derivee.txt + consignes du profil
' ---------------------------------------------------------------------
Public Function ConstruirePrompt(ByVal p As Object, ByVal modalite As String, ByVal tutoiement As Boolean, _
                                 ByVal identitePatient As String, ByVal motif As String, _
                                 ByVal phrasesPrescription As String) As String
    Dim systeme As String
    ' les lettres de consultation ne servent PAS de reference de style :
    ' c'est precisement ce qui produisait des demandes en forme de resume
    systeme = modApiConfiguration.ChargerPrompt("derivee.txt", False)
    systeme = Replace(systeme, "{{TYPE_DEMANDE}}", modDemandes.LibelleProfil(p, modalite))
    systeme = Replace(systeme, "{{CONSIGNES_TYPE}}", _
        modDemandes.ConsignesProfil(p, modalite, tutoiement, identitePatient, motif, phrasesPrescription))
    systeme = Replace(systeme, "{{PATIENT}}", identitePatient)
    systeme = Replace(systeme, "{{TUTOIEMENT}}", IIf(tutoiement, _
        "Le destinataire est un confrère proche : tutoiement (« Je te serais reconnaissant », « ton service »).", _
        "Le destinataire est vouvoyé (« Je vous serais reconnaissant », « votre service »)."))
    If Len(Trim$(motif)) > 0 Then
        systeme = Replace(systeme, "{{MOTIF}}", "Motif ou question précisé par le médecin (à reprendre tel quel) : " & Trim$(motif))
    Else
        systeme = Replace(systeme, "{{MOTIF}}", "Aucun motif particulier n'a été saisi : déduis-le sobrement de la phrase de prescription et du courrier source, sans le surinterpréter.")
    End If
    ConstruirePrompt = systeme
End Function

' Generation sans interface (testable). Renvoie Nothing si le medecin
' refuse l'envoi apres le scan residuel.
Public Function GenererDepuisProfil(ByVal docSource As Document, ByVal p As Object, _
                                    ByVal destNouveau As Object, ByVal modalite As String, _
                                    ByVal motif As String, ByVal phrasesPrescription As String) As Document
    Err.Raise vbObjectError + 988, "Compatibilite U2", "Moteur historique retire. Utilisez Unifie_D_Finaliser pour conserver les controles et reprises."
End Function
