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
Private Function ChoisirProfil() As String
    Dim items As Collection, f As ufListe
    Set items = modDemandes.ListerProfils(modConfig.Config("DERIVEES", "Types", ""))
    If items.Count = 0 Then
        Err.Raise vbObjectError + 502, "modDerivees", _
            "Aucun profil de demande dans " & modDemandes.DossierProfils()
    End If
    Set f = New ufListe
    f.Configurer "Type de demande (tapez pour filtrer)", items, Array("Libelle", "Code"), "220 pt;150 pt"
    f.Show vbModal
    If Not f.Annule Then ChoisirProfil = f.Resultat("Code")
    Unload f
End Function

' Modalite du profil (libelle). "" si le profil n'en a pas, vbNullChar si annulation.
Private Function ChoisirModalite(ByVal p As Object) As String
    Dim items As Collection, f As ufListe
    Set items = modDemandes.ModalitesProfil(p)
    If items.Count = 0 Then ChoisirModalite = "": Exit Function
    Set f = New ufListe
    f.Configurer modDemandes.LibelleProfil(p) & " : modalite", items, Array("Libelle"), "280 pt"
    f.Show vbModal
    If f.Annule Then ChoisirModalite = vbNullChar Else ChoisirModalite = f.Resultat("Libelle")
    Unload f
End Function

' ---------------------------------------------------------------------
' Destinataire
' ---------------------------------------------------------------------
Private Function ChoisirDestinataire(ByVal p As Object, ByVal modalite As String) As Object
    Dim items As Collection, f As ufListe, titre As String
    titre = "Destinataire : " & modDemandes.LibelleProfil(p, modalite)
    Set items = CorrespondantsPourTypes(modDemandes.ValeurProfil(p, "IDENTITE", "TYPE_BASE"))
    If items.Count > 0 Then
        Set f = New ufListe
        f.Configurer titre & "  (Nouveau... = autre correspondant)", items, _
                     Array("NomDestinataire", "Structure", "Ville"), "160 pt;150 pt;80 pt", "", True
        f.Show vbModal
        If Not f.Annule Then
            Set ChoisirDestinataire = f.Resultat
            Unload f
            Exit Function
        End If
        If Not f.NouveauDemande Then Unload f: Exit Function
        Unload f
    End If
    ' pas de specialiste reference pour ce type, ou "autre correspondant"
    Set ChoisirDestinataire = modPatient.ChoisirCorrespondant("", titre)
End Function

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
    Dim res As Collection, fichier As String, specialistes As Collection, parType As Collection
    Dim codes() As String, i As Long, s As Object, l As Object, d As Object
    Dim parId As Object, vus As Object, cle As String
    Set res = New Collection
    fichier = FichierSpecialistes()
    If Len(fichier) = 0 Or Len(Trim$(typesExamen)) = 0 Then Set CorrespondantsPourTypes = res: Exit Function

    Set specialistes = modBase.LireTable(fichier, modConfig.Config("DERIVEES", "FeuilleSpecialistes", "Specialistes"), "ID")
    Set parType = modBase.LireTable(fichier, modConfig.Config("DERIVEES", "FeuilleParType", "Specialistes_ParType"), "ID_Ligne")
    Set parId = CreateObject("Scripting.Dictionary")
    parId.CompareMode = 1
    For Each s In specialistes
        If Not parId.Exists(s("ID")) Then Set parId(s("ID")) = s
    Next s

    Set vus = CreateObject("Scripting.Dictionary")
    vus.CompareMode = 1
    codes = Split(typesExamen, ",")
    For i = LBound(codes) To UBound(codes)
        For Each l In parType
            If StrComp(Trim$(l("TypeExamen")), Trim$(codes(i)), vbTextCompare) = 0 And EstOui(l("Actif")) Then
                cle = l("ID_Specialiste") & "|" & l("NomDestinataire")
                If Not vus.Exists(cle) Then
                    vus(cle) = 1
                    If parId.Exists(l("ID_Specialiste")) Then
                        Set s = parId(l("ID_Specialiste"))
                    Else
                        Set s = Nothing
                    End If
                    Set d = CorrespondantDepuisSpecialiste(s, l)
                    InsererParPriorite res, d
                End If
            End If
        Next l
    Next i
    Set CorrespondantsPourTypes = res
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
    systeme = modClaude.ChargerPrompt("derivee.txt", False)
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
    Dim pat As Object, corSource As Object, ctx As Object
    Dim corps As String, anonyme As String, problemes As String
    Dim systeme As String, reponse As String, final As String
    Dim nouveauDoc As Document, prog As ufProgression
    Dim identite As String, tutoiement As Boolean, libelle As String, phrasesAnonymes As String
    Dim motifAnonyme As String

    Set pat = modClaude.PatientDuDocument(docSource)
    Set corSource = modClaude.CorrespondantDuDocument(docSource)
    corps = modCourrier.RecupererCorps(docSource)
    If Len(Trim$(corps)) < 10 Then
        Err.Raise vbObjectError + 500, "modDerivees", "Le courrier source est vide."
    End If
    libelle = modDemandes.LibelleProfil(p, modalite)

    Set ctx = modAnonymise.Construire(pat, corSource)
    ' le nouveau destinataire peut aussi etre cite dans le courrier source
    If Not destNouveau Is Nothing Then
        modAnonymise.AjouterCorrespondant ctx, destNouveau, "DEST2"
    End If
    anonyme = modAnonymise.Anonymiser(corps, ctx)
    ' identite du patient (civilite, prenom, nom, age) et phrases de
    ' prescription : balisees elles aussi
    identite = modAnonymise.Anonymiser(modCourrier.TexteIdentitePatient(pat), ctx)
    phrasesAnonymes = modAnonymise.Anonymiser(phrasesPrescription, ctx)
    ' le motif est une saisie libre du medecin : il est balise et scanne
    ' comme le reste avant de partir vers l'API
    motifAnonyme = modAnonymise.Anonymiser(motif, ctx)
    problemes = modAnonymise.ScanResiduel(anonyme & vbCr & identite & vbCr & _
                                          phrasesAnonymes & vbCr & motifAnonyme, ctx)
    If Len(problemes) > 0 Then
        If MsgBox("L'anonymisation a detecte un risque avant envoi (" & libelle & ") :" & vbCrLf & vbCrLf & _
                  problemes & vbCrLf & "Envoyer QUAND MEME a l'API ?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Cabinet - protection des donnees") <> vbYes Then
            Exit Function
        End If
    End If

    tutoiement = modCourrier.EstTutoye(destNouveau)
    systeme = ConstruirePrompt(p, modalite, tutoiement, identite, motifAnonyme, phrasesAnonymes)

    Set prog = New ufProgression
    prog.Show vbModeless
    prog.Definir "Redaction de la demande : " & libelle & "..."
    On Error GoTo ErreurApi
    reponse = modClaude.AppelerClaude(systeme, "Patient : " & identite & vbLf & vbLf & _
                                      "Courrier de consultation source :" & vbLf & anonyme)
    On Error GoTo 0
    Unload prog
    Set prog = Nothing

    problemes = modAnonymise.VerifierBalisesRetour(reponse, ctx, _
                    identite & vbCr & anonyme & vbCr & phrasesAnonymes & vbCr & motifAnonyme)
    If Len(problemes) > 0 Then
        modFichiers.EcrireTexteUTF8 modConfig.chemin("Logs") & "\reponse_rejetee.txt", reponse
        Err.Raise vbObjectError + 501, "modDerivees", _
            "La reponse de l'API a altere des balises d'identite :" & vbCrLf & problemes
    End If
    final = modAnonymise.Reinjecter(reponse, ctx)

    Set nouveauDoc = modCourrier.CreerCourrierPour(pat, destNouveau, "demande - " & libelle)
    nouveauDoc.Variables("ProfilDemande") = p("Code")
    modCourrier.RemplacerCorps nouveauDoc, modClaude.NettoyerReponse(final)
    On Error Resume Next
    modGras.AppliquerGras nouveauDoc
    If Err.Number <> 0 Then modLog.LogErreur "Gras lettre derivee : " & Err.Description
    On Error GoTo 0
    Application.StatusBar = "Demande '" & libelle & "' generee : relisez avant validation."
    Set GenererDepuisProfil = nouveauDoc
    Exit Function
ErreurApi:
    If Not prog Is Nothing Then Unload prog
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
