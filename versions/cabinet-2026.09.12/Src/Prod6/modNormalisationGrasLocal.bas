Attribute VB_Name = "modNormalisationGrasLocal"
Option Explicit

'===============================================================================
' MODULE : modNormalisationGrasLocal
' VERSION : R5 - IDENTITE PATIENT EN GRAS LOCAL
'
' Traitement 100 % local, sans appel API :
' - lit les deux dictionnaires partagés du NAS ;
' - médicaments : forme canonique en MAJUSCULES + gras ;
' - expressions : gras uniquement, sans modifier la casse ;
' - s'applique au courrier principal ET aux lettres complémentaires déjà fusionnées.
' - impose localement le gras sur Monsieur/Madame NOM Prénom ;
' - l'identité du patient n'est jamais ajoutée aux dictionnaires NAS.
'
' Les dictionnaires restent sur le NAS afin que les apprentissages réalisés
' au secrétariat soient immédiatement utilisables sur les autres postes.
'===============================================================================

Private Const NGL_DOSSIER_CONFIG As String = _
    "\\DS224\home\logiciel_chatgpt_cabinet"

Private Const NGL_FICHIER_MEDICAMENTS As String = _
    "\\DS224\home\logiciel_chatgpt_cabinet\medicaments.txt"

Private Const NGL_FICHIER_EXPRESSIONS As String = _
    "\\DS224\home\logiciel_chatgpt_cabinet\termes_gras.txt"

Private Const NGL_VAR_IDENTITE_PATIENT As String = _
    "MCP_IDENTITE_PATIENT_V1"

Public Function NGL_NormaliserDocument(ByVal doc As Document) As Boolean
    Call modGras.AppliquerGrasDocumentComplet(doc)
    NGL_NormaliserDocument = True
End Function
Public Sub NGL_TesterConfiguration()
    MsgBox CStr(modGras.NbMedicaments()) & " medicaments et " & CStr(modGras.NbExpressions()) & " expressions dans les dictionnaires Cabinet.", vbInformation, "Gras"
End Sub






'-------------------------------------------------------------------------------
' IDENTITE DU PATIENT : REGLE STRUCTURELLE LOCALE
'-------------------------------------------------------------------------------












