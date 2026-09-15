Attribute VB_Name = "modGlobals"
Option Explicit

Public Type PatientInfo
    civilite As String
    nom As String
    prenom As String
    Age As String
    NomComplet As String
    marqueur As String
End Type

Public gPatient As PatientInfo

Public gTexteAnonymise As String
Public gTexteCorrige As String
Public gPlageOriginale As Range

Public gDemandeExamenAnonymisee As String
Public gDemandeExamenRestauree As String

Public Const MARQUEUR_PATIENT As String = "[[PATIENT]]"
Public Const APP_VERSION As String = _
    "ModeleCourrierChatGPT CABINET PROD RAPIDE v1"

Public Const BALISE_DEBUT_CORPS As String = "---CORPS_COURRIER---"
Public Const BALISE_FIN_CORPS As String = "---FIN_CORPS_COURRIER---"

Public Const BALISE_DEBUT_DEMANDE_EXAMEN As String = "---DEMANDE_EXAMEN---"
Public Const BALISE_FIN_DEMANDE_EXAMEN As String = "---FIN_DEMANDE_EXAMEN---"

Public Const OPENAI_API_URL As String = "https://api.openai.com/v1/responses"
Public Const OPENAI_MODEL As String = "gpt-4.1-mini"



Public Const FEUILLE_BASE_GENERALISTES As String = "Generalistes"
Public Const FEUILLE_BASE_SPECIALISTES As String = "Specialistes"
Public Const FEUILLE_BASE_SPECIALISTES_PARTYPE As String = "Specialistes_ParType"

Public Const BALISE_DEBUT_DEMANDE_DESTINATION As String = "---DEMANDE_DESTINATION---"
Public Const BALISE_FIN_DEMANDE_DESTINATION As String = "---FIN_DEMANDE_DESTINATION---"

Public Const PREFIXE_CLE_DESTINATION As String = "CLE_DESTINATION="
Public Const BALISE_DEBUT_CORPS_DESTINATION As String = "---CORPS_DESTINATION---"
Public Const BALISE_FIN_CORPS_DESTINATION As String = "---FIN_CORPS_DESTINATION---"

    '=========================================================
' État de l'interface CABINET TEST
'=========================================================


Public gReponseAPICabinetTest As String
Public gCorpsCorrigeAnonymiseCabinetTest As String

Public gApercuCabinetTestActif As Boolean
Public gCorrectionCabinetTestValidee As Boolean

Public gNombreDemandesPretesCabinetTest As Long
Public gNombreDestinationsACompleterCabinetTest As Long
Public gDernieresDestinationsACompleterCabinetTest As String

Public gReponseAPICabinetTestAvantTest As String

Public gCleDestinationChoisieCabinetTest As String
Public gDemandesMultipagesAjouteesProdRapide As Boolean

'=========================================================
' Sécurités de la branche CABINET TEST
'=========================================================

Public Const MODE_CABINET_TEST As Boolean = True

Public Const AUTORISER_MODIFICATION_DIRECTE_ORIGINAL As Boolean = False
Public Const AUTORISER_SAUVEGARDE_AUTOMATIQUE As Boolean = False
Public Const AUTORISER_IMPRESSION_AUTOMATIQUE As Boolean = False

Public Const EXIGER_VALIDATION_AVANT_INSERTION As Boolean = True
Public Const EXIGER_VALIDATION_AVANT_CREATION_DEMANDES As Boolean = True

Public Const REFUSER_BLOC_DEMANDE_INCOMPLET As Boolean = True
Public Const REFUSER_DESTINATION_NON_RESOLUE As Boolean = True

Public Sub ReinitialiserContexteTraitement()

    gPatient.civilite = ""
    gPatient.nom = ""
    gPatient.prenom = ""
    gPatient.Age = ""
    gPatient.NomComplet = ""
    gPatient.marqueur = MARQUEUR_PATIENT

    gTexteAnonymise = ""
    gTexteCorrige = ""

    Set gPlageOriginale = Nothing

    gDemandeExamenAnonymisee = ""
    gDemandeExamenRestauree = ""




End Sub


Public Function CHEMIN_BASE_CORRESPONDANTS() As String
    CHEMIN_BASE_CORRESPONDANTS = modConfig.CheminNasConfigure("PROD6", "BaseCorrespondants", "Base\base_travail_correspondants_v1.xlsx")
End Function
