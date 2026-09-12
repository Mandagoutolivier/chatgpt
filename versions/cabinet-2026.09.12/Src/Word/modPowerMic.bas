Attribute VB_Name = "modPowerMic"
Option Explicit
' =====================================================================
' modPowerMic - Points d'entree des boutons du PowerMic (Dragon).
' Dragon execute une commande VBA du type :
'     wd.Run "Cabinet_A_NouvelleLettre"
' Les noms sont propres a Cabinet.dotm (l'ancien complement garde les
' siens, PowerMic_A_NouvelleLettre...) : aucune ambiguite entre les deux.
'
'   A : nouveau courrier pour le patient ARRIVE (medecin traitant de la
'       fiche) ; curseur dans le bloc adresse s'il est vide, sinon dans
'       le corps. Zero clavier.
'   B : formule d'appel selectionnee -> la dictee la remplace.
'   D : finaliser = corriger + lettres de demande a la suite + enregistrer
'       (dossier patient, sortiedragon) + secretariat.
'   P : identite du patient au curseur (equivalent F6).
' =====================================================================

Public Sub Cabinet_A_NouvelleLettre()
    modPowerMicUnifie.Unifie_A_NouvelleLettre
End Sub
Public Sub Cabinet_B_FormuleAppel()
    modPowerMicUnifie.Unifie_B_FormuleAppel
End Sub
Public Sub Cabinet_D_Finaliser()
    modPowerMicUnifie.Unifie_D_Finaliser
End Sub
Public Sub Cabinet_P_Patient()
    modPowerMicUnifie.Unifie_C_InsererPatient
End Sub
Public Sub Cabinet_Destinataire()
    modCourrier.AllerDestinataire
End Sub

Public Sub Cabinet_Corps()
    modCourrier.AllerCorps
End Sub
