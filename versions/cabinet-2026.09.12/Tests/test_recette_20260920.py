import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def source(path):
    return (ROOT / path).read_text(encoding="utf-8-sig")


def procedure(text, name):
    m = re.search(r"(?ms)^(?:Public|Private) (?:Sub|Function) " + re.escape(name) + r"\b.*?^End (?:Sub|Function)$", text)
    if not m:
        raise AssertionError("Procedure absente : " + name)
    return m[0]


class Recette20Septembre(unittest.TestCase):
    def test_correction_ne_fabrique_pas_une_relecture(self):
        text = procedure(source("Src/Integration/modControleCourrier.bas"), "PreparerRelecture")
        for interdit in ('TransmettreSecretariat', 'ValiderEtTransmettre', '"RelectureValidee", "1"', '"publish"'):
            self.assertNotIn(interdit, text)
        self.assertIn('InvaliderRelecture doc', text)
        self.assertIn('"RelectureEnAttente", "1"', text)
        self.assertLess(text.index('RemplirSignet doc, "DESTINATAIRE"'), text.index('"RelectureEnAttente", "1"'))
        self.assertLess(text.index('"RelectureEnAttente", "1"'), text.index('doc.Save'))

    def test_d_separe_correction_et_unique_validation(self):
        text = procedure(source("Src/Integration/modPowerMicUnifie.bas"), "Unifie_D_Finaliser")
        self.assertIn('"RelectureEnAttente")) = "1" Then', text)
        self.assertEqual(text.count('modControleCourrier.ValiderEtTransmettre'), 1)
        self.assertEqual(text.count('modProdRapide.PR_CorrigerToutEnUnClic'), 1)
        self.assertIn('Else', text)
        self.assertIn('If mOccupe Or modProdRapide.PR_EnCours() Then Exit Sub', text)

    def test_transmission_sans_confirmation_redondante_ni_substitution(self):
        text = procedure(source("Src/Integration/modControleCourrier.bas"), "ValiderEtTransmettre")
        self.assertNotIn('vbYesNo', text)
        self.assertNotIn('AssurerDestinataire(doc)', text)
        self.assertNotIn('RemplirSignet', text)
        self.assertLess(text.index('"RelectureEnAttente"'), text.index('TransmettreSecretariat'))
        self.assertLess(text.index('DestinataireRelectureConforme'), text.index('"RelectureValidee", "1"'))
        self.assertLess(text.index('"RelectureValidee", "1"'), text.index('TransmettreSecretariat'))
        self.assertIn('If Trim$(modIntegrationUnifie.VariableDoc(doc, "PublicationRevision")) <> revision Then', text)
        self.assertIn('"RelectureValidee", "0"', text.split('Echec:', 1)[1])

    def test_ancien_droit_de_validation_est_invalide_avant_correction(self):
        text = procedure(source("Src/Integration/modCycleCourrier.bas"), "ExecuterCycleCourrier")
        self.assertLess(text.index('PrendreVerrou'), text.index('InvaliderRelecture'))
        self.assertLess(text.index('InvaliderRelecture'), text.index('SauvegarderBrouillon'))
        self.assertLess(text.index('InvaliderRelecture'), text.index('AppelerOpenAIStructure'))

    def test_gdt_minimal_ne_lit_pas_les_champs_ecartes(self):
        text = procedure(source("Src/Commun/modGdt.bas"), "ConstruireGdt")
        champs = re.findall(r'LigneGdt\("(\d{4})"', text)
        self.assertEqual(champs, ['8000', '9206', '9218', '3000', '3101', '3102', '8402', '8100'])
        self.assertIn('LigneGdt("3000", pat("ID"))', text)
        self.assertNotIn('pat("DDN")', text)
        self.assertNotIn('pat("Sexe")', text)
        identity = procedure(source("Src/Integration/modIntegrationUnifie.bas"), "PatientVerifie")
        self.assertIn('Array("Nom", "Prenom", "DDN", "Sexe")', identity)
        self.assertIn('DateFrValide', identity)
        self.assertIn('SexeNormalise', identity)

    def test_nir_facultatif_sans_invention_ni_appel_vide(self):
        text = source("Src/Excel/modCerfaPrint.bas")
        helper = procedure(text, "NirFacultatifPourFeuille")
        self.assertLess(helper.index('If Len(valeur) = 0 Then Exit Function'), helper.index('Appeler("nir.validate"'))
        self.assertIn('valeur = Trim$(valeur)', helper)
        self.assertIn('CStr(r("nir"))', helper)
        self.assertIn('valeurs("ASSURE_NIR") = NirFacultatifPourFeuille(nir)', procedure(text, 'ImprimerFeuille'))
        frozen = procedure(text, 'VerifierAvantReimpression')
        self.assertIn('For Each cle In Array("PatientID", "Nom", "Prenom", "DDN")', frozen)
        self.assertNotIn('"NIR"', frozen)

    def test_nir_optionnel_ne_supprime_pas_les_protections_impression(self):
        text = procedure(source("Src/Excel/modCerfaPrint.bas"), 'ImprimerFeuille')
        for controle in ('CalageValide', 'PraticienPreimprime', 'NumeroAM', 'RPPS', 'DateFrValide', 'actes.Count > 4', 'Fiche de l assure incomplete'):
            self.assertIn(controle, text)
        self.assertIn('ImprimerDocumentCale', text)
        form = source('Src/Excel/ufChoixActe.vba')
        self.assertIn('chkFds.Value = False', form)
        self.assertIn('If chkFds.Value Then modCerfaPrint.VerifierAvantFacturation', form)
        self.assertIn('TraitementPeutEtreCloture', form)

    def test_recettes_couvrent_nir_vide_et_changement_destinataire(self):
        excel = source('Tests/Vba/excel/modRecetteU1Excel.bas')
        for controle in ('NirFacultatifPourFeuille("")', 'NirFacultatifPourFeuille("   ")', 'ligneFigee.Remove "NIR"', 'ligneFigee.Remove "AssureNIR"'):
            self.assertIn(controle, excel)
        word = source('Tests/Vba/word/modRecetteU2.bas')
        self.assertIn('TesterDestinataireRelecture', word)
        self.assertIn('changement adresse serveur refuse', word)
        audit = source('Tests/Vba/word/modAuditTests.bas')
        for controle in ('GDT sans sexe ni naissance', 'GDT identifiant long integral', 'GDT deux identifiants de meme prefixe distincts', 'GDT accent CP1252 conserve'):
            self.assertIn(controle, audit)


if __name__ == '__main__':
    unittest.main()
