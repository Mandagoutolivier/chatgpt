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

    def test_gdt_transmet_la_ddn_validee_sans_inventer_de_sexe(self):
        text = procedure(source("Src/Commun/modGdt.bas"), "ConstruireGdt")
        champs = re.findall(r'LigneGdt\("(\d{4})"', text)
        self.assertEqual(champs, ['8000', '9206', '9218', '3000', '3101', '3102', '3103', '8402', '8100'])
        self.assertIn('LigneGdt("3000", pat("ID"))', text)
        self.assertIn('ddn = modTexte.DdnPatient(pat)', text)
        self.assertIn('If Not modTexte.DateFrValide(ddn) Then', text)
        self.assertIn('naissance = modTexte.DateFr(ddn)', text)
        self.assertIn('If naissance > Date Then', text)
        self.assertIn('LigneGdt("3103", Format$(naissance, "dd.mm.yyyy"))', text)
        self.assertLess(text.index('DateFrValide(ddn)'), text.index('modTexte.DateFr(ddn)'))
        self.assertLess(text.index('If naissance > Date Then'), text.index('Set lignes = New Collection'))
        for conversion_regionale in ('CDate(', 'DateValue('):
            self.assertNotIn(conversion_regionale, text)
        self.assertNotIn('pat("Sexe")', text)
        identity = procedure(source("Src/Integration/modIntegrationUnifie.bas"), "PatientVerifie")
        self.assertIn('Array("Nom", "Prenom", "DDN", "Sexe")', identity)
        self.assertIn('DateFrValide', identity)
        self.assertIn('SexeNormalise', identity)

    def test_gdt_conserve_longueurs_crlf_et_ecriture_cp1252_atomique(self):
        gdt = source("Src/Commun/modGdt.bas")
        construction = procedure(gdt, "ConstruireGdt")
        self.assertIn('total = 14', construction)
        self.assertIn('total = total + Len(l) + 2', construction)
        self.assertIn('LigneGdt("8100", Format$(total, "00000")) & vbCrLf', construction)
        self.assertIn('contenu = contenu & l & vbCrLf', construction)
        ligne = procedure(gdt, "LigneGdt")
        self.assertIn('st.Charset = "windows-1252"', ligne)
        self.assertIn('If retour <> valeur Then Err.Raise', ligne)
        self.assertIn('Format$(Len(champ & valeur) + 5, "000")', ligne)
        ecriture = procedure(gdt, "EcrireGdtPatient")
        self.assertLess(ecriture.index('contenu = ConstruireGdt(pat)'), ecriture.index('EcrireTexteAnsi tmp, contenu'))
        self.assertIn('RenommerAtomique tmp, chemin, True', ecriture)

    def test_recette_gdt_couvre_dates_validees_et_rejets(self):
        audit = source('Tests/Vba/word/modAuditTests.bas')
        for date in ('15/01/1980', '15/11/1980', '29/02/1980', '31/12/1980', '1/2/1980', '29/02/2000'):
            self.assertIn('"' + date + '"', audit)
        for controle in ('GDT DDN avec points', 'GDT DDN courte normalisee', 'GDT DDN absente refusee',
                         'GDT DDN invalide refusee', 'GDT DDN future refusee', 'GDT DDN du jour acceptee',
                         'GDT sans sexe', 'GDT longueurs CRLF', 'GDT octets CP1252 sans BOM',
                         'GDT DDN refusee avant ecriture', 'GDT existant preserve apres refus DDN',
                         'GDT aucun temporaire apres refus DDN', 'GDT futur ne cree aucun import'):
            self.assertIn(controle, audit)
        for date in ('31/02/1960', '29/02/1900', '29/02/1981', '15.01.1980', '1980-01-15', '15011980'):
            self.assertIn('"' + date + '"', audit)

    def test_gdt_points_entree_et_limite_du_calcul_age_explicites(self):
        ecg = procedure(source('Src/Word/modEcg.bas'), 'EnvoyerECG')
        consultation = procedure(source('Src/Integration/modPowerMicUnifie.bas'), 'Unifie_DemarrerConsultation')
        for parcours in (ecg, consultation):
            self.assertIn('modGdt.EcrireGdtPatient(pat)', parcours.replace('modGdt.EcrireGdtPatient pat', 'modGdt.EcrireGdtPatient(pat)'))
        for consigne in ('F2 / Nouveau Patient', "controlez l'ID et la DDN", 'retablissez exactement la DDN', "Ce geste n'est pas automatise", 'renseignez le sexe'):
            self.assertIn(consigne, ecg)
        for commande in ('SendKeys', 'SendMessage', 'Shell', 'AppActivate'):
            self.assertNotIn(commande, ecg)

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
        for controle in ('GDT sans sexe', 'GDT identifiant long integral', 'GDT deux identifiants de meme prefixe distincts', 'GDT accent CP1252 conserve'):
            self.assertIn(controle, audit)


if __name__ == '__main__':
    unittest.main()
