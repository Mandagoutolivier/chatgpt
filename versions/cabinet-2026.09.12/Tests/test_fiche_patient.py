"""Contrat HTTP et controles d integration du nouveau formulaire Word."""
import json
from pathlib import Path
import sys
import unittest
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'Serveur'))
from cabinet.contract import validate
from cabinet.domain import Refus


class FichePatient(unittest.TestCase):
    def test_contrat_accepte_consultation_et_revision_texte(self):
        validate('patient.update', {'consultation_id': 'FICTIF-C', 'data': {'ID': 'FICTIF-P', '_revision': '1', 'Prenom': 'Fictif'}})

    def test_contrat_refuse_champs_et_types_inattendus(self):
        for params in ({'data': {}}, {'consultation_id': '', 'data': {}},
                       {'consultation_id': 'FICTIF-C', 'genre': 'ACTES', 'data': {}},
                       {'consultation_id': 'FICTIF-C', 'data': {'DDN': 1980}}):
            with self.subTest(params=params), self.assertRaises(Refus):
                validate('patient.update', params)

    def test_ruban_et_formulaire_sont_construits_dans_word(self):
        manifest = json.loads((ROOT / 'Build/manifest.json').read_text())
        form = next(e for e in manifest['word'] if e['name'] == 'ufPatientMedecin')
        controls = {c['name'] for c in form['designer']['controls']}
        fields = ('Nom', 'Prenom', 'NomNaissance', 'DDN', 'Sexe', 'NIR', 'Adresse1', 'Adresse2',
                  'CP', 'Ville', 'Tel', 'Mobile', 'Email', 'Mutuelle', 'ALD', 'Notes')
        self.assertTrue({'txt' + field for field in fields} <= controls)
        self.assertTrue({'btnEnregistrer', 'btnFermer', 'btnConfirmerIdentite', 'btnMedecin'} <= controls)
        self.assertEqual(len(controls), len(form['designer']['controls']))
        ribbon = ET.parse(ROOT / 'Build/ruban_unifie.xml')
        buttons = [e for e in ribbon.iter() if e.attrib.get('id') == 'cabFichePatient']
        self.assertEqual(len(buttons), 1)
        self.assertEqual(buttons[0].attrib['label'], 'Fiche patient')


if __name__ == '__main__':
    unittest.main()
