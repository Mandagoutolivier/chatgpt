import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
import audit_statique
from inventaire_architecture import inventorier


class ArchitectureU2(unittest.TestCase):
    def test_production_et_recette_sont_coherentes(self):
        for recette in [False,True]:
            erreurs,inventory=audit_statique.verifier(recette)
            self.assertEqual(erreurs,[])
            self.assertTrue(inventory)
        result=inventorier()
        self.assertEqual(result['erreurs'],[])
        self.assertEqual(len(result['archives_exclues']),3)

    def test_reference_production_vers_recette_est_refusee(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)/'copie';shutil.copytree(audit_statique.ROOT,root)
            path=root/'Src/Word/modPowerMic.bas'
            with path.open('a') as f:f.write('\nPublic Sub InjectionFictive()\n    modRecetteU1.Executer "FICTIF"\nEnd Sub\n')
            with patch.object(audit_statique,'ROOT',root):
                erreurs,_=audit_statique.verifier()
                self.assertTrue(any('module absent modRecetteU1' in e for e in erreurs))

    def test_composant_inconnu_est_refuse(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)/'copie';shutil.copytree(audit_statique.ROOT,root)
            (root/'Src/Word/modFictif.bas').write_text('Attribute VB_Name = "modFictif"\n')
            self.assertTrue(any('modFictif' in e for e in inventorier(root)['erreurs']))

    def test_analyse_preserve_commentaires_et_litteraux(self):
        self.assertEqual(audit_statique.sans_litteraux('x = "a""b" \' modAbsent.Faux'), 'x = "" ')
        with self.assertRaises(ValueError):audit_statique.sans_litteraux('x="chaine ouverte')


if __name__=='__main__':unittest.main()
