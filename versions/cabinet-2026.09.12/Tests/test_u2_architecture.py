import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
import audit_statique
from inventaire_architecture import inventorier
from cabinet import APPLICATION_VERSION, PROTOCOL_VERSION, SERVICE_REVISION, TARGET_SCHEMA_VERSION


class ArchitectureU2(unittest.TestCase):
    def test_revision_et_environnement_isole_sont_coherents(self):
        root=audit_statique.ROOT
        release=json.loads((root/'Build/release.json').read_text(encoding='utf-8'))
        self.assertEqual(release, {
            'release':SERVICE_REVISION,'version':APPLICATION_VERSION,'statut':'recette-isolee',
            'protocole':PROTOCOL_VERSION,'schema':TARGET_SCHEMA_VERSION,
            'revision_service':SERVICE_REVISION})
        env={}
        for line in (root/'Serveur/.env.u2-test.example').read_text(encoding='utf-8').splitlines():
            if line and not line.startswith('#') and '=' in line:
                key,value=line.split('=',1);env[key]=value
        self.assertEqual(env['CABINET_PROJECT_NAME'],'cabinetcardio-test-u2')
        self.assertEqual(env['CABINET_API_PORT'],'8766')
        self.assertEqual(env['CABINET_MAINTENANCE_IMAGE'],'cabinet-maintenance:'+SERVICE_REVISION)
        self.assertEqual(env['CABINET_DATA_VOLUME'],'/volume1/CabinetCardioTestU2')
        self.assertEqual(env['CABINET_DB_VOLUME'],'/volume1/docker/cabinetcardio-test-u2/postgres')
        self.assertEqual(env['CABINET_BACKUP_VOLUME'],'/volume1/docker/cabinetcardio-test-u2/sauvegardes')
        self.assertEqual(env['CABINET_UNC'],r'\\NAS-RECETTE\CabinetCardioTestU2')
        compose=(root/'Serveur/compose.yaml').read_text(encoding='utf-8')
        self.assertIn('name: ${CABINET_PROJECT_NAME:',compose)
        self.assertIn('127.0.0.1:${CABINET_API_PORT:',compose)
        self.assertIn("user: ${CABINET_UID:?Renseigner l UID API}:${CABINET_GID:?Renseigner le GID API}",compose)
        self.assertNotIn('127.0.0.1:8765:8765',compose)
        for name in ('compose.yaml','compose.recette.yaml','compose.verification.yaml'):
            self.assertNotIn('cabinet-maintenance:2026.09.14-u0',(root/'Serveur'/name).read_text(encoding='utf-8'))

    def test_contrat_recette_office_est_declare_par_les_suites(self):
        root=audit_statique.ROOT
        word=(root/'Tests/Vba/word/modRecetteU2.bas').read_text(encoding='utf-8-sig')
        excel=(root/'Tests/Vba/excel/modRecetteU1Excel.bas').read_text(encoding='utf-8-sig')
        self.assertRegex(word,r'NOMBRE_ATTENDU_U2 As Long = 17\b')
        self.assertRegex(excel,r'NOMBRE_ATTENDU_EXCEL As Long = 25\b')
        for source in (word,excel):
            self.assertIn('""attendus"":',source)
            self.assertRegex(source,r'(?:reussis|mNombre) <> NOMBRE_ATTENDU_')
        runner=(root/'Build/Tester_U1_Office.ps1').read_text(encoding='utf-8')
        helper=(root/'Build/outils_recette_u1.ps1').read_text(encoding='utf-8')
        self.assertNotRegex(runner,r'\.reussis\s+-ne\s+10')
        self.assertIn('[int]$recette.reussis -lt 34',runner)
        self.assertEqual(runner.count('Verifier-ResultatRecetteOffice'),2)
        self.assertIn("@('reussis','attendus','echec')",helper)

    def test_export_et_reimpression_exigent_un_contrat_explicite(self):
        root=audit_statique.ROOT
        config=(root/'Src/ConfigDefaut/config.ini').read_text(encoding='utf-8')
        self.assertRegex(config,r'(?ms)^\[SORTIE\].*^ExportActif=[01]$.*^Dossier=.+$.*^NomFichier=(?:PublicationID|IdentitePublication)$')
        sortie=(root/'Src/Prod6/modSortieDragon.bas').read_text(encoding='utf-8-sig')
        self.assertIn('Config("SORTIE", "ExportActif", "")',sortie)
        self.assertIn('Config("SORTIE", "NomFichier", "")',sortie)
        self.assertNotIn('Config("SORTIE", "ExportActif", "0")',sortie)
        actes=(root/'Src/Excel/modActes.bas').read_text(encoding='utf-8-sig')
        debut=actes.index('Public Sub ImprimerSeanceEnregistree')
        reimpression=actes[debut:]
        self.assertIn('DonneesImpressionFigees',reimpression)
        self.assertNotIn('LireID("PATIENTS"',reimpression)
        self.assertIn('CodeCerfa',actes)

    def test_inventaires_commis_sont_a_jour(self):
        erreurs,sources=audit_statique.verifier(False)
        self.assertEqual(erreurs,[])
        # Les condensats gardent le controle de fraicheur sans republier les inventaires internes detailles.
        inventaire_sources=json.dumps(sources,sort_keys=True,ensure_ascii=False,separators=(',',':')).encode('utf-8')
        self.assertEqual(hashlib.sha256(inventaire_sources).hexdigest(),'d9a6645005fd8a0314ecca85df8a9d82551f93815400a5da8b5780cf5c43b933')
        inventaire=json.dumps(inventorier(),sort_keys=True,ensure_ascii=False,separators=(',',':')).encode('utf-8')
        self.assertEqual(hashlib.sha256(inventaire).hexdigest(),'857a0ee73cee6a592d3e6a934ee98f3eb64159adf753a1137dad32613035f0ff')

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
