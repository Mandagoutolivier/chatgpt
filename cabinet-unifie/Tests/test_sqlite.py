"""Execute le SQL extrait de la source VBA avec SQLite.
Verifie les transactions et les cles; ne pretend pas executer VBA ou cmd.exe.
"""
from pathlib import Path
import re, sqlite3, tempfile, unittest
ROOT=Path(__file__).resolve().parents[1]
SOURCE=(ROOT/'Src/Integration/modAttenteLocale.bas').read_text()

def litteraux(expression):
    return ''.join(x[1:-1].replace('""','"') for x in re.findall(r'"(?:[^"]|"")*"',expression))

PREFIX=litteraux('"PRAGMA'+SOURCE.split('sql = "PRAGMA',1)[1].split('    For Each f',1)[0])
SUFFIX=litteraux('"COMMIT;'+SOURCE.split('sql = sql & "COMMIT;',1)[1].split('    ExecuterSql',1)[0])
# Les parametres de test sont lies par SQLite; le quoting VBA est un essai Windows distinct.
INSERT='INSERT INTO attentes_v2 VALUES('+','.join(['?']*13)+')'
def executer_fragment(db, sql):
    # execute() conserve BEGIN/COMMIT explicites, contrairement au pre-COMMIT de executescript().
    for statement in sql.split(";"):
        if statement.strip():db.execute(statement)

def ligne(rdv,nom="FICTIF"):
    return ('TEST-001',nom,'Elodie','29/02/1960','F','',rdv,'11/09/2026','10:00','09:55','Arrive','NAS fictif','2026-09-11')

class CacheLocal(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.path=Path(self.temp.name)/'test.sqlite'
        self.db=sqlite3.connect(self.path,timeout=0)
        executer_fragment(self.db,PREFIX)
        self.db.execute(INSERT,ligne('ANCIEN'))
        executer_fragment(self.db,SUFFIX)
    def tearDown(self):
        self.db.close();self.temp.cleanup()
    def test_un_patient_deux_rendez_vous(self):
        executer_fragment(self.db,PREFIX)
        self.db.execute(INSERT,ligne('RDV-1'));self.db.execute(INSERT,ligne('RDV-2'))
        executer_fragment(self.db,SUFFIX)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM attentes_v2').fetchone()[0],2)
    def test_erreur_annule_aussi_la_suppression(self):
        executer_fragment(self.db,PREFIX)
        self.db.execute(INSERT,ligne('DOUBLE'))
        with self.assertRaises(sqlite3.IntegrityError):self.db.execute(INSERT,ligne('DOUBLE'))
        # Comme sqlite3.exe -bail : la fermeture sans COMMIT annule la transaction.
        self.db.close();self.db=sqlite3.connect(self.path,timeout=0)
        self.assertEqual(self.db.execute('SELECT rdv_id FROM attentes_v2').fetchall(),[('ANCIEN',)])
    def test_lecteur_ne_voit_pas_la_suppression_non_validee(self):
        executer_fragment(self.db,PREFIX)
        reader=sqlite3.connect(self.path,timeout=0)
        try:self.assertEqual(reader.execute('SELECT rdv_id FROM attentes_v2').fetchall(),[('ANCIEN',)])
        finally:reader.close()
    def test_deux_ecrivains_non_simultanes(self):
        executer_fragment(self.db,PREFIX)
        other=sqlite3.connect(self.path,timeout=0)
        try:
            with self.assertRaises(sqlite3.OperationalError):other.execute('BEGIN IMMEDIATE')
        finally:other.close()

if __name__=='__main__':
    print('SQLite de verification :',sqlite3.sqlite_version)
    unittest.main(verbosity=2)
