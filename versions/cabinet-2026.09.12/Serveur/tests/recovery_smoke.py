"""Recette PostgreSQL 17 native, seulement dans compose.verification.yaml."""
from pathlib import Path
from contextlib import contextmanager
from cabinet.migration_u1 import migrer
import hashlib
import json
import os
import uuid
import zipfile
from psycopg import sql
from psycopg.types.json import Jsonb
from cabinet import recovery as r


def main():
    r.require(os.environ['PGDATABASE'] == 'cabinet_verification', 'Base de recette obligatoire.')
    source = Path('/source'); source.mkdir()
    config = Path('/config-source'); config.mkdir()
    backups = Path('/backups-test'); backups.mkdir()
    (source/'Documents').mkdir(); (source/'Patients').mkdir()
    (config/'.env').write_text('CABINET_UNC=\\\\NAS-FICTIF\\Source\n')
    (config/'secrets').mkdir(mode=0o700)
    (config/'secrets'/'fictif.txt').write_text('SECRET FICTIF DE RECETTE')
    os.chmod(config/'secrets'/'fictif.txt', 0o600)
    docx = source/'Patients'/'fictif.docx'
    with zipfile.ZipFile(docx, 'w') as z:
        z.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr('word/document.xml', '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>RECETTE U0 FICTIVE</w:t></w:r></w:p></w:body></w:document>')
    # Une page PDF fictive complete, avec table de references et positions calculees.
    stream = b'BT /F1 12 Tf 36 780 Td (RECETTE U0 FICTIVE) Tj ET\n'
    objects = [b'<< /Type /Catalog /Pages 2 0 R >>',
               b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
               b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
               b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
               b'<< /Length '+str(len(stream)).encode()+b' >>\nstream\n'+stream+b'endstream']
    data = b'%PDF-1.4\n'; offsets=[]
    for i, obj in enumerate(objects, 1):
        offsets.append(len(data)); data+=str(i).encode()+b' 0 obj\n'+obj+b'\nendobj\n'
    xref=len(data); data+=b'xref\n0 6\n0000000000 65535 f \n'
    for offset in offsets: data+=f'{offset:010d} 00000 n \n'.encode()
    data+=b'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n'+str(xref).encode()+b'\n%%EOF\n'
    pdf = source/'Patients'/'fictif.pdf'; pdf.write_bytes(data)
    old = r'\\NAS-FICTIF\Source'; new = r'\\NAS-RECETTE\AutrePartage'
    publication = {}
    for key, path, hash_key in [('CheminDocx',docx,'sha_docx'),('CheminPdf',pdf,'sha_pdf')]:
        sha = r.digest(path); name = sha+path.suffix
        (source/'Documents'/name).write_bytes(path.read_bytes())
        os.chmod(source/'Documents'/name, 0o640)
        publication[key] = old+'\\Documents\\'+name; publication[hash_key] = sha
    with r.connection() as db:
        db.execute(Path('/tests/schema.sql').read_text())
        db.execute("INSERT INTO comptes VALUES ('compte-fictif', %s, '[\"medecin\"]', true)", (hashlib.sha256(b'FICTIF').hexdigest(),))
        db.execute("INSERT INTO ressources(genre,id,donnees) VALUES ('PATIENTS','P123456789A',%s)", (Jsonb({'ID':'P123456789A','Nom':'FICTIF'}),))
        db.execute("INSERT INTO consultations(id,rdv_id,patient_id,etat,donnees) VALUES ('consult-fictive','rdv-fictif','P123456789A','publie',%s)", (Jsonb({'CheminBrouillon':old+r'\Patients\fictif.docx'}),))
        db.execute("INSERT INTO publications(id,consultation_id,etat,donnees) VALUES ('publication-fictive','consult-fictive','a_traiter',%s)", (Jsonb(publication),))
        db.execute("INSERT INTO commandes(compte,id,empreinte,resultat) VALUES ('compte-fictif','commande-fictive','empreinte-conservee',%s)", (Jsonb({'items':[publication]}),))
    # Une sauvegarde U0 reste inspectable avec le moteur U1.
    legacy = r.backup(source, backups, config, old)
    legacy_manifest = r.inspect_bundle(backups/legacy['sauvegarde'])
    assert legacy_manifest['schema'] == 1
    assert legacy_manifest['revision'] == r.SERVICE_REVISION == '2026.09.16-u2b'
    class MigrationService:
        @contextmanager
        def connexion(self):
            with r.connection() as db, db.transaction():
                yield db
    acte = {'ID': 'CS', 'Code': 'CS', 'LibelleCourt': 'Consultation FICTIVE',
            'Tarif': '30.00', 'Depassement': '0.00'}
    with r.connection() as db:
        db.execute("INSERT INTO ressources(genre,id,donnees) VALUES ('ACTES','CS',%s)", (Jsonb(acte),))
    service = MigrationService()
    simulation = migrer(service)
    assert simulation['modifications'] == 1 and not simulation['conflits']
    assert migrer(service, simulation['empreinte'])['applique']
    avant = [{'CodeActe': 'CS', 'Montant': '30.00', 'Paye': 'N'}]
    apres = [{'CodeActe': 'CS', 'Montant': '30.00', 'Paye': 'O', 'ModePaiement': 'CB'}]
    with r.connection() as db:
        db.execute("INSERT INTO seances(id,patient_id,empreinte,lignes,impression_etat,impression_tentative) "
                   "VALUES ('consult-fictive','P123456789A','empreinte-fictive',%s,'inconnue','tentative-fictive')", (Jsonb(apres),))
        db.execute("INSERT INTO reglements_audit(seance_id,compte,avant,apres) VALUES "
                   "('consult-fictive','compte-fictif',%s,%s)", (Jsonb(avant), Jsonb(apres)))
    report = r.backup(source, backups, config, old)
    bundle = backups/report['sauvegarde']
    manifest = r.inspect_bundle(bundle)
    assert manifest['schema'] == 2
    assert manifest['revision'] == r.SERVICE_REVISION == '2026.09.16-u2b'
    source_db = os.environ['PGDATABASE']
    test_db = 'u1_restore_'+uuid.uuid4().hex
    with r.connection() as db:
        db.execute(sql.SQL('CREATE DATABASE {}').format(sql.Identifier(test_db)))
    try:
        os.environ['PGDATABASE'] = test_db
        target = Path('/target'); target.mkdir()
        recovered_config = Path('/config-target'); recovered_config.mkdir()
        result = r.restore(bundle, target, recovered_config, new, 'postgres')
        assert result['publications']==1 and result['archives']==2 and result['brouillons']==1
        assert r.regular_tree(target)==r.regular_tree(source)
        assert r.regular_tree(recovered_config)==r.regular_tree(config)
        with r.connection() as db:
            row = db.execute('SELECT donnees FROM publications').fetchone()['donnees']
            assert row['CheminDocx'].startswith(new) and row['sha_docx']==publication['sha_docx']
            command = db.execute('SELECT resultat,empreinte FROM commandes').fetchone()
            assert command['resultat']['items'][0]['CheminPdf'].startswith(new)
            assert command['empreinte']=='empreinte-conservee'
            assert db.execute("SELECT id FROM ressources WHERE genre='PATIENTS'").fetchone()['id']=='P123456789A'
            assert db.execute('SELECT max(version) AS v FROM schema_version').fetchone()['v'] == 2
            acte_restaure = db.execute("SELECT donnees,revision FROM ressources WHERE genre='ACTES' AND id='CS'").fetchone()
            assert acte_restaure['revision'] == 2
            assert acte_restaure['donnees']['Libelle'] == 'Consultation FICTIVE'
            assert acte_restaure['donnees']['Depassement'] == '0.00'
            migration = db.execute('SELECT avant,apres FROM migrations_ressources').fetchone()
            assert migration['avant'] == acte and migration['apres'] == acte_restaure['donnees']
            audit = db.execute('SELECT compte,avant,apres FROM reglements_audit').fetchone()
            assert audit == {'compte': 'compte-fictif', 'avant': avant, 'apres': apres}
            seance = db.execute('SELECT lignes,impression_etat,impression_tentative FROM seances').fetchone()
            assert seance == {'lignes': apres, 'impression_etat': 'inconnue', 'impression_tentative': 'tentative-fictive'}
        # Refus d'une base non vide, meme avec des nouveaux dossiers vides.
        t2=Path('/target2'); t2.mkdir(); c2=Path('/config-target2'); c2.mkdir()
        try:
            r.restore(bundle, t2, c2, new, 'postgres')
        except ValueError as exc:
            assert 'Base cible non vide' in str(exc)
        else:
            raise AssertionError('Base non vide acceptee')
        assert not any(t2.iterdir()) and not any(c2.iterdir())
        # Un composant altere doit echouer avant restauration.
        with (bundle/'base.dump').open('ab') as f: f.write(b'ALTERATION')
        try:
            r.inspect_bundle(bundle)
        except ValueError as exc:
            assert 'altere' in str(exc)
        else:
            raise AssertionError('Dump altere accepte')
        print(json.dumps({'postgresql_native':True, 'schema_u1_restaure':True,
                          'historique_migration_reglement':True, 'sauvegarde_u0_lisible':True, 'restauration_complete':True,
                          'changement_partage':True, 'base_non_vide_refusee':True,
                          'dump_altere_refuse':True, 'recette_office_smb_executee':False}))
    finally:
        os.environ['PGDATABASE'] = source_db
        with r.connection() as db:
            db.execute(sql.SQL('DROP DATABASE {}').format(sql.Identifier(test_db)))


if __name__ == '__main__': main()
