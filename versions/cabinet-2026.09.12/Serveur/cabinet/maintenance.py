"""Maintenance locale explicite ; aucune tache automatique, aucun acces HTTP admin."""
from datetime import datetime, timezone
from pathlib import Path
import hashlib
import os
import re
import secrets
from psycopg.types.json import Jsonb
from .domain import Refus, empreinte

RECU = '_cabinet_resultat_archive'
VERROU = 20260912


def resultat_rejouable(resultat):
    if isinstance(resultat, dict) and resultat.get(RECU) == 1:
        raise Refus('Commande deja executee ; resultat ancien archive. Rechargez les donnees et faites verifier la reprise, sans recreer automatiquement cette commande.',
                    409, 'resultat_archive')
    return resultat


def compacter_commandes(service, avant, appliquer=None, limite=100):
    """Retire uniquement les copies de resultats, conserve la preuve de deduplication.

    La duree de retention appartient a l'exploitant : aucun delai arbitraire par defaut.
    Une simulation ne verrouille ni ne modifie les donnees. L'application recalcule son
    empreinte sous le meme verrou transactionnel que les mutations metier.
    """
    if not isinstance(avant, datetime) or avant.tzinfo is None or avant.utcoffset() is None:
        raise Refus('Date limite avec fuseau obligatoire.', 422)
    if avant >= datetime.now(timezone.utc):
        raise Refus('La date limite doit etre passee.', 422)
    if type(limite) is not int or not 1 <= limite <= 500:
        raise Refus('Lot attendu entre 1 et 500 commandes.', 422)
    with service.connexion() as db:
        if appliquer is not None:
            db.execute('SELECT pg_advisory_xact_lock(%s)', (VERROU,))
        rows = db.execute('SELECT compte,id,empreinte,resultat,cree_le FROM commandes '
                          'WHERE cree_le < %s AND NOT (resultat ? %s) '
                          'ORDER BY cree_le,compte,id LIMIT %s', (avant, RECU, limite)).fetchall()
        preuves = [empreinte({'compte':r['compte'], 'id':r['id'], 'commande':r['empreinte'],
                              'resultat':r['resultat'], 'date':r['cree_le'].isoformat()}) for r in rows]
        digest = empreinte({'version':1, 'avant':avant.astimezone(timezone.utc).isoformat(),
                            'limite':limite, 'preuves':preuves})
        if appliquer is not None:
            if not secrets.compare_digest(appliquer, digest):
                raise Refus('Simulation perimee ou empreinte incorrecte ; aucune modification.')
            for row in rows:
                recu = {RECU:1, 'empreinte_resultat':empreinte(row['resultat'])}
                db.execute('UPDATE commandes SET resultat=%s WHERE compte=%s AND id=%s',
                           (Jsonb(recu), row['compte'], row['id']))
            if rows:
                db.execute('INSERT INTO evenements(compte,operation,objet) VALUES (%s,%s,%s)',
                           ('maintenance', 'commandes.compaction', digest))
        return {'version':1, 'commandes':len(rows), 'empreinte':digest,
                'avant':avant.astimezone(timezone.utc).isoformat(), 'limite':limite,
                'applique':appliquer is not None, 'preuves_deduplication_conservees':True}


def etat_exploitation(service):
    """Comptages uniquement : pas de nom, token, dossier ni texte clinique."""
    with service.connexion() as db:
        commandes = db.execute('SELECT count(*) AS total, count(*) FILTER (WHERE resultat ? %s) AS archives FROM commandes', (RECU,)).fetchone()
        return {'commandes':commandes, 'comptes_actifs':db.execute('SELECT count(*) AS n FROM comptes WHERE actif').fetchone()['n'],
                'schema':db.execute('SELECT max(version) AS n FROM schema_version').fetchone()['n']}


def etat_compte(service, identifiant):
    with service.connexion() as db:
        row = db.execute('SELECT identifiant,roles,actif,token_sha256 FROM comptes WHERE identifiant=%s', (identifiant,)).fetchone()
        if not row:
            raise Refus('Compte introuvable.', 404)
        return {'identifiant':row['identifiant'], 'roles':row['roles'], 'actif':row['actif'],
                'empreinte_rotation':empreinte(row)}


def renouveler_jeton(service, identifiant, attendu, sortie):
    """Le nouveau jeton est ecrit dans un fichier neuf 0600 avant le commit.

    L'ancien reste valide si la creation du fichier ou la transaction echoue.
    Aucun secret n'est renvoye dans le resultat ni inscrit dans le journal.
    """
    if not re.fullmatch('[a-f0-9]{64}', attendu):
        raise Refus('Empreinte du compte obligatoire.', 422)
    path = Path(sortie)
    cree = False
    try:
        with service.connexion() as db:
            db.execute('SELECT pg_advisory_xact_lock(%s)', (VERROU,))
            row = db.execute('SELECT identifiant,roles,actif,token_sha256 FROM comptes WHERE identifiant=%s FOR UPDATE', (identifiant,)).fetchone()
            if not row or not row['actif']:
                raise Refus('Compte absent ou revoque ; aucune reactivation automatique.', 409)
            if not secrets.compare_digest(attendu, empreinte(row)):
                raise Refus('Compte modifie depuis la verification ; relire son etat.')
            token = secrets.token_urlsafe(48)
            fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, 'O_NOFOLLOW', 0), 0o600)
            cree = True
            with os.fdopen(fd, 'w', encoding='utf-8', newline='\n') as output:
                output.write(token+'\n'); output.flush(); os.fsync(output.fileno())
            db.execute('UPDATE comptes SET token_sha256=%s WHERE identifiant=%s',
                       (hashlib.sha256(token.encode()).hexdigest(), identifiant))
            db.execute('INSERT INTO evenements(compte,operation,objet) VALUES (%s,%s,%s)',
                       ('maintenance', 'compte.rotation', identifiant))
        return {'renouvele':True, 'ancien_jeton_invalide':True}
    except BaseException:
        if cree:
            path.unlink(missing_ok=True)
        raise


def revoquer_compte(service, identifiant):
    with service.connexion() as db:
        db.execute('SELECT pg_advisory_xact_lock(%s)', (VERROU,))
        row = db.execute('SELECT actif FROM comptes WHERE identifiant=%s FOR UPDATE', (identifiant,)).fetchone()
        if not row:
            raise Refus('Compte introuvable.', 404)
        if row['actif']:
            db.execute('UPDATE comptes SET actif=false WHERE identifiant=%s', (identifiant,))
            db.execute('INSERT INTO evenements(compte,operation,objet) VALUES (%s,%s,%s)',
                       ('maintenance', 'compte.revocation', identifiant))
        return {'revoque':True, 'deja_revoque':not row['actif']}
