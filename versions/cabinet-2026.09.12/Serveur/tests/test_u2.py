from datetime import datetime, timedelta, timezone
import hashlib
import json
import logging
import os
import uuid
import pytest
from fastapi.testclient import TestClient
from psycopg.types.json import Jsonb
from cabinet.api import create_app
from cabinet.domain import Refus
from cabinet.maintenance import (compacter_commandes, etat_compte, etat_exploitation,
                                 renouveler_jeton, revoquer_compte, resultat_rejouable, RECU)
from test_service import service, rpc


def ancienne_commande(service):
    rid=uuid.uuid4().hex
    params={'genre':'CORRESPONDANTS','data':{'Nom':'IDENTITE FICTIVE U2'}}
    result=rpc(service,'table.add',params,rid=rid)
    with service.connexion() as db:
        db.execute("UPDATE commandes SET cree_le=now()-interval '120 days' WHERE id=%s",(rid,))
    return rid,params,result


def test_compaction_preserve_preuve_et_refuse_reexecution(service):
    rid,params,result=ancienne_commande(service)
    avant=datetime.now(timezone.utc)-timedelta(days=60)
    plan=compacter_commandes(service,avant)
    assert plan['commandes']==1 and not plan['applique']
    assert 'IDENTITE' not in json.dumps(plan)
    assert rpc(service,'command.result',{'id':rid})['resultat']==result
    with service.connexion() as db:
        before=db.execute('SELECT empreinte,cree_le FROM commandes WHERE id=%s',(rid,)).fetchone()
    assert compacter_commandes(service,avant,plan['empreinte'])['applique']
    for op,p in [('table.add',params),('command.result',{'id':rid})]:
        with pytest.raises(Refus) as error:rpc(service,op,p,rid=rid)
        assert error.value.code=='resultat_archive'
    with pytest.raises(Refus,match='contenu different'):
        rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'AUTRE FICTIF'}},rid=rid)
    assert rpc(service,'command.result',{'id':rid},'medecin')['trouve'] is False
    with service.connexion() as db:
        row=db.execute('SELECT empreinte,cree_le,resultat FROM commandes WHERE id=%s',(rid,)).fetchone()
        assert {k:row[k] for k in before}==before
        assert row['resultat'][RECU]==1 and 'IDENTITE' not in json.dumps(row['resultat'])
        assert db.execute("SELECT count(*) AS n FROM ressources WHERE genre='CORRESPONDANTS'").fetchone()['n']==1
    assert compacter_commandes(service,avant)['commandes']==0
    assert rpc(service,'record.get',{'genre':'CORRESPONDANTS','id':result['ID']})==result


def test_compaction_plan_perime_et_lots_bornes(service):
    rid,params,result=ancienne_commande(service)
    avant=datetime.now(timezone.utc)-timedelta(days=60)
    plan=compacter_commandes(service,avant)
    with service.connexion() as db:
        db.execute('UPDATE commandes SET resultat=%s WHERE id=%s',(Jsonb({'ok':'CHANGEMENT FICTIF'}),rid))
    with pytest.raises(Refus,match='perimee'):compacter_commandes(service,avant,plan['empreinte'])
    assert RECU not in rpc(service,'command.result',{'id':rid})['resultat']
    rpc(service,'table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'RECENT FICTIF'}})
    assert compacter_commandes(service,avant,limite=1)['commandes']==1
    for limit in [True,0,501]:
        with pytest.raises(Refus):compacter_commandes(service,avant,limite=limit)


@pytest.mark.parametrize('date',[datetime(2020,1,1),datetime.now(timezone.utc)+timedelta(days=1)])
def test_retention_exige_date_explicite_valide(date):
    class Interdit:
        def connexion(self):pytest.fail('Connexion avant validation')
    with pytest.raises(Refus):compacter_commandes(Interdit(),date)


def test_resultat_recent_inchange():
    for value in [None,True,{'ID':'FICTIF'},[1,2]]:
        assert resultat_rejouable(value)==value


def test_rotation_atomique_et_commandes_conservees(service,tmp_path):
    rid,_,result=ancienne_commande(service)
    old=('test-secretariat-')*6
    state=etat_compte(service,'secretariat');dest=tmp_path/'jeton-fictif.token'
    out=renouveler_jeton(service,'secretariat',state['empreinte_rotation'],dest)
    new=dest.read_text().strip()
    assert len(new)>=32 and new!=old and new not in json.dumps(out)
    assert service.compte(new)['roles']==['secretariat']
    with pytest.raises(Refus):service.compte(old)
    assert rpc(service,'command.result',{'id':rid})['resultat']==result
    assert not (dest.stat().st_mode & 0o077)
    with service.connexion() as db:
        events=db.execute("SELECT operation,objet FROM evenements WHERE operation='compte.rotation'").fetchall()
        assert events==[{'operation':'compte.rotation','objet':'secretariat'}]
    with pytest.raises(Refus):renouveler_jeton(service,'secretariat',state['empreinte_rotation'],tmp_path/'autre.token')
    assert not (tmp_path/'autre.token').exists()


def test_rotation_echec_fichier_conserve_ancien_jeton(service,tmp_path):
    old=('test-medecin-')*6;state=etat_compte(service,'medecin')
    dest=tmp_path/'existant.token';dest.write_text('FICTIF A CONSERVER')
    with pytest.raises(FileExistsError):renouveler_jeton(service,'medecin',state['empreinte_rotation'],dest)
    assert dest.read_text()=='FICTIF A CONSERVER' and service.compte(old)['identifiant']=='medecin'
    assert etat_compte(service,'medecin')==state


def test_rotation_echec_ecriture_annule_transaction(service,tmp_path,monkeypatch):
    state=etat_compte(service,'medecin');dest=tmp_path/'neuf.token'
    def fail(_):raise OSError('ECHEC FICTIF FSYNC')
    monkeypatch.setattr(os,'fsync',fail)
    with pytest.raises(OSError):renouveler_jeton(service,'medecin',state['empreinte_rotation'],dest)
    assert not dest.exists() and etat_compte(service,'medecin')==state


def test_revocation_refuse_compte_absent_et_acteur_deja_lu(service,tmp_path):
    actor=service.compte(('test-medecin-')*6)
    state=etat_compte(service,'medecin')
    assert revoquer_compte(service,'medecin')['deja_revoque'] is False
    assert revoquer_compte(service,'medecin')['deja_revoque'] is True
    with pytest.raises(Refus):service.compte(('test-medecin-')*6)
    with pytest.raises(Refus,match='revoque'):
        service.executer(actor,'dictionary.add',{'genre':'MEDICAMENTS','texte':'FICTIF'},uuid.uuid4().hex)
    with pytest.raises(Refus):renouveler_jeton(service,'medecin',state['empreinte_rotation'],tmp_path/'interdit.token')
    assert not (tmp_path/'interdit.token').exists()
    with pytest.raises(Refus,match='introuvable'):revoquer_compte(service,'absent')


def test_diagnostics_bornes_sans_donnees_ni_operation_arbitraire(service,caplog):
    caplog.set_level(logging.INFO,logger='cabinet')
    token=('test-secretariat-')*6
    with TestClient(create_app(service)) as client:
        for operation,params,status in [('table.add',{'genre':'CORRESPONDANTS','data':{'Nom':'DONNEE FICTIVE PRIVEE'}},200),
                                        ('DONNEE FICTIVE PRIVEE',{},404)]:
            response=client.post('/v1/rpc',headers={'Authorization':'Bearer '+token},json={'operation':operation,'params':params,'request_id':'REQUETE-FICTIVE-CONFIDENTIELLE'})
            assert response.status_code==status
    messages='\n'.join(r.getMessage() for r in caplog.records if r.name=='cabinet')
    assert 'duree_ms=' in messages and 'operation=inconnue' in messages and 'statut=200' in messages
    for secret in [token,'DONNEE FICTIVE PRIVEE','REQUETE-FICTIVE-CONFIDENTIELLE']:assert secret not in messages
    state=etat_exploitation(service)
    assert state['commandes']['total']==1 and state['commandes']['archives']==0
    assert 'DONNEE' not in json.dumps(state)
