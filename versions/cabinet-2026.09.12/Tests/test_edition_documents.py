"""Tests des regles XML des annexes, sans SQL ni appels reseau."""
import io
import sys
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'Serveur'))
from cabinet.editing import reperes, verifier_revision, W
from cabinet.domain import Refus
from cabinet.contract import validate

V={'PatientID':'FICTIF-P1','ConsultationID':'FICTIF-S1','Patient_Nom':'FICTIF',
   'Patient_Prenom':'ESSAI','Patient_DDN':'01/01/1980','Patient_Sexe':'M',
   'AnnexeDestinataire_001':'FICTIF-DEST-A'}


def docx(dest='Docteur ANNEXE',principal='Docteur PRINCIPAL',changes=None,annexe=True):
    v=dict(V);v.update(changes or {})
    settings=ET.Element(W+'settings');vs=ET.SubElement(settings,W+'docVars')
    for k,x in v.items():ET.SubElement(vs,W+'docVar',{W+'name':k,W+'val':x})
    doc=ET.Element(W+'document');body=ET.SubElement(doc,W+'body')
    for number,name,text in [('1','DESTINATAIRE',principal)]+([('2','U2ANN_DEST_001',dest)] if annexe else []):
        p=ET.SubElement(body,W+'p');ET.SubElement(p,W+'bookmarkStart',{W+'id':number,W+'name':name})
        ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text=text
        ET.SubElement(p,W+'bookmarkEnd',{W+'id':number})
    b=io.BytesIO()
    with zipfile.ZipFile(b,'w') as z:
        z.writestr('word/settings.xml',ET.tostring(settings));z.writestr('word/document.xml',ET.tostring(doc))
    return b.getvalue()


class CorrespondantsFictifs:
    def _record(self,db,genre,id):
        return {'ID':id,'Actif':'0' if id=='INACTIF' else '1','AValider':'0',
                'BlocDestinataire':'Docteur ANNEXE' if id=='FICTIF-DEST-A' else 'Docteur REMPLACEMENT\nADRESSE FICTIVE'}


class AnnexesXML(unittest.TestCase):
    def check(self,new):
        return verifier_revision(docx(),new,V,CorrespondantsFictifs(),None)
    def test_revision_inchangee(self):
        self.assertEqual(self.check(docx()),[{'Numero':'001','DestinataireID':'FICTIF-DEST-A'}])
    def test_nouveau_destinataire(self):
        new=docx('Docteur REMPLACEMENT\nADRESSE FICTIVE',changes={'AnnexeDestinataire_001':'FICTIF-DEST-B'})
        self.assertEqual(self.check(new)[0]['DestinataireID'],'FICTIF-DEST-B')
    def test_identite_interdite(self):
        with self.assertRaises(Refus):self.check(docx(changes={'PatientID':'AUTRE'}))
    def test_destinataire_principal_preserve(self):
        with self.assertRaises(Refus):self.check(docx(principal='AUTRE'))
    def test_repere_annexe_obligatoire(self):
        with self.assertRaises(Refus):self.check(docx(annexe=False))
    def test_destinataire_vide_refuse(self):
        with self.assertRaises(Refus):self.check(docx(''))
    def test_destinataire_inactif_refuse(self):
        with self.assertRaises(Refus):self.check(docx(changes={'AnnexeDestinataire_001':'INACTIF'}))
    def test_texte_adresse_non_concordant_refuse(self):
        with self.assertRaises(Refus):self.check(docx('ADRESSE NON CONCORDANTE'))
    def test_selection_type_strict(self):
        validate('claim',{'id':'FICTIF-S1','selection':'empreinte-fictive'})
        with self.assertRaises(Refus):validate('claim',{'id':'FICTIF-S1','selection':False})
    def test_contrat_revision(self):
        p={k:'FICTIF' for k in ('source_id','source_sha','revision_id','patient_id','docx','pdf','sha_docx','sha_pdf')}
        validate('publication.revise',p)
        p.pop('source_sha')
        with self.assertRaises(Refus):validate('publication.revise',p)


if __name__=='__main__':unittest.main()
