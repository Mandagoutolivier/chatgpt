"""Verifier les octets du lanceur contre les objets Git de son commit source."""
import re
from generer_lanceur import ROOT, HEADER, payload

def verifier():
    target = ROOT / 'Installateur'
    actual = (target / 'Demarrer_Installation_Cabinet.ps1').read_bytes()
    match = re.search(rb"(?m)^\$Commit='([a-f0-9]{40})'\r?$", actual)
    if not match:
        raise ValueError('Commit source du lanceur absent')
    code = payload(match.group(1).decode('ascii'))
    def windows(text):
        return text.replace('\r\n', '\n').replace('\n', '\r\n').encode('ascii')
    if actual != windows(code):
        raise ValueError('Payload PowerShell different du commit source')
    if (target / 'Demarrer_Installation_Cabinet.cmd').read_bytes() != windows(HEADER + code):
        raise ValueError('Lanceur CMD different du commit source')
    print('Lanceur conforme aux objets Git :', match.group(1).decode('ascii'))

if __name__ == '__main__':
    verifier()
