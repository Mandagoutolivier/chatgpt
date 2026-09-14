"""Contrat ACTES compatible avec les classeurs et imports anterieurs."""
from .domain import Refus, montant

FIELDS = ['Code', 'Libelle', 'LibelleCourt', 'Tarif', 'CodeAssocie', 'TarifAssocie',
          'LibelleCerfa', 'Actif', 'Notes', 'Depassement']


def normaliser(data, ident=None, prior=None):
    d = dict(data)
    if set(d) - set(FIELDS) - {'ID', '_revision'}:
        raise Refus('Champ ACTES inconnu : migration manuelle requise.', 422)
    if any(type(v) is not str for v in d.values()):
        raise Refus('Les champs ACTES doivent etre du texte.', 422)
    code = d.get('Code') or d.get('ID') or ident
    if not code or (ident and code != ident) or (d.get('ID') and d['ID'] != code):
        raise Refus('ID et Code ACTES doivent etre identiques et immuables.', 422)
    a, b = d.get('Libelle', ''), d.get('LibelleCourt', '')
    if a and b and a != b:
        # Une edition d'un seul alias d'une fiche canonique est non ambigue.
        old = (prior or {}).get('Libelle', (prior or {}).get('LibelleCourt', ''))
        if prior and a == old: a = b
        elif prior and b == old: b = a
        else: raise Refus('Libelles ACTES en conflit : migration manuelle requise.', 422)
    label = a or b or code
    d = dict.fromkeys(FIELDS, '') | d | {'ID': code, 'Code': code, 'Libelle': label, 'LibelleCourt': label}
    montant(d['Tarif'])
    if d['CodeAssocie']: montant(d['TarifAssocie'])
    if d['Depassement']: montant(d['Depassement'])
    # Conservation exacte du depassement historique ; ne jamais l'ajouter
    # implicitement au tarif (sens non documente dans le classeur source).
    if d['Actif'] not in {'', '0', '1'}: raise Refus('Actif ACTES invalide.', 422)
    d['Actif'] = d['Actif'] or '1'
    return d
