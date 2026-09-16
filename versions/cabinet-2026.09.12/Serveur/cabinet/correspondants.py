"""Indicateurs canoniques des correspondants, y compris les fiches U0.

Les anciennes fiches omettaient AValider/ParDefaut : vide signifiait faux,
et Actif vide signifiait actif. Cette compatibilite est normalisee sur le
serveur ; les clients ne doivent jamais accepter un statut incomplet.
"""
from .domain import Refus, plier


def normaliser_indicateurs(data, historique=False):
    d = dict(data)
    for champ, defaut in (('Actif', '1'), ('AValider', '0'), ('ParDefaut', '0')):
        valeur = d.get(champ, '')
        if not isinstance(valeur, str):
            raise Refus('Indicateur correspondant invalide : ' + champ, 422)
        if historique:
            valeur = plier(valeur)
            if valeur in {'oui', 'o'}: valeur = '1'
            elif valeur in {'non', 'n'}: valeur = '0'
        if valeur not in {'', '0', '1'}:
            raise Refus('Indicateur correspondant invalide : ' + champ, 422)
        d[champ] = valeur or defaut
    return d
