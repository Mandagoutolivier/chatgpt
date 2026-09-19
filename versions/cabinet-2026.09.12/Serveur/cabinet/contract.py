"""Validation complete avant transaction et avant tout effet sur les fichiers."""
from .domain import Refus

# Obligatoires / facultatifs. Les objets de fiches n'acceptent que du texte.
SPECS = {
    'whoami': ({}, {}), 'attentes': ({}, {}), 'reprises': ({}, {}),
    'publications': ({}, {}), 'stale_arrivals': ({}, {}),
    'command.result': ({'id': str}, {}),
    'record.get': ({'genre': str, 'id': str}, {}),
    'table.read': ({'genre': str}, {'q': str, 'year': str, 'date': str, 'offset': int, 'limit': int}),
    'dictionary.read': ({'genre': str}, {'offset': int, 'limit': int}),
    'journal.read': ({}, {'id': str, 'year': str, 'date': str, 'offset': int, 'limit': int}),
    'table.add': ({'genre': str, 'data': dict}, {}),
    'table.update': ({'genre': str, 'data': dict}, {}),
    'correspondent.save': ({'data': dict}, {}),
    'correspondent.resolve': ({}, {'id': str, 'cle': str, 'examen': str}),
    'dictionary.add': ({'genre': str, 'texte': str}, {}),
    'clinical.compare': ({'source': str, 'resultat': str}, {}),
    'nir.validate': ({'nir': str}, {}),
    'draft': ({'id': str, 'path': str}, {}),
    'agenda.status': ({'id': str, 'statut': str, 'revision': str}, {}),
    'bill': ({'id': str, 'publication_id': str, 'lignes': list}, {}),
    'print.request': ({'id': str, 'reimpression_confirmee': bool}, {}),
    'printed': ({'id': str, 'tentative': str, 'confirmee': bool}, {}),
    'payment': ({'id': str, 'date': str, 'mode': str, 'empreinte': str}, {'payeur': str, 'montant': str}),
    'publish': ({k: str for k in ('ConsultationID', 'PublicationID', 'PatientID',
        'DestinataireID', 'DateActe', 'CheminDocx', 'CheminPdf',
        'Patient_Nom', 'Patient_Prenom', 'Patient_DDN', 'Patient_Sexe')} | {'Relu': bool},
        {k: str for k in ('Nom', 'Prenom', 'DDN', 'NIR', 'TypeCourrier', 'SeanceID',
                          'RdvID', 'AnneeAgenda', 'DateValidation', 'Poste')}),
}
for op in ('arrive', 'cancel_arrival', 'claim', 'release', 'ack', 'billing.get'):
    SPECS[op] = ({'id': str}, {})

LINE_REQUIRED = {'SeanceID', 'PatientID', 'Date', 'CodeActe', 'Montant', 'TiersPayant', 'Paye'}
LINE_OPTIONAL = {'Nom', 'Prenom', 'DDN', 'NIR', 'AssureNom', 'AssurePrenom', 'AssureDDN',
                 'AssureNIR', 'CodeCerfa', 'ModePaiement', 'DateEncaissement', 'FeuilleSoinsImprimee', 'Notes'}


def validate(operation, params):
    if operation not in SPECS:
        raise Refus('Operation inconnue.', 404)
    required, optional = SPECS[operation]
    if type(params) is not dict or set(required) - params.keys() or params.keys() - (required.keys() | optional.keys()):
        raise Refus('Parametres invalides ou incomplets.', 422)
    for key, value in params.items():
        if type(value) is not (required | optional)[key]:
            raise Refus('Type invalide pour ' + key + '.', 422)
        if isinstance(value, str) and (len(value) > 1_000_000 or '\x00' in value):
            raise Refus('Texte invalide pour ' + key + '.', 422)
        if key in required and isinstance(value, str) and not value.strip():
            raise Refus('Champ obligatoire : ' + key + '.', 422)
    if 'data' in params and (any(type(v) is not str for v in params['data'].values())
                             or any(type(k) is not str for k in params['data'])):
        raise Refus('Les champs doivent etre du texte.', 422)
    for key in ('offset', 'limit'):
        if key in params and not (0 if key == 'offset' else 1) <= params[key] <= (1_000_000 if key == 'offset' else 200):
            raise Refus('Pagination invalide.', 422)
    if operation == 'bill':
        if not 1 <= len(params['lignes']) <= 4:
            raise Refus('Une feuille comporte de une a quatre lignes. Aucune comptabilisation.', 422)
        for line in params['lignes']:
            if (type(line) is not dict or LINE_REQUIRED - line.keys()
                    or line.keys() - (LINE_REQUIRED | LINE_OPTIONAL)
                    or any(type(v) is not str for v in line.values())):
                raise Refus('Ligne comptable invalide ou incomplete.', 422)
    if operation == 'publish' and params['DestinataireID'].strip().upper() == 'A_COMPLETER':
        raise Refus('Destinataire a completer avant publication.', 422)
