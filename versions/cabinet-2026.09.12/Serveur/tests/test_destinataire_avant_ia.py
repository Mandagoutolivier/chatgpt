"""Contrats de cablage VBA : les deux chemins IA appliquent le meme garde-fou.

Les cas metier du validateur sont executes par la recette Word U2. Ces tests
statiques ne pretendent pas remplacer la compilation ni l'execution Office.
"""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
OPENAI = (ROOT / 'Src/Prod6/modOpenAI_v22_corrige.bas').read_text(encoding='utf-8-sig')


def procedure(name):
    match = re.search(
        rf'(?ms)^(?:Public|Private) (?:Function|Sub) {name}\b.*?^End (?:Function|Sub)$',
        OPENAI,
    )
    assert match is not None, name
    return match.group(0)


def test_les_deux_chemins_revalident_juste_avant_l_envoi():
    for name in ('AppelerOpenAI', 'AppelerOpenAIStructure'):
        body = procedure(name)
        initial = 'Set cor = ResoudreDestinataireAvantIA(docSource)'
        final = 'Set cor = ResoudreDestinataireAvantIA(docSource, ident)'
        assert body.count('ResoudreDestinataireAvantIA(') == 2
        assert body.index(initial) < body.index('modAnonymise.AjouterCorrespondant ctx, cor, "DEST"')
        assert body.index('PreparerRequeteSortante(') < body.index(final)
        lines = [line.strip() for line in body.splitlines() if line.strip() and not line.lstrip().startswith("'")]
        sends = [index for index, line in enumerate(lines) if 'AppelerOpenAIRaw(' in line]
        assert len(sends) == 1
        assert lines[sends[0] - 1] == final
        assert 'modBase.CorrespondantParID(ident)' not in body


def test_resolution_exige_id_explicite_et_aucun_repli():
    body = procedure('ResoudreDestinataireAvantIA')
    rpc = 'modServiceNas.Appeler("correspondent.resolve", resolution)'
    assert body.index('If Len(ident) = 0 Then Err.Raise') < body.index(rpc)
    assert body.index('If Len(idAttendu) > 0 And ident <> idAttendu Then Err.Raise') < body.index(rpc)
    assert 'resolution("id") = ident' in body
    assert 'resolution("cle")' not in body
    assert 'resolution("examen")' not in body
    assert body.index(rpc) < body.index('VerifierDestinataireAvantIA ident, cor')


def test_validateur_refuse_selection_ou_statuts_incomplets():
    body = procedure('VerifierDestinataireAvantIA')
    for guard in (
        'If Len(ident) = 0 Then Err.Raise',
        'If cor Is Nothing Then Err.Raise',
        'If Not cor.Exists("ID") Then Err.Raise',
        'If CStr(cor("ID")) <> ident Then Err.Raise',
        'If Not cor.Exists("Actif") Or Not cor.Exists("AValider") Then Err.Raise',
        'If CStr(cor("Actif")) <> "1" Or CStr(cor("AValider")) <> "0" Then Err.Raise',
    ):
        assert guard in body
    assert 'modServiceNas.Appeler(' not in body


def test_chemin_cycle_et_recette_exercent_le_contrat():
    cycle = (ROOT / 'Src/Integration/modCycleCourrier.bas').read_text(encoding='utf-8-sig')
    assert cycle.count('modOpenAI_v22_corrige.AppelerOpenAIStructure(') == 2
    recette = (ROOT / 'Tests/Vba/word/modRecetteU2.bas').read_text(encoding='utf-8-sig')
    assert 'Private Const NOMBRE_ATTENDU_U2 As Long = 36' in recette
    assert 'modOpenAI_v22_corrige.VerifierDestinataireAvantIA ident, cor' in recette
    assert '    TesterDestinataireAvantIA\n' in recette
    assert 'DestinataireIARefuse(" ", cor)' in recette
    assert 'cor("Actif") = "0"' in recette
    assert 'cor("AValider") = "1"' in recette
    assert 'cor("AValider") = ""' in recette
    assert 'cor("AValider") = "inconnu"' in recette
    assert 'cor("Actif") = ""' in recette
    assert 'cor("Actif") = "inconnu"' in recette
