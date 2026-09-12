"""Regles metier independantes d'Office, HTTP et PostgreSQL."""
from __future__ import annotations
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
import hashlib
import json
import re
import unicodedata

class Refus(Exception):
    def __init__(self, message: str, status: int = 409):
        super().__init__(message)
        self.status = status


def date_fr(value: str) -> date:
    if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", value):
        raise Refus("Date attendue au format JJ/MM/AAAA.", 422)
    try:
        return datetime.strptime(value, "%d/%m/%Y").date()
    except ValueError as exc:
        raise Refus("Date inexistante.", 422) from exc


def plier(value: str) -> str:
    return " ".join("".join(c for c in unicodedata.normalize("NFD", value.casefold())
                           if not unicodedata.combining(c)).split())


def montant(value: str) -> Decimal:
    if not re.fullmatch(r"\d{1,7}([.,]\d{1,2})?", value):
        raise Refus("Montant invalide : deux decimales au maximum.", 422)
    return Decimal(value.replace(",", ".")).quantize(Decimal("0.01"))


def patient_valide(data: dict, today: date) -> None:
    for name in ("Nom", "Prenom", "DDN", "Sexe"):
        if not str(data.get(name, "")).strip():
            raise Refus("Fiche incomplete : " + name, 422)
    if date_fr(data["DDN"]) > today:
        raise Refus("Date de naissance future.", 422)
    if data["Sexe"] not in ("M", "F"):
        raise Refus("Sexe attendu M ou F ; aucune deduction automatique.", 422)
    for name in ("AssureDDN",):
        if data.get(name): date_fr(data[name])
    for name in ("NIR", "AssureNIR"):
        if data.get(name): valider_nir(data[name])


def valider_nir(value: str) -> str:
    nir = re.sub(r"[\s.-]", "", value).upper()
    if not re.fullmatch(r"[12]\d{4}(?:\d{2}|2[AB])\d{6}\d{2}", nir):
        raise Refus("NIR incomplet ou format non pris en charge : verifier manuellement.", 422)
    numeric = nir[:13].replace("2A", "19").replace("2B", "18")
    if 97 - int(numeric) % 97 != int(nir[13:]):
        raise Refus("Cle de controle du NIR incorrecte.", 422)
    return nir


def creneau(data: dict) -> tuple[date, int, int]:
    day = date_fr(data["Date"])
    try:
        if not re.fullmatch(r"\d{2}:\d{2}", data["Heure"]): raise ValueError()
        hour, minute = map(int, data["Heure"].split(":"))
        duration = int(data["DureeMin"])
        start = hour * 60 + minute
        if not (0 <= hour <= 23 and 0 <= minute <= 59 and 1 <= duration <= 480
                and start + duration <= 1440): raise ValueError()
    except (ValueError, KeyError) as exc:
        raise Refus("Horaire ou duree du rendez-vous invalide.", 422) from exc
    return day, start, start + duration


def chevauche(a: dict, b: dict) -> bool:
    da, sa, ea = creneau(a)
    db, sb, eb = creneau(b)
    return a.get("Statut") != "Annule" and b.get("Statut") != "Annule" and da == db and sa < eb and sb < ea


def empreinte(data: object) -> str:
    return hashlib.sha256(json.dumps(data, sort_keys=True, ensure_ascii=False,
                                    separators=(",", ":")).encode()).hexdigest()


def comparer_clinique(source: str, resultat: str) -> dict:
    """Sentinelles de regression, jamais une certification medicale du texte."""
    patterns = {
        "nombres_et_unites": r"\b\d+(?:[.,]\d+)?\s*(?:mg|µg|mcg|g|mmhg|mm|cm|ml|min|bpm|%|ui)?",
        "negations": r"\b(?:ne|n['’]|pas|sans|absence|aucun|aucune|negatif|negative)\b",
        "identifiants_masques": r"\[\[[A-Z0-9_]+\]\]|\{\{[A-Z0-9_]+\}\}",
    }
    from collections import Counter
    differences = []
    for name, pattern in patterns.items():
        left = Counter(re.findall(pattern, plier(source))) if name != "identifiants_masques" else Counter(re.findall(pattern, source))
        right = Counter(re.findall(pattern, plier(resultat))) if name != "identifiants_masques" else Counter(re.findall(pattern, resultat))
        if left != right:
            differences.append({"controle": name, "retires": list((left-right).elements()),
                                "ajoutes": list((right-left).elements())})
    return {"differences": differences, "relecture_obligatoire": True}
