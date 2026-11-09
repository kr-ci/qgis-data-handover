"""Wykrywanie kolumn, które prawdopodobnie zawierają dane osobowe.

Dwie reguły: nazwa pola (lub nagłówek) zawiera znane słowo, albo większość
próbki wartości ma kształt danej osobowej (PESEL z poprawną sumą kontrolną,
e-mail, telefon, numer księgi wieczystej). Wykrycie to podpowiedź dla
autora eksportu — decyzję podejmuje człowiek.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Sequence
from dataclasses import dataclass

from .model import Value

# Słowa w nazwie pola. Porównujemy całe tokeny (fragmenty nazwy rozdzielone
# „_”, spacją lub cyfrą) albo ich początek — „tel” nie trafi w „hotel”.
_EXACT = {"tel", "kw", "mail", "nip", "owner", "osoba", "person", "surname"}
_PREFIX = (
    "imie", "imię", "imion", "nazwisk", "wlasciciel", "właściciel", "wlasc", "właść",
    "pesel", "email", "e-mail", "telefon", "phone", "mobile", "komork", "komórk",
    "adres", "address", "ulica", "street", "zamieszk", "dowod", "dowód",
    "paszport", "passport", "firstname", "lastname", "surname", "ksiega", "księga",
)

_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$")
_PHONE = re.compile(r"^(\+?48[\s-]?)?(\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2})$")
_KW = re.compile(r"^[A-Z]{2}\d[A-Z]/\d{8}/\d$")
_PESEL_WEIGHTS = (1, 3, 7, 9, 1, 3, 7, 9, 1, 3)


@dataclass(frozen=True)
class Finding:
    field: str
    reason: str


def is_pesel(text: str) -> bool:
    """Prawdziwy PESEL: 11 cyfr i zgodna cyfra kontrolna."""
    if len(text) != 11 or not text.isdigit():
        return False
    total = sum(int(d) * w for d, w in zip(text[:10], _PESEL_WEIGHTS))
    return (10 - total % 10) % 10 == int(text[10])


def _tokens(name: str) -> list[str]:
    spaced = re.sub(r"([a-ząćęłńóśźż])([A-ZĄĆĘŁŃÓŚŹŻ])", r"\1 \2", name)
    return [t for t in re.split(r"[\W_\d]+", spaced.lower()) if t]


def _name_reason(name: str) -> str | None:
    for token in _tokens(name):
        if token in _EXACT or token.startswith(_PREFIX):
            return token
    return None


def _value_reason(values: Iterable[Value]) -> str | None:
    texts = [str(v).strip() for v in values if v is not None and str(v).strip()]
    if len(texts) < 3:
        return None
    checks = (("PESEL", is_pesel), ("e-mail", _EMAIL.match), ("telefon", _PHONE.match), ("KW", _KW.match))
    for label, check in checks:
        hits = sum(1 for t in texts if check(t))
        if hits * 2 >= len(texts):
            return label
    return None


def detect(fields: Sequence[tuple[str, str]], samples: dict[str, list[Value]]) -> list[Finding]:
    """`fields` to pary (nazwa pola, nagłówek); `samples` — próbka wartości."""
    findings = []
    for name, header in fields:
        reason = _name_reason(name) or _name_reason(header)
        if reason is not None:
            findings.append(Finding(name, f"nazwa: {reason}"))
            continue
        value_reason = _value_reason(samples.get(name, []))
        if value_reason is not None:
            findings.append(Finding(name, f"wartości: {value_reason}"))
    return findings


def mentions(text: str, names: Iterable[str]) -> bool:
    """Czy tekst (np. filtr warstwy) odwołuje się do któregoś z pól — całym słowem, bez wielkości liter."""
    lowered = text.lower()
    return any(re.search(rf"(?<!\w){re.escape(n.lower())}(?!\w)", lowered) for n in names if n)


def mask(value: Value) -> str | None:
    """„Jan Kowalski” → „J*** K***”, cyfry i symbole → „***”.

    Maska ma stałą długość, więc nie zdradza długości oryginału.
    """
    if value is None:
        return None
    out = []
    for token in str(value).split():
        out.append(token[0] + "***" if token[0].isalpha() else "***")
    return " ".join(out) or None
