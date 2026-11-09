"""Porównanie dwóch wersji danych według kolumny-klucza."""

from __future__ import annotations

import datetime as dt
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Optional, Union

from .model import Value

Comparable = Union[None, str, int, float, bool]


class DuplicateKeyError(ValueError):
    """Klucz nie wskazuje obiektów jednoznacznie (powtórzenia lub braki)."""

    def __init__(self, key: str, bad_rows: int) -> None:
        super().__init__(key, bad_rows)
        self.key = key
        self.bad_rows = bad_rows


@dataclass(frozen=True)
class Change:
    kind: str  # "added" | "removed" | "modified"
    key: str
    column: Optional[str] = None
    old: Value = None
    new: Value = None


def normalize(value: Value) -> Comparable:
    """Sprowadza wartość do postaci, która przeżywa zapis i odczyt z XLSX.

    Excel trzyma liczby jako float, a daty jako datetime z dokładnością
    do sekundy; pusty tekst i brak wartości są dla odbiorcy tym samym.
    """
    if value is None or isinstance(value, (bool, int)):
        return value  # int zostaje int: float zlałby identyfikatory powyżej 2**53
    if isinstance(value, float):
        return round(float(value), 9)
    if isinstance(value, dt.datetime):
        return value.replace(microsecond=0, tzinfo=None).isoformat()
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day).isoformat()
    if isinstance(value, dt.time):
        return value.replace(microsecond=0).isoformat()
    text = str(value)
    return text if text.strip() else None


def key_text(value: Value) -> Optional[str]:
    norm = normalize(value)
    if norm is None:
        return None
    if isinstance(norm, float) and norm.is_integer():
        return str(int(norm))
    return str(norm)


def index_rows(columns: Sequence[str], rows: Sequence[Sequence[Value]], key: str) -> dict[str, dict[str, Value]]:
    """Wiersze jako słownik {klucz: {kolumna: wartość}}; klucz musi być unikalny."""
    position = list(columns).index(key)
    indexed: dict[str, dict[str, Value]] = {}
    bad = 0
    for row in rows:
        k = key_text(row[position])
        if k is None or k in indexed:
            bad += 1
            continue
        indexed[k] = dict(zip(columns, row))
    if bad:
        raise DuplicateKeyError(key, bad)
    return indexed


def compare(old: dict[str, dict[str, Value]], new: dict[str, dict[str, Value]], columns: Sequence[str]) -> list[Change]:
    """Zmiany od `old` do `new`, tylko w kolumnach obecnych w obu wersjach.

    Kolejność: zmienione i dodane w kolejności nowych danych, potem usunięte.
    """
    changes: list[Change] = []
    for k, row in new.items():
        before = old.get(k)
        if before is None:
            changes.append(Change("added", k))
            continue
        for col in columns:
            if normalize(before.get(col)) != normalize(row.get(col)):
                changes.append(Change("modified", k, col, before.get(col), row.get(col)))
    changes.extend(Change("removed", k) for k in old if k not in new)
    return changes
