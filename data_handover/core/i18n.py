"""Teksty wtyczki i pliku wynikowego w dziesięciu językach.

Teksty leżą w `core/lang/<kod>.json` — jeden plik na język, te same klucze
co w `en.json`. Jeden słownik obsługuje zarówno okno wtyczki, jak i treść
arkuszy, dzięki czemu nie potrzebujemy narzędzi Qt do kompilacji tłumaczeń
(lrelease). Brak tekstu w danym języku = tekst angielski.
"""

from __future__ import annotations

import json
import os

# Kolejność listy w oknie; angielski jest domyślny.
LANGUAGES = ("en", "de", "fr", "es", "pt", "pl", "ru", "ko", "ja", "zh")

# Nazwy języków w ich własnym zapisie — tak użytkownik szuka swojego języka na liście.
LANGUAGE_NAMES = {
    "en": "English",
    "de": "Deutsch",
    "fr": "Français",
    "es": "Español",
    "pt": "Português",
    "pl": "Polski",
    "ru": "Русский",
    "ko": "한국어",
    "ja": "日本語",
    "zh": "中文（简体）",
}

# Separator tysięcy i dziesiętny w tekstach (arkusz „Zmiany”, podgląd).
# Komórki z liczbami pokazują separatory według ustawień Excela odbiorcy.
NUMBER_SEPARATORS = {
    "en": (",", "."), "de": (".", ","), "fr": (" ", ","), "es": (".", ","), "pt": (".", ","),
    "pl": (" ", ","), "ru": (" ", ","), "ko": (",", "."), "ja": (",", "."), "zh": (",", "."),
}

# Format daty: (kod formatu Excela, wzorzec strftime).
DATE_FORMATS = {
    "en": ("YYYY-MM-DD", "%Y-%m-%d"), "de": ("DD.MM.YYYY", "%d.%m.%Y"), "fr": ("DD/MM/YYYY", "%d/%m/%Y"),
    "es": ("DD/MM/YYYY", "%d/%m/%Y"), "pt": ("DD/MM/YYYY", "%d/%m/%Y"), "pl": ("DD.MM.YYYY", "%d.%m.%Y"),
    "ru": ("DD.MM.YYYY", "%d.%m.%Y"), "ko": ("YYYY-MM-DD", "%Y-%m-%d"), "ja": ("YYYY-MM-DD", "%Y-%m-%d"),
    "zh": ("YYYY-MM-DD", "%Y-%m-%d"),
}

_DIR = os.path.join(os.path.dirname(__file__), "lang")


def _load(lang: str) -> dict[str, str]:
    path = os.path.join(_DIR, f"{lang}.json")
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        data: dict[str, str] = json.load(f)
    return data


TEXT: dict[str, dict[str, str]] = {lang: _load(lang) for lang in LANGUAGES}


def supported(lang: str) -> str:
    """Kod obsługiwanego języka dla ustawień regionalnych („pt_BR” → „pt”); inaczej angielski."""
    code = (lang or "").lower().replace("-", "_").split("_")[0]
    return code if code in LANGUAGES else "en"


def tr(lang: str, text_id: str, /, **kwargs: object) -> str:
    """Zwraca tekst w języku `lang`; brak tekstu w danym języku = angielski."""
    table = TEXT.get(lang, TEXT["en"])
    text = table.get(text_id) or TEXT["en"][text_id]
    return text.format(**kwargs) if kwargs else text


def plural_form(lang: str, n: int) -> str:
    """Klucz formy liczebnika: „count.one”, „count.few” albo „count.many”."""
    if lang in ("pl", "ru"):
        if n == 1 or (lang == "ru" and n % 10 == 1 and n % 100 != 11):
            return "count.one"
        return "count.few" if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14) else "count.many"
    if lang == "fr":
        return "count.one" if n in (0, 1) else "count.many"
    return "count.one" if n == 1 else "count.many"


def format_number(lang: str, text: str) -> str:
    """Zamienia separatory z zapisu Pythona („1,234.5”) na separatory języka."""
    thousands, decimal = NUMBER_SEPARATORS.get(lang, NUMBER_SEPARATORS["en"])
    return text.replace(",", "\0").replace(".", decimal).replace("\0", thousands)


def count_label(lang: str, n: int) -> str:
    """„1 obiekt”, „3 obiekty”, „5 obiektów” — liczebnik w formie właściwej dla języka."""
    return tr(lang, plural_form(lang, n), n=format_number(lang, f"{n:,}"))
