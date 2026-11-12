"""Model eksportu: dane, ustawienia wyglądu i wybór arkuszy — bez zależności od QGIS."""

from __future__ import annotations

import datetime as dt
from collections.abc import Sequence
from dataclasses import asdict, dataclass, field, fields, is_dataclass, replace
from typing import Any, Optional, Union

Value = Union[None, str, int, float, bool, dt.date, dt.datetime, dt.time]

KINDS = ("text", "integer", "decimal", "date", "datetime", "time", "boolean", "link")

# Arkusz Excela ma 1 048 576 wierszy; odejmujemy wiersze tytułu i nagłówka.
# Bez tytułu nad tabelą nagłówek trafia do wiersza 1 (`Look.header_row`).
HEADER_ROW = 4
SHEET_ROWS = 1_048_576
MAX_ROWS = SHEET_ROWS - HEADER_ROW


@dataclass
class Column:
    """Jedna kolumna arkusza „Dane” i jej opis w słowniku."""

    name: str
    header: str
    kind: str = "text"
    unit: str = ""
    description: str = ""
    decimals: int = 2
    is_identifier: bool = False
    represented: bool = False
    masked: bool = False
    codes: list[tuple[str, str]] = field(default_factory=list)
    summed: bool = False  # kolumna wchodzi do wiersza sum i „Podsumowania”
    grouping: bool = True  # separator tysięcy (wyłączany np. dla roku)


@dataclass
class Summary:
    """Ustawienia wiersza sum i arkusza „Podsumowanie” (domyślnie wszystko wyłączone)."""

    row: bool = False
    sheet: bool = False
    group_by: str = ""
    measure: bool = True


def row_limit(columns: Sequence[Column], summary: Summary) -> int:
    """Ile wierszy danych się zmieści: wiersz sum pod danymi zabiera jedno miejsce."""
    totals = summary.row and any(c.summed and not c.masked for c in columns)
    return MAX_ROWS - int(totals)


@dataclass
class Package:
    """Wszystko, co trafia do pliku: kolumny, wiersze i fakty o pochodzeniu.

    `facts` to pary (klucz z i18n bez prefiksu „fact.”, wartość).
    `notes` to gotowe zdania dla odbiorcy (np. co usunięto z danych).
    """

    title: str
    columns: list[Column]
    rows: list[list[Value]]
    facts: list[tuple[str, str]] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    language: str = "pl"
    key: Optional[str] = None
    summary: Summary = field(default_factory=Summary)
    exported_at: dt.datetime = field(default_factory=lambda: dt.datetime.now().replace(microsecond=0))


@dataclass(frozen=True)
class Theme:
    """Kolorystyka pliku (kolory w zapisie RRGGBB)."""

    header_fill: str
    header_font: str
    accent: str
    soft: str


THEMES: dict[str, Theme] = {
    "navy": Theme(header_fill="1F3A5F", header_font="FFFFFF", accent="1F3A5F", soft="E8EEF6"),
    "forest": Theme(header_fill="2D5A3D", header_font="FFFFFF", accent="2D5A3D", soft="E9F2EC"),
    "graphite": Theme(header_fill="3A3F44", header_font="FFFFFF", accent="3A3F44", soft="EDEEF0"),
    "terracotta": Theme(header_fill="9A4A2F", header_font="FFFFFF", accent="9A4A2F", soft="F7ECE7"),
    "plum": Theme(header_fill="5B3A6B", header_font="FFFFFF", accent="5B3A6B", soft="F1EBF4"),
}


def theme_from_color(rgb: str) -> Theme:
    """Motyw z jednego koloru użytkownika: tło nagłówka + jasny odcień.

    Kolor czcionki nagłówka dobieramy do jasności tła, żeby tekst był czytelny.
    """
    rgb = rgb.lstrip("#").upper()
    r, g, b = (int(rgb[i : i + 2], 16) for i in (0, 2, 4))
    luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
    light = luminance > 150
    font = "000000" if light else "FFFFFF"
    # Akcent to kolor tytułów na białym tle — jasny kolor przyciemniamy.
    accent = "".join(f"{int(c * 0.45):02X}" for c in (r, g, b)) if light else rgb
    soft = "".join(f"{int(c + (255 - c) * 0.88):02X}" for c in (r, g, b))
    return Theme(header_fill=rgb, header_font=font, accent=accent, soft=soft)


# --------------------------------------------------------------------------
# Wygląd pliku i wybór arkuszy
# --------------------------------------------------------------------------

ALIGNMENTS = ("auto", "left", "center", "right")
BORDER_STYLES = ("none", "hair", "thin", "medium", "thick", "dashed", "dotted", "double")
BORDER_SCOPES = ("all", "horizontal", "outline")

# Pozycje arkusza „Słownik” (kolumna z nagłówkiem jest zawsze) i „Metryka”.
DICTIONARY_ITEMS = ("field", "kind", "unit", "description", "codes")
ABOUT_ITEMS = (
    "exported_at", "title", "layer", "note", "source", "provider", "crs", "scope", "filter",
    "count", "total", "author", "qgis", "previous", "guide", "notes",
)


@dataclass
class CellLook:
    """Wygląd grupy komórek (nagłówki albo zawartość). Kolory w zapisie RRGGBB."""

    font: str = "Calibri"
    size: int = 11
    color: str = "000000"
    bold: bool = False
    italic: bool = False
    underline: bool = False
    fill: str = ""  # puste = bez tła
    align: str = "auto"  # auto: identyfikatory i tekst do lewej, liczby do prawej
    wrap: bool = False


@dataclass
class Look:
    """Wszystkie ustawienia wyglądu arkusza z danymi."""

    as_table: bool = True  # tabela Excela (Wstaw → Tabela): filtry, nazwa, rozszerza się o nowe wiersze
    autofilter: bool = True  # filtry bez tabeli Excela
    show_title: bool = True  # tytuł i podpis nad tabelą
    title_color: str = "1F3A5F"
    title_size: int = 16
    header: CellLook = field(default_factory=lambda: CellLook(color="FFFFFF", bold=True, fill="1F3A5F", wrap=True))
    body: CellLook = field(default_factory=CellLook)
    stripes: bool = True
    stripe_color: str = "EEF2F7"
    border_style: str = "thin"
    border_color: str = "BFBFBF"
    border_scope: str = "all"
    freeze_header: bool = True
    header_comments: bool = True  # opis kolumny w dymku nagłówka
    gridlines: bool = True
    max_width: int = 60  # maksymalna szerokość kolumny dopasowywanej do treści (znaki)
    landscape: bool = True
    tab_color: str = "1F3A5F"

    @property
    def header_row(self) -> int:
        return HEADER_ROW if self.show_title else 1


@dataclass
class Sheets:
    """Które arkusze dodatkowe trafiają do pliku i co w nich jest."""

    dictionary: bool = False
    dictionary_items: tuple[str, ...] = DICTIONARY_ITEMS
    about: bool = False
    about_items: tuple[str, ...] = ABOUT_ITEMS


# --------------------------------------------------------------------------
# Kolumny liczone z geometrii
# --------------------------------------------------------------------------

GEO_PREFIX = "@"  # nazwy kolumn z geometrii („@area_ha”) — nie są polami warstwy
AREA_UNITS = ("ha", "m2")
LENGTH_UNITS = ("m", "km")
MAP_SERVICES = ("osm", "geoportal")


@dataclass
class GeoColumns:
    """Kolumny z geometrii dopisywane na końcu arkusza „Dane” (ustawienia warstwy).

    `measure` to powierzchnia na warstwie poligonowej, a długość na liniowej.
    """

    measure: bool = False
    area_unit: str = "ha"
    length_unit: str = "m"
    xy: bool = False
    wgs84: bool = False
    link: bool = False
    map: str = "osm"

    def __post_init__(self) -> None:
        # Zapisane ustawienia mogą być uszkodzone albo ze starszej wersji: nieznana wartość = domyślna.
        if self.area_unit not in AREA_UNITS:
            self.area_unit = "ha"
        if self.length_unit not in LENGTH_UNITS:
            self.length_unit = "m"
        if self.map not in MAP_SERVICES:
            self.map = "osm"

    @property
    def any(self) -> bool:
        return self.measure or self.xy or self.wgs84 or self.link


def map_url(service: str, lat: float, lon: float) -> str:
    """Adres mapy w punkcie WGS84. Parametr `mapview` Geoportalu nie jest opisany w dokumentacji
    (Geoportal sam ustawia go w adresie) — sprawdzony 2026-11-10."""
    la, lo = f"{lat:.6f}", f"{lon:.6f}"
    if service == "geoportal":
        return f"https://mapy.geoportal.gov.pl/imapnext/imap/?mapview={la},{lo},1000s"
    return f"https://www.openstreetmap.org/?mlat={la}&mlon={lo}#map=18/{la}/{lo}"


def look_with_theme(look: Look, theme: Theme) -> Look:
    """Gotowa kolorystyka nakładana na bieżące ustawienia (czcionki i układ zostają)."""
    return replace(
        look,
        title_color=theme.accent,
        header=replace(look.header, fill=theme.header_fill, color=theme.header_font),
        stripe_color=theme.soft,
        tab_color=theme.accent,
    )


def to_dict(obj: Any) -> dict[str, Any]:
    """Ustawienia jako słownik (do zapisu w QgsSettings jako JSON)."""
    return asdict(obj)


def from_dict(cls: Any, data: Any, default: Any = None) -> Any:
    """Odtwarza ustawienia ze słownika; nieznane klucze i złe typy pomija.

    Zapisane ustawienia mogą pochodzić ze starszej wersji wtyczki, więc
    każde pole sprawdzamy osobno, a brakujące dostają wartość domyślną.
    """
    default = cls() if default is None else default
    if not isinstance(data, dict):
        return default
    values: dict[str, Any] = {}
    for f in fields(cls):
        if f.name not in data:
            continue
        current = getattr(default, f.name)
        value = data[f.name]
        if is_dataclass(current):
            values[f.name] = from_dict(type(current), value, current)
        elif isinstance(current, tuple):
            if isinstance(value, (list, tuple)) and all(isinstance(v, str) for v in value):
                values[f.name] = tuple(value)
        elif type(value) is type(current):
            values[f.name] = value
    return replace(default, **values)
