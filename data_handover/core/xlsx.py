"""Zapis paczki do XLSX (openpyxl) i odczyt poprzedniej paczki do porównania.

Jedyne miejsce, które zna układ pliku. `export()` to jedyny punkt wejścia
dla okna wtyczki i testów: przygotowuje wartości raz i z tego samego wyniku
liczy zmiany i zapisuje arkusze, żeby plik nie przeczył sam sobie.
"""

from __future__ import annotations

import contextlib
import dataclasses
import datetime as dt
import hashlib
import math
import os
import uuid
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Any, Callable, Optional, Union, cast

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell, MergedCell
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from . import diff, privacy
from .i18n import DATE_FORMATS, count_label, format_number, tr
from .model import (
    DICTIONARY_ITEMS,
    GEO_PREFIX,
    HEADER_ROW,
    MAX_ROWS,
    SHEET_ROWS,
    CellLook,
    Column,
    Look,
    Package,
    Sheets,
    Value,
    row_limit,
)

FORMAT_VERSION = "1"
META_SHEET = "_handover"
CELL_LIMIT = 32_767
# Excel trzyma liczby jako double: pewnych jest 15 cyfr znaczących.
EXACT_INT_LIMIT = 10**15
YEAR_RANGE = (1000, 2100)  # liczby z tego zakresu w polu całkowitym uznajemy za rok
WIDTH_SAMPLE = 300
TOTALS_TABLE = "HandoverData"  # nazwa tabeli Excela w arkuszu „Dane”
# Excel przyjmuje najwyżej 65 530 hiperłączy w arkuszu; plik z większą liczbą uznaje za uszkodzony.
HYPERLINK_LIMIT = 65_530
LINK_COLOR = "0563C1"  # kolor zwykłego hiperłącza w Excelu

FILL_ADDED = "E2EFDA"
FILL_REMOVED = "FCE4D6"
FILL_MODIFIED = "FFF2CC"
GREY_TEXT = "595959"


class NotHandoverFileError(ValueError):
    """Plik nie jest paczką Data Handover (lub jest uszkodzony)."""


class EditedFileError(NotHandoverFileError):
    """Paczka zmieniona ręcznie po eksporcie: wiersze lub kolumny nie zgadzają się z zapisem."""


class CanceledError(Exception):
    """Użytkownik anulował eksport; plik docelowy został nietknięty."""


Progress = Optional[Callable[[float], None]]  # dostaje 0–100
CancelCheck = Optional[Callable[[], bool]]


def _check(canceled: CancelCheck) -> None:
    if canceled is not None and canceled():
        raise CanceledError


class PackageError(ValueError):
    """Błąd zgłaszany użytkownikowi gotowym zdaniem w jego języku."""


@dataclass
class Snapshot:
    """Poprzednia paczka odczytana z pliku."""

    file: str
    key: Optional[str]
    exported_at: str
    columns: list[str]
    headers: dict[str, str]
    rows: list[list[Value]]


@dataclass
class ExportResult:
    path: str
    rows: int
    changes: list[diff.Change] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# Przygotowanie wartości (jedno źródło prawdy dla zapisu i porównania)
# --------------------------------------------------------------------------


def prepare_rows(pkg: Package) -> tuple[list[list[Value]], list[str]]:
    """Wartości dokładnie takie, jakie trafią do komórek, oraz uwagi o poprawkach."""
    truncated = illegal = 0
    out: list[list[Value]] = []
    for row in pkg.rows:
        cells: list[Value] = []
        for value, col in zip(row, pkg.columns):
            if col.masked:
                value = privacy.mask(value)
            if isinstance(value, bool):
                value = tr(pkg.language, "bool.true" if value else "bool.false")
            elif isinstance(value, float) and not math.isfinite(value):
                value = None
            elif isinstance(value, int) and abs(value) >= EXACT_INT_LIMIT:
                value = str(value)  # Excel zgubiłby ostatnie cyfry identyfikatora
            elif isinstance(value, (dt.datetime, dt.time)) and value.tzinfo is not None:
                value = value.replace(tzinfo=None)
            if isinstance(value, str):
                clean = ILLEGAL_CHARACTERS_RE.sub("", value)
                illegal += clean != value
                clean = clean.replace("\r\n", "\n").replace("\r", "\n")  # w pliku nie ma samego CR
                if len(clean) > CELL_LIMIT:
                    clean = clean[: CELL_LIMIT - 1] + "…"
                    truncated += 1
                value = clean
            cells.append(value)
        out.append(cells)
    notes = []
    if truncated:
        notes.append(tr(pkg.language, "note.truncated", n=truncated))
    if illegal:
        notes.append(tr(pkg.language, "note.illegal", n=illegal))
    return out, notes


def _plain_years(pkg: Package, rows: Sequence[Sequence[Value]]) -> Package:
    """Kopia paczki, w której kolumny z samymi latami nie mają separatora tysięcy."""
    low, high = YEAR_RANGE
    columns = list(pkg.columns)
    for i, col in enumerate(columns):
        if col.kind != "integer" or col.is_identifier or col.unit:
            continue
        values = [row[i] for row in rows if row[i] is not None]
        if values and all(
            isinstance(v, int) and not isinstance(v, bool) and low <= v <= high for v in values
        ):
            columns[i] = dataclasses.replace(col, grouping=False)
    return dataclasses.replace(pkg, columns=columns)


def number_format(col: Column, lang: str) -> str:
    clean_unit = col.unit.replace('"', "").strip()
    unit = f' "{clean_unit}"' if clean_unit else ""
    if col.kind == "integer":
        return ("0" if col.is_identifier or not col.grouping else "#,##0") + unit
    if col.kind == "decimal":
        decimals = max(0, min(col.decimals, 10))
        return ("#,##0." + "0" * decimals if decimals else "#,##0") + unit
    date = DATE_FORMATS.get(lang, DATE_FORMATS["en"])[0]
    if col.kind == "date":
        return date
    if col.kind == "datetime":
        return date + " HH:MM"
    if col.kind == "time":
        return "HH:MM:SS"
    return "@"


def display(value: Value, col: Optional[Column], lang: str) -> str:
    """Wartość jako tekst dla człowieka (arkusz „Zmiany”, komentarze, szerokość)."""
    if value is None:
        return ""
    if col is not None and col.kind == "link":
        return tr(lang, "geo.link.label")  # w komórce jest ten tekst, adres siedzi w hiperłączu
    date = DATE_FORMATS.get(lang, DATE_FORMATS["en"])[1]
    if isinstance(value, dt.datetime):
        return value.strftime(date + " %H:%M")
    if isinstance(value, dt.date):
        return value.strftime(date)
    if isinstance(value, dt.time):
        return value.strftime("%H:%M:%S")
    unit = f" {col.unit}" if col is not None and col.unit else ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if col is not None and col.kind == "decimal":
            text = f"{value:,.{max(0, min(col.decimals, 10))}f}"
        elif col is not None and (col.is_identifier or not col.grouping):
            text = str(value)
        else:
            text = f"{value:,}"
        return format_number(lang, text) + unit
    return str(value)


# --------------------------------------------------------------------------
# Odczyt poprzedniej paczki
# --------------------------------------------------------------------------


def read_snapshot(path: str) -> Snapshot:
    name = os.path.basename(path)
    try:
        wb = load_workbook(path, read_only=True)
    except Exception as exc:  # uszkodzony plik, nie-zip, nie-xlsx
        raise NotHandoverFileError(name) from exc
    try:
        if META_SHEET not in wb.sheetnames:
            raise NotHandoverFileError(name)
        meta: dict[str, str] = {}
        columns: list[str] = []
        headers: dict[str, str] = {}
        for row in wb[META_SHEET].iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            label = str(row[0])
            if label == "column":
                columns.append(str(row[1]))
                headers[str(row[1])] = str(row[2])
            else:
                meta[label] = "" if row[1] is None else str(row[1])
        if meta.get("format") != FORMAT_VERSION or meta.get("data_sheet") not in wb.sheetnames:
            raise NotHandoverFileError(name)
        header_row = int(meta["header_row"])
        count = int(meta["rows"])
        extra = 1 if meta.get("totals_row") == "1" else 0  # wiersz sum pod danymi
        width = len(columns)
        # Komórki z formułami (wiersz sum) rozpoznajemy po typie, nie po tekście: zwykły tekst
        # zaczynający się od „=” jest zapisany jako tekst. Czytamy do końca arkusza,
        # żeby zauważyć dopiski także za pustymi wierszami pod danymi.
        used = count + 2 + extra  # nagłówek, dane, wiersz sum i jeden wiersz zapasu
        cells = [
            list(row) + [None] * (width + 1 - len(row))
            for row in wb[meta["data_sheet"]].iter_rows(min_row=header_row, max_col=width + 1)
        ]
        cells += [[None] * (width + 1)] * (used - len(cells))
        found = [[c.value if c is not None else None for c in row] for row in cells]
        is_formula = [any(getattr(c, "data_type", "") == "f" for c in row) for row in cells]
        total = [getattr(c, "data_type", "") == "f" for c in cells[count + 1]] if extra else []
        # Ręczne zmiany po eksporcie (przestawione lub dopisane kolumny, dopisane
        # lub usunięte wiersze, zmienione wartości) dałyby cicho błędny raport — odrzucamy taki plik.
        edited = found[0] != [headers[c] for c in columns] + [None]
        edited |= any(all(v is None for v in row[:width]) for row in found[1 : count + 1])
        # Pod danymi (i wierszem sum) nie może być niczego; przy pustym eksporcie wiersz 1 to komunikat „brak danych”.
        edited |= any(v is not None for row in found[count + 1 + extra if count else 2 :] for v in row)
        if extra:
            # Wiersz sum stoi zaraz pod danymi i zawiera wyłącznie sumy częściowe; inaczej ktoś usunął wiersz
            # danych albo podmienił formułę. Kolumny z formułą przy eksporcie muszą ją nadal mieć.
            edited |= not any(total) or any(is_formula[1 : count + 1])
            edited |= any(is_f and not str(v).startswith("=SUBTOTAL(109,") for is_f, v in zip(total, found[count + 1]))
            wanted = [int(n) for n in meta.get("totals_cols", "").split(",") if n]
            edited |= any(not total[n - 1] for n in wanted)
        # Brak skrótu (plik sprzed jego wprowadzenia) — wartości przyjmujemy bez sprawdzenia.
        edited |= bool(meta.get("digest")) and meta["digest"] != _digest([row[:width] for row in found[1 : count + 1]])
        if edited:
            raise EditedFileError(name)
        rows = [cast(list[Value], row[:width]) for row in found[1 : count + 1]]
        return Snapshot(name, meta.get("key") or None, meta.get("exported_at", ""), columns, headers, rows)
    except (KeyError, ValueError, TypeError) as exc:
        if isinstance(exc, NotHandoverFileError):
            raise
        raise NotHandoverFileError(name) from exc
    finally:
        wb.close()


def changes_against(
    snapshot: Snapshot, pkg: Package, rows: Sequence[Sequence[Value]]
) -> tuple[list[diff.Change], list[str]]:
    key = pkg.key
    lang = pkg.language
    masked = {c.name for c in pkg.columns if c.masked}
    if key is None or key in masked:
        # Bez jawnego klucza porównanie się nie odbywa — nie wolno udawać „brak różnic”.
        raise PackageError(tr(lang, "error.key_unusable"))
    if key not in snapshot.columns:
        raise PackageError(tr(lang, "error.key_missing", key=key))
    names = [c.name for c in pkg.columns]
    # Link wynika ze współrzędnych, a z poprzedniego pliku odczytalibyśmy tylko jego tekst.
    links = {c.name for c in pkg.columns if c.kind == "link"}
    common = [n for n in names if n in snapshot.columns and n != key and n not in links]
    # Kolumna maskowana teraz mogła być jawna poprzednio: maskujemy też stare
    # wartości, inaczej pełne dane wróciłyby w „Było” i w dymkach komórek.
    old_rows = [
        [privacy.mask(v) if col in masked else v for col, v in zip(snapshot.columns, row)]
        for row in snapshot.rows
    ]
    try:
        old = diff.index_rows(snapshot.columns, old_rows, key)
        new = diff.index_rows(names, rows, key)
    except diff.DuplicateKeyError as exc:
        header = next((c.header for c in pkg.columns if c.name == key), key)
        raise PackageError(tr(lang, "error.key_duplicate", key=header, n=exc.bad_rows)) from exc
    notes = []
    added_cols = [c.header for c in pkg.columns if c.name not in snapshot.columns]
    removed_cols = [snapshot.headers[n] for n in snapshot.columns if n not in names]
    if added_cols:
        notes.append(tr(lang, "note.columns.added", names=", ".join(added_cols)))
    if removed_cols:
        notes.append(tr(lang, "note.columns.removed", names=", ".join(removed_cols)))
    return diff.compare(old, new, common), notes


# --------------------------------------------------------------------------
# Zapis
# --------------------------------------------------------------------------


def export(
    path: str,
    pkg: Package,
    look: Optional[Look] = None,
    sheets: Optional[Sheets] = None,
    previous: Optional[str] = None,
    progress: Progress = None,
    canceled: CancelCheck = None,
) -> ExportResult:
    """Zapisuje plik; z `previous` dodaje arkusz zmian względem tamtego pliku.

    `canceled` sprawdzamy co setną część wierszy i tuż przed podmianą pliku:
    po anulowaniu leci `CanceledError`, a plik docelowy zostaje taki, jaki był.
    """
    look = look or Look()
    sheets = sheets or Sheets()
    lang = pkg.language
    limit = row_limit(pkg.columns, pkg.summary)  # wiersz sum pod danymi zabiera jedno miejsce
    if len(pkg.rows) > limit:
        raise PackageError(tr(lang, "error.too_many_rows", n=len(pkg.rows), limit=limit))
    pkg, link_notes = _drop_excess_links(pkg)
    rows, notes = prepare_rows(pkg)
    pkg = _plain_years(pkg, rows)
    notes = link_notes + notes
    summary_data = summarize(pkg, rows) if pkg.summary.sheet else None
    if summary_data is not None and summary_data.group_missing:
        notes.append(tr(lang, "note.summary.group_missing", field=pkg.summary.group_by))
    if summary_data is not None and (needed := _summary_rows(pkg, summary_data)) > SHEET_ROWS:
        raise PackageError(tr(lang, "error.summary_too_long", name=tr(lang, "sheet.summary"),
                              n=needed, limit=SHEET_ROWS))
    changes: Optional[list[diff.Change]] = None
    snapshot: Optional[Snapshot] = None
    if previous:
        try:
            snapshot = read_snapshot(previous)
        except EditedFileError as exc:
            raise PackageError(tr(lang, "error.previous_edited", file=str(exc))) from exc
        except NotHandoverFileError as exc:
            raise PackageError(tr(lang, "error.not_handover", file=str(exc))) from exc
        changes, change_notes = changes_against(snapshot, pkg, rows)
        if len(changes) > MAX_ROWS:
            raise PackageError(tr(lang, "error.too_many_changes", n=len(changes), limit=MAX_ROWS))
        notes += change_notes
    dirty = sum(1 for text in _texts(pkg) if text and _clean(text) != text)
    if dirty:
        notes.append(tr(lang, "note.illegal.meta", n=dirty))
    notes = pkg.notes + notes

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    _check(canceled)
    _write_data(ws, pkg, rows, look, sheets, changes, progress, canceled)
    if summary_data is not None:
        _write_summary(wb.create_sheet(), pkg, look, summary_data)
    if sheets.dictionary:
        _write_dictionary(wb.create_sheet(), pkg, look, sheets.dictionary_items)
    if sheets.about:
        _write_about(wb.create_sheet(), pkg, look, sheets, notes, snapshot)
    if changes is not None and snapshot is not None:
        _write_changes(wb.create_sheet(), pkg, look, changes, snapshot)
    totals = _has_totals(pkg, rows)
    # Skrót liczymy z komórek tak, jak trafiły do arkusza (linki, maski, teksty „=…”).
    stored = [list(r) for r in ws.iter_rows(min_row=look.header_row + 1, max_row=look.header_row + len(rows),
                                            max_col=len(pkg.columns), values_only=True)]
    total_cols = [c for c in range(1, len(pkg.columns) + 1)
                  if totals and ws.cell(row=look.header_row + len(rows) + 1, column=c).data_type == "f"]
    _write_meta(wb.create_sheet(META_SHEET), pkg, len(rows), look.header_row, totals, _digest(stored), total_cols)
    wb.properties.title = _clean(pkg.title)
    wb.properties.creator = "Data Handover (QGIS)"
    _save(wb, path, canceled)
    if progress is not None:
        progress(100.0)
    return ExportResult(path, len(rows), changes or [], notes)


def _drop_excess_links(pkg: Package) -> tuple[Package, list[str]]:
    """Bez kolumn linku, gdy hiperłączy byłoby więcej, niż Excel przyjmie w arkuszu."""
    keep = [i for i, c in enumerate(pkg.columns) if c.kind != "link"]
    if len(keep) == len(pkg.columns) or len(pkg.rows) <= HYPERLINK_LIMIT:
        return pkg, []
    note = tr(pkg.language, "note.link.dropped", limit=format_number(pkg.language, f"{HYPERLINK_LIMIT:,}"))
    return dataclasses.replace(pkg, columns=[pkg.columns[i] for i in keep],
                               rows=[[row[i] for i in keep] for row in pkg.rows]), [note]


def _save(wb: Workbook, path: str, canceled: CancelCheck = None) -> None:
    """Zapis przez plik tymczasowy obok docelowego: błąd albo anulowanie nie niszczy poprzedniej wersji."""
    folder, base = os.path.split(os.path.abspath(path))
    # Unikalna nazwa: dwa eksporty w tle do tego samego pliku nie mogą dzielić pliku tymczasowego.
    # (mkstemp odpada: nadaje prawa 0600, a plik bywa dzielony z innymi).
    tmp = os.path.join(folder, f".~{base}.{uuid.uuid4().hex[:12]}.tmp")
    try:
        wb.save(tmp)
        _check(canceled)  # zapis openpyxl trwa długo i nie da się go przerwać w środku
        os.replace(tmp, path)
    except BaseException:
        with contextlib.suppress(OSError):
            os.remove(tmp)
        raise


def _clean(text: str) -> str:
    """Tekst, który Excel przyjmie: bez znaków sterujących i w limicie długości komórki."""
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    return text if len(text) <= CELL_LIMIT else text[: CELL_LIMIT - 1] + "…"


def _texts(pkg: Package) -> list[str]:
    """Teksty spoza wierszy danych, które trafiają do pliku (do zliczenia poprawek)."""
    texts = [pkg.title, *pkg.notes, *(value for _, value in pkg.facts)]
    for col in pkg.columns:
        texts += [col.name, col.header, col.unit, col.description, *(t for pair in col.codes for t in pair)]
    return texts


def _set_text(cell: Union[Cell, MergedCell], value: Value) -> None:
    """Wpisuje wartość; tekst zaczynający się od „=” zostaje tekstem, nie formułą.

    Każdy tekst przechodzi przez `_clean()`, więc znak sterujący w aliasie,
    opisie albo metryce nie przerywa zapisu.
    """
    cell.value = _clean(value) if isinstance(value, str) else value
    if cell.data_type == "f":
        cell.data_type = "s"


def _unique_headers(columns: Sequence[Column]) -> list[str]:
    used: set[str] = set()
    out = []
    for col in columns:
        base = _clean((col.header or col.name).strip() or col.name)
        name, n = base, 1
        while name.lower() in used:  # także nazwy już nadane z numerem
            n += 1
            name = f"{base} ({n})"
        used.add(name.lower())
        out.append(name)
    return out


# --- style ------------------------------------------------------------------


def _font(look: CellLook, **override: object) -> Font:
    values: dict[str, object] = {
        "name": look.font, "size": look.size, "bold": look.bold, "italic": look.italic,
        "underline": "single" if look.underline else None, "color": look.color,
    }
    values.update(override)
    return Font(**values)  # type: ignore[arg-type]


def _solid(color: str) -> PatternFill:
    return PatternFill("solid", start_color=color, end_color=color)


def _rule(formula: str, color: str) -> Any:
    """Reguła formatowania warunkowego: wypełnienie, gdy formuła jest prawdziwa."""
    return FormulaRule(formula=[formula], fill=_solid(color))  # type: ignore[no-untyped-call]


def _alignment(look: CellLook, identifier: bool = False, vertical: str = "top") -> Alignment:
    horizontal = None if look.align == "auto" else look.align
    if look.align == "auto" and identifier:
        # Identyfikatory to etykiety, nie wielkości — wyrównujemy je do lewej jak tekst.
        horizontal = "left"
    return Alignment(horizontal=horizontal, vertical=vertical, wrap_text=look.wrap)


class _Borders:
    """Obramowanie komórki zależne od położenia w tabeli (dla zakresu „obrys”)."""

    def __init__(self, look: Look, rows: int, cols: int) -> None:
        self.scope = look.border_scope
        style = cast(Any, look.border_style)  # wartość z BORDER_STYLES
        self.side = Side(style=style, color=look.border_color) if look.border_style != "none" else None
        self.last_row, self.last_col = rows - 1, cols - 1
        self.cache: dict[tuple[bool, bool, bool, bool], Border] = {}

    def at(self, r: int, c: int) -> Border:
        """`r`, `c` liczone od 0 od komórki nagłówka w lewym górnym rogu."""
        if self.side is None:
            return Border()
        if self.scope == "outline":
            key = (r == 0, r == self.last_row, c == 0, c == self.last_col)
        elif self.scope == "horizontal":
            key = (True, True, False, False)
        else:
            key = (True, True, True, True)
        border = self.cache.get(key)
        if border is None:
            top, bottom, left, right = (self.side if flag else Side() for flag in key)
            border = self.cache[key] = Border(top=top, bottom=bottom, left=left, right=right)
        return border


def _title(ws: Worksheet, title: str, subtitle: str, look: Look) -> None:
    _set_text(ws["A1"], title)
    ws["A1"].font = Font(name=look.body.font, size=look.title_size, bold=True, color=look.title_color)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = max(30, look.title_size * 1.9)
    _set_text(ws["A2"], subtitle)
    ws["A2"].font = Font(name=look.body.font, size=10, italic=True, color=GREY_TEXT)
    ws.row_dimensions[3].height = 6


def _header_row(ws: Worksheet, row: int, headers: Sequence[str], look: Look, borders: _Borders) -> None:
    style = look.header
    font = _font(style)
    fill = _solid(style.fill) if style.fill else None
    align = _alignment(style, vertical="center")
    for i, text in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=i)
        _set_text(cell, text)
        cell.font, cell.alignment, cell.border = font, align, borders.at(0, i - 1)
        if fill is not None:
            cell.fill = fill


def _structured(name: str) -> str:
    """Nazwa kolumny w odwołaniu strukturalnym: znaki specjalne ([ ] # @ ' i odwrotny apostrof) poprzedza `'`."""
    return "".join(f"'{ch}" if ch in "[]#@'`" else ch for ch in name)


def _grid(ws: Worksheet, look: Look, header_row: int, cols: int, rows: int, table_name: str,
          stripes: bool = True, headers: Sequence[str] = (), totals: Optional[Sequence[bool]] = None,
          label: str = "") -> None:
    """Tabela Excela albo zwykły zakres z filtrami, plus pasy wierszy.

    `totals` (flaga „sumowana” dla każdej kolumny) oznacza, że pod danymi jest wiersz sum:
    wchodzi do tabeli, ale filtr go nie obejmuje.
    """
    if not rows:
        return
    last = f"{get_column_letter(cols)}{header_row + rows}"
    ref = f"A{header_row}:{last}"
    if look.as_table:
        table = Table(displayName=table_name, ref=ref)
        if totals is not None:
            table.ref = f"A{header_row}:{get_column_letter(cols)}{header_row + rows + 1}"
            table.totalsRowCount = 1
            table.autoFilter = AutoFilter(ref=ref)
            # Własne kolumny: openpyxl sam ich nie opisze wiersza sum.
            table.tableColumns = [
                TableColumn(
                    id=i, name=headers[i - 1],
                    totalsRowFunction="sum" if totals[i - 1] else None,
                    totalsRowLabel=label if i == 1 and not totals[0] else None,
                )
                for i in range(1, cols + 1)
            ]
        # Styl tabeli „Brak”: cały wygląd pochodzi z ustawień użytkownika.
        table.tableStyleInfo = TableStyleInfo(name=None, showRowStripes=False)
        ws.add_table(table)
    elif look.autofilter:
        ws.auto_filter.ref = ref
    if stripes and look.stripes and rows > 1:
        # Formatowanie warunkowe zamiast wypełnienia komórek: pasy zostają
        # równe po sortowaniu, a kolor można wybrać dowolnie.
        body = f"A{header_row + 1}:{get_column_letter(cols)}{header_row + rows}"
        rule = _rule(f"MOD(ROW()-{header_row},2)=0", look.stripe_color)
        ws.conditional_formatting.add(body, rule)


def _page_setup(ws: Worksheet, title: str, lang: str, header_row: Optional[int], landscape: bool) -> None:
    if header_row is not None:
        ws.print_title_rows = f"{header_row}:{header_row}"
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties()
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    footer = ws.oddFooter
    assert footer is not None  # openpyxl tworzy stopkę z arkuszem
    footer.left.text = _clean(title).replace("&", "&&")
    footer.right.text = tr(lang, "footer.page")


def _sheet_basics(ws: Worksheet, look: Look, freeze_row: Optional[int]) -> None:
    ws.sheet_properties.tabColor = look.tab_color
    if freeze_row is not None and look.freeze_header:
        ws.freeze_panes = f"A{freeze_row + 1}"


# --- arkusz „Dane” ------------------------------------------------------------


def _write_data(
    ws: Worksheet, pkg: Package, rows: list[list[Value]], look: Look, sheets: Sheets,
    changes: Optional[list[diff.Change]], progress: Progress = None, canceled: CancelCheck = None,
) -> None:
    lang = pkg.language
    hr = look.header_row
    ws.title = tr(lang, "sheet.data")
    if look.show_title:
        subtitle = tr(lang, "data.subtitle", count=count_label(lang, len(rows)),
                      date=display(pkg.exported_at, None, lang))
        if pkg.summary.sheet:
            subtitle += tr(lang, "data.subtitle.summary", name=tr(lang, "sheet.summary"))
        if sheets.dictionary:
            subtitle += tr(lang, "data.subtitle.dictionary", name=tr(lang, "sheet.dictionary"))
        if sheets.about:
            subtitle += tr(lang, "data.subtitle.about", name=tr(lang, "sheet.about"))
        _title(ws, pkg.title, subtitle, look)
    headers = _unique_headers(pkg.columns)
    summed = [c.summed and not c.masked for c in pkg.columns]
    totals = _has_totals(pkg, rows)
    borders = _Borders(look, len(rows) + 1 + totals, len(headers))
    _header_row(ws, hr, headers, look, borders)

    style = look.body
    font = _font(style) if style != CellLook() else None  # domyślna czcionka skoroszytu = bez zapisu
    fill = _solid(style.fill) if style.fill else None
    formats = [number_format(c, lang) for c in pkg.columns]
    aligns = [_alignment(style, c.is_identifier) for c in pkg.columns]
    links = {i for i, c in enumerate(pkg.columns) if c.kind == "link"}
    link_font = _font(style, color=LINK_COLOR, underline="single")
    link_label = tr(lang, "geo.link.label")
    step = max(1, len(rows) // 100)
    for r, row in enumerate(rows, start=1):
        if r % step == 0:
            _check(canceled)
            if progress is not None:
                progress(90.0 * r / len(rows))  # ostatnie 10% to zapis pliku
        for c, value in enumerate(row):
            cell = ws.cell(row=hr + r, column=c + 1)
            if c in links and isinstance(value, str):
                _set_text(cell, link_label)
                cell.hyperlink = value
                cell.font = link_font
            else:
                _set_text(cell, value)
            cell.number_format = formats[c]
            cell.alignment = aligns[c]
            cell.border = borders.at(r, c)
            if font is not None and c not in links:
                cell.font = font
            if fill is not None:
                cell.fill = fill

    label = tr(lang, "totals.label")
    if totals:
        _write_totals(ws, look, hr, headers, summed, label, formats, aligns, borders, len(rows))

    if look.header_comments:
        for i, col in enumerate(pkg.columns, start=1):
            tip = _header_tip(col)
            if tip:
                comment = Comment(_clean(tip), "Data Handover")
                comment.width, comment.height = 320, 140
                ws.cell(row=hr, column=i).comment = comment

    if not rows:
        _set_text(ws.cell(row=hr + 1, column=1), tr(lang, "data.empty"))
        ws.cell(row=hr + 1, column=1).font = Font(name=style.font, italic=True, color=GREY_TEXT)
    if changes:
        _highlight(ws, pkg, rows, changes, hr)  # przed pasami: wyższe pierwszeństwo
    _grid(ws, look, hr, len(headers), len(rows), TOTALS_TABLE,
          headers=headers, totals=summed if totals else None, label=label)

    # Etykieta wiersza sum też ma się zmieścić w pierwszej kolumnie.
    sized = rows + [[label] + [None] * (len(headers) - 1)] if totals and not summed[0] else rows
    _fit_columns(ws, pkg.columns, headers, sized, lang, look, hr)
    ws.sheet_view.showGridLines = look.gridlines
    _sheet_basics(ws, look, hr)
    _page_setup(ws, pkg.title, lang, hr, look.landscape)


def _has_totals(pkg: Package, rows: Sequence[Sequence[Value]]) -> bool:
    """Czy pod „Danymi” jest wiersz sum: włączony, są wiersze i jest co sumować."""
    return bool(pkg.summary.row and rows and any(c.summed and not c.masked for c in pkg.columns))


def _style_total(cell: Union[Cell, MergedCell], look: Look, base: Border, align: Alignment) -> None:
    """Wygląd wiersza „Razem”: pogrubienie, tło pasów, gruba górna krawędź w kolorze tytułu."""
    edge = Side(style="medium", color=look.title_color)
    cell.border = Border(top=edge, bottom=base.bottom, left=base.left, right=base.right)
    cell.font, cell.fill, cell.alignment = _font(look.body, bold=True), _solid(look.stripe_color), align


def _write_totals(
    ws: Worksheet, look: Look, hr: int, headers: Sequence[str], summed: Sequence[bool], label: str,
    formats: Sequence[str], aligns: Sequence[Alignment], borders: _Borders, count: int,
) -> None:
    """Wiersz sum tuż pod danymi: SUBTOTAL(109) liczy tylko wiersze widoczne po filtrze."""
    r = count + 1
    for c, name in enumerate(headers):
        cell = ws.cell(row=hr + r, column=c + 1)
        letter = get_column_letter(c + 1)
        if summed[c]:
            if look.as_table and "\n" not in name:
                cell.value = f"=SUBTOTAL(109,{TOTALS_TABLE}[[{_structured(name)}]])"
            else:  # w zakresie albo przy nagłówku z nową linią: zwykły zakres A1
                cell.value = f"=SUBTOTAL(109,{letter}{hr + 1}:{letter}{hr + count})"
            cell.number_format = formats[c]
        elif c == 0:
            _set_text(cell, label)
        _style_total(cell, look, borders.at(r, c), aligns[c])


def _header_tip(col: Column) -> str:
    parts = [p for p in (col.description, f"[{col.unit}]" if col.unit else "") if p]
    if col.codes and not col.masked:  # etykiety kolumny maskowanej to też dane osobowe
        parts.append("; ".join(f"{code} = {label}" for code, label in col.codes[:12]))
    return "\n".join(parts)


def _highlight(
    ws: Worksheet, pkg: Package, rows: list[list[Value]], changes: list[diff.Change], header_row: int
) -> None:
    """Podświetla dodane wiersze i zmienione komórki; dymek pokazuje dawną wartość.

    Podświetlenie to formatowanie warunkowe dodane przed pasami wierszy, więc
    pasy go nie zasłaniają.
    """
    key = pkg.key
    if key is None:
        return
    names = [c.name for c in pkg.columns]
    key_pos = names.index(key)
    row_of = {diff.key_text(row[key_pos]): r for r, row in enumerate(rows, start=header_row + 1)}
    col_by_name = {c.name: c for c in pkg.columns}
    last = get_column_letter(len(names))
    added: list[str] = []
    modified: list[str] = []
    for change in changes:
        r = row_of.get(change.key)
        if r is None:
            continue
        if change.kind == "added":
            added.append(f"A{r}:{last}{r}")
        elif change.kind == "modified" and change.column is not None:
            cell = ws.cell(row=r, column=names.index(change.column) + 1)
            modified.append(cell.coordinate)
            old = display(change.old, col_by_name[change.column], pkg.language) or "—"
            cell.comment = Comment(_clean(tr(pkg.language, "change.comment", old=old)), "Data Handover")
    # ponytail: adresy stałe — po sortowaniu w Excelu podświetlenie zostaje w miejscu; arkusz „Zmiany” jest wzorcem
    for ranges, color in ((modified, FILL_MODIFIED), (added, FILL_ADDED)):
        if ranges:
            ws.conditional_formatting.add(" ".join(ranges), _rule("TRUE", color))


def _fit_columns(
    ws: Worksheet, columns: Sequence[Column], headers: Sequence[str], rows: list[list[Value]], lang: str,
    look: Look, header_row: int,
) -> None:
    # Szerokość kolumny w Excelu liczy się w znakach czcionki 11 pt; większa czcionka = szerzej.
    body_scale = look.body.size / 11 * (1.1 if look.body.bold else 1.0)
    head_scale = look.header.size / 11 * (1.2 if look.header.bold else 1.05)
    lines = 1
    for i, (col, header) in enumerate(zip(columns, headers)):
        sample = [len(display(row[i], col, lang)) * body_scale for row in rows[:WIDTH_SAMPLE]]
        longest_word = max((len(w) for w in header.split()), default=0) * head_scale
        whole = (min(len(header), 18) if look.header.wrap else len(header)) * head_scale
        width = math.ceil(max([longest_word, whole] + sample)) + 3
        width = max(8, min(width, look.max_width))
        ws.column_dimensions[get_column_letter(i + 1)].width = width
        if look.header.wrap:
            lines = max(lines, math.ceil(len(header) * head_scale / max(width - 3, 1)))
    ws.row_dimensions[header_row].height = look.header.size * 1.45 * min(lines, 4) + 8


# --- arkusz „Podsumowanie” ----------------------------------------------------


@dataclass
class SummaryData:
    """Liczby do arkusza „Podsumowanie”, policzone z wyniku `prepare_rows()`."""

    count: int
    sums: list[tuple[Column, float]]  # kolumny sumowane (bez maskowanych) i suma całości
    groups: Optional[tuple[Column, list[tuple[Value, int, list[float]]]]]  # (grupa, liczba, sumy jak w `sums`)
    gaps: list[tuple[str, int]]  # nagłówek kolumny i liczba pustych komórek (tylko > 0)
    masked: list[str]  # nagłówki kolumn maskowanych (pominięte)
    group_missing: bool = False  # `group_by` nie jest kolumną pliku


def _total(values: Sequence[Value]) -> float:
    """Suma wartości liczbowych; tekst (np. długi identyfikator), `None` i `bool` pomijamy."""
    return math.fsum(v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool))


def summarize(pkg: Package, rows: Sequence[Sequence[Value]]) -> SummaryData:
    """Całość, grupy według `Summary.group_by` i braki danych — jedna funkcja dla całego arkusza."""
    lang, name = pkg.language, pkg.summary.group_by
    headers = _unique_headers(pkg.columns)
    summed = [i for i, c in enumerate(pkg.columns) if c.summed and not c.masked]
    sums = [(pkg.columns[i], _total([row[i] for row in rows])) for i in summed]
    gaps = [] if not rows else [
        (headers[i], n) for i, c in enumerate(pkg.columns)
        if not c.masked and (n := sum(1 for row in rows if row[i] is None or row[i] == ""))
    ]
    masked = [headers[i] for i, c in enumerate(pkg.columns) if c.masked]
    index = next((i for i, c in enumerate(pkg.columns) if c.name == name), None)
    groups = None
    if rows and index is not None and not pkg.columns[index].masked:
        col = pkg.columns[index]
        members: dict[Value, list[Sequence[Value]]] = {}
        for row in rows:
            members.setdefault(None if row[index] == "" else row[index], []).append(row)
        ordered = sorted(members, key=lambda v: (-len(members[v]), v is None, display(v, col, lang)))
        groups = (col, [(v, len(members[v]), [_total([r[i] for r in members[v]]) for i in summed])
                        for v in ordered])
    return SummaryData(len(rows), sums, groups, gaps, masked, group_missing=bool(name) and index is None)


def _write_summary(ws: Worksheet, pkg: Package, look: Look, data: SummaryData) -> None:
    lang = pkg.language
    ws.title = tr(lang, "sheet.summary")
    ws.sheet_view.showGridLines = False
    subtitle = tr(lang, "summary.subtitle", count=count_label(lang, data.count),
                  date=display(pkg.exported_at, None, lang))
    _title(ws, tr(lang, "sheet.summary"), subtitle, look)
    headers = {c.name: h for c, h in zip(pkg.columns, _unique_headers(pkg.columns))}
    line = Side(style="thin", color=look.border_color)
    body, plain = _font(look.body), _alignment(look.body)
    count_format = "#,##0"
    widths = {1: 26.0}

    def section(r: int, text: str) -> None:
        cell = ws.cell(row=r, column=1)
        _set_text(cell, text)
        cell.font = _font(look.body, size=look.body.size + 1, bold=True, color=look.title_color)

    def table(r: int, head: Sequence[str], rows: Sequence[Sequence[Value]], formats: Sequence[str],
              total: Optional[Sequence[Value]] = None) -> int:
        """Rysuje tabelę od wiersza `r`; zwraca numer jej ostatniego wiersza."""
        everything = list(rows) + ([total] if total is not None else [])
        borders = _Borders(look, len(everything) + 1, len(head))
        _header_row(ws, r, head, look, borders)
        for j, values in enumerate(everything, start=1):
            for c, value in enumerate(values):
                cell = ws.cell(row=r + j, column=c + 1)
                _set_text(cell, value)
                cell.number_format, cell.border, cell.font, cell.alignment = formats[c], borders.at(j, c), body, plain
                if look.stripes and j % 2 == 0:
                    cell.fill = _solid(look.stripe_color)
                if total is not None and j == len(everything):
                    _style_total(cell, look, borders.at(j, c), plain)
                widths[c + 1] = max(widths.get(c + 1, 16.0), min(len(str(display(value, None, lang))) + 3.0, 50.0))
        for c, text in enumerate(head):
            widths[c + 1] = max(widths.get(c + 1, 16.0), min(len(text) * 1.2 + 3, 40.0))
        return r + len(everything)

    r = HEADER_ROW
    section(r, tr(lang, "summary.whole"))
    facts: list[tuple[str, Value, str]] = [(tr(lang, "summary.count"), data.count, count_format)]
    facts += [(tr(lang, "summary.sum", header=headers[c.name]), total, number_format(c, lang))
              for c, total in data.sums]
    for label, value, fmt in facts:
        r += 1
        a, b = ws.cell(row=r, column=1), ws.cell(row=r, column=2)
        _set_text(a, label)
        a.font, a.fill = _font(look.body, bold=True, color=look.title_color), _solid(look.stripe_color)
        a.border = Border(bottom=line)
        _set_text(b, value)
        b.number_format, b.font, b.border, b.alignment = fmt, body, Border(bottom=line), plain
        widths[1] = max(widths[1], min(len(label) + 3.0, 50.0))
    r += 2

    if data.groups is not None:
        col, groups = data.groups
        section(r, tr(lang, "summary.by", header=headers[col.name]))
        head = [headers[col.name], tr(lang, "summary.count")] + [
            f"{headers[c.name]} [{c.unit}]" if c.unit else headers[c.name] for c, _ in data.sums]
        formats = ["@" if col.kind == "link" else number_format(col, lang), count_format] + [
            number_format(c, lang) for c, _ in data.sums]
        rows: list[list[Value]] = [
            [tr(lang, "summary.empty_group") if v is None else v, n, *sums] for v, n, sums in groups]
        total: list[Value] = [tr(lang, "summary.total"), data.count, *(t for _, t in data.sums)]
        # Etykieta „(brak)” jest tekstem: format kolumny grupującej (np. daty) jej nie zmienia.
        r = table(r + 1, head, rows, formats, total) + 2

    if data.count:
        section(r, tr(lang, "summary.gaps"))
        sentences = []
        if data.gaps:
            share: list[list[Value]] = [[name, n, n / data.count] for name, n in data.gaps]
            r = table(r + 1, [tr(lang, "summary.gaps.column"), tr(lang, "summary.gaps.empty"),
                              tr(lang, "summary.gaps.share")], share, ["@", count_format, "0.0%"])
            if len(data.gaps) < len(pkg.columns) - len(data.masked):
                sentences.append(tr(lang, "summary.gaps.rest"))
        else:
            sentences.append(tr(lang, "summary.gaps.none"))
        if data.masked:
            sentences.append(tr(lang, "summary.gaps.masked", names=", ".join(data.masked)))
        if sentences:
            cell = ws.cell(row=r + 1, column=1)
            _set_text(cell, " ".join(sentences))
            cell.font = Font(name=look.body.font, italic=True, color=GREY_TEXT)
    for c, width in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = width
    _sheet_basics(ws, look, None)
    _page_setup(ws, pkg.title, lang, None, landscape=False)


# --- arkusze dodatkowe ------------------------------------------------------


def _write_dictionary(ws: Worksheet, pkg: Package, look: Look, items: Sequence[str]) -> None:
    lang = pkg.language
    ws.title = tr(lang, "sheet.dictionary")
    ws.sheet_view.showGridLines = False
    _title(ws, tr(lang, "dict.title"), tr(lang, "dict.subtitle", data=tr(lang, "sheet.data")), look)
    keys = ["header"] + [k for k in DICTIONARY_ITEMS if k in items]
    borders = _Borders(look, len(pkg.columns) + 1, len(keys))
    _header_row(ws, HEADER_ROW, [tr(lang, f"dict.col.{k}") for k in keys], look, borders)
    wrap = Alignment(vertical="top", wrap_text=True)
    font, bold = _font(look.body), _font(look.body, bold=True)
    for r, (col, header) in enumerate(zip(pkg.columns, _unique_headers(pkg.columns)), start=1):
        remarks = [col.description] if col.description else []
        if col.is_identifier:
            remarks.append(tr(lang, "dict.identifier"))
        if col.represented:
            remarks.append(tr(lang, "dict.represented"))
        if col.masked:
            remarks.append(tr(lang, "dict.masked"))
        values = {
            "header": header,
            "field": tr(lang, "dict.geometry_source") if col.name.startswith(GEO_PREFIX) else col.name,
            "kind": tr(lang, f"kind.{col.kind}"),
            "unit": col.unit,
            "description": "\n".join(remarks),
            "codes": "" if col.masked else "\n".join(f"{code} = {label}" for code, label in col.codes),
        }
        for c, key in enumerate(keys):
            cell = ws.cell(row=HEADER_ROW + r, column=c + 1)
            _set_text(cell, values[key] or None)
            cell.alignment = wrap
            cell.font = bold if key == "header" else font
            cell.border = borders.at(r, c)
    _grid(ws, look, HEADER_ROW, len(keys), len(pkg.columns), "HandoverDictionary", stripes=False)
    widths = {"header": 28, "field": 20, "kind": 18, "unit": 12, "description": 55, "codes": 45}
    for i, key in enumerate(keys, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths[key] * max(1.0, look.body.size / 11)
    _sheet_basics(ws, look, HEADER_ROW)
    _page_setup(ws, pkg.title, lang, HEADER_ROW, look.landscape)


def _write_about(
    ws: Worksheet, pkg: Package, look: Look, sheets: Sheets, notes: Sequence[str], snapshot: Optional[Snapshot]
) -> None:
    lang = pkg.language
    items = sheets.about_items
    ws.title = tr(lang, "sheet.about")
    ws.sheet_view.showGridLines = False
    _title(ws, tr(lang, "about.title"), tr(lang, "about.subtitle"), look)
    label_fill = _solid(look.stripe_color)
    line = Side(style="thin", color=look.border_color)
    wrap = Alignment(vertical="top", wrap_text=True)
    label_font = _font(look.body, bold=True, color=look.title_color)
    value_font = _font(look.body)
    facts = [("exported_at", display(pkg.exported_at, None, lang))] + list(pkg.facts)
    if snapshot is not None:
        facts.append(("previous", f"{snapshot.file} ({_snapshot_date(snapshot, lang)})"))
    r = HEADER_ROW
    for key, value in facts:
        if key not in items:
            continue
        label = ws.cell(row=r, column=1)
        _set_text(label, tr(lang, f"fact.{key}"))
        label.font, label.fill, label.alignment = label_font, label_fill, wrap
        cell = ws.cell(row=r, column=2)
        _set_text(cell, value if value else tr(lang, "none"))
        cell.font, cell.alignment = value_font, wrap
        for c in (label, cell):
            c.border = Border(bottom=line)
        r += 1

    def section(title: str, lines: Sequence[str], start: int, numbered: bool) -> int:
        cell = ws.cell(row=start, column=1)
        _set_text(cell, title)
        cell.font = _font(look.body, size=look.body.size + 1, bold=True, color=look.title_color)
        for i, text in enumerate(lines, start=1):
            body = ws.cell(row=start + i, column=1)
            _set_text(body, f"{i}. {text}" if numbered else f"• {text}")
            body.font = value_font
            body.alignment = Alignment(vertical="top", wrap_text=True)
            ws.merge_cells(start_row=start + i, start_column=1, end_row=start + i, end_column=2)
            ws.row_dimensions[start + i].height = 30
        return start + len(lines) + 2

    r += 1
    if "guide" in items:
        names = {"data": tr(lang, "sheet.data"), "dictionary": tr(lang, "sheet.dictionary"),
                 "changes": tr(lang, "sheet.changes"), "summary": tr(lang, "sheet.summary")}
        guide = ["rows"]
        if look.as_table or look.autofilter:
            guide.append("filters")
        if look.header_comments:
            guide.append("tips")
        if sheets.dictionary:
            guide.append("dictionary")
        if pkg.summary.sheet:
            guide.append("summary")
        if _has_totals(pkg, pkg.rows):
            guide.append("totals")
        if snapshot is not None:
            guide.append("changes")
        guide.append("source")
        r = section(tr(lang, "about.guide"), [tr(lang, f"about.guide.{g}", **names) for g in guide], r, True)
    if "notes" in items and notes:
        section(tr(lang, "about.notes"), notes, r, False)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 90
    _sheet_basics(ws, look, None)
    _page_setup(ws, pkg.title, lang, None, landscape=False)


def _write_changes(
    ws: Worksheet, pkg: Package, look: Look, changes: list[diff.Change], snapshot: Snapshot
) -> None:
    lang = pkg.language
    ws.title = tr(lang, "sheet.changes")
    ws.sheet_view.showGridLines = False
    counts = {k: sum(1 for c in changes if c.kind == k) for k in ("added", "removed", "modified")}
    subtitle = tr(lang, "changes.subtitle", file=snapshot.file, date=_snapshot_date(snapshot, lang), **counts)
    _title(ws, tr(lang, "changes.title"), subtitle, look)
    _sheet_basics(ws, look, None)
    by_name = {c.name: c for c in pkg.columns}
    key_col = by_name.get(pkg.key or "")
    headers = [
        tr(lang, "changes.col.kind"),
        key_col.header if key_col else (pkg.key or ""),
        tr(lang, "changes.col.column"),
        tr(lang, "changes.col.old"),
        tr(lang, "changes.col.new"),
    ]
    if not changes:
        _set_text(ws.cell(row=HEADER_ROW, column=1), tr(lang, "changes.none"))
        ws.column_dimensions["A"].width = 60
        return
    borders = _Borders(look, len(changes) + 1, len(headers))
    _header_row(ws, HEADER_ROW, headers, look, borders)
    fills = {kind: _solid(color) for kind, color in
             (("added", FILL_ADDED), ("removed", FILL_REMOVED), ("modified", FILL_MODIFIED))}
    wrap = Alignment(vertical="top", wrap_text=True)
    font = _font(look.body)
    for r, change in enumerate(changes, start=1):
        col = by_name.get(change.column or "")
        values = (
            tr(lang, f"change.{change.kind}"),
            change.key,
            col.header if col else "",
            display(change.old, col, lang),
            display(change.new, col, lang),
        )
        for c, value in enumerate(values):
            cell = ws.cell(row=HEADER_ROW + r, column=c + 1)
            _set_text(cell, value or None)
            cell.number_format = "@"
            cell.alignment, cell.font, cell.border = wrap, font, borders.at(r, c)
        ws.cell(row=HEADER_ROW + r, column=1).fill = fills[change.kind]
    _grid(ws, look, HEADER_ROW, len(headers), len(changes), "HandoverChanges", stripes=False)
    for letter, width in zip("ABCDE", (14, 22, 26, 40, 40)):
        ws.column_dimensions[letter].width = width * max(1.0, look.body.size / 11)
    _sheet_basics(ws, look, HEADER_ROW)
    _page_setup(ws, pkg.title, lang, HEADER_ROW, look.landscape)


def _snapshot_date(snapshot: Snapshot, lang: str) -> str:
    try:
        return display(dt.datetime.fromisoformat(snapshot.exported_at), None, lang)
    except ValueError:
        return snapshot.exported_at


def _to_second(moment: dt.datetime) -> dt.datetime:
    return moment.replace(microsecond=0) + dt.timedelta(seconds=moment.microsecond >= 500_000)


def _digest_cell(value: object) -> str:
    """Wartość komórki w postaci, która nie zmienia się po zapisie przez Excela.

    Excel trzyma 15 cyfr znaczących, daty i czasy z dokładnością do sekundy,
    a pusty tekst zapisuje jako pustą komórkę; openpyxl czyta datę jako `datetime`.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return format(float(value) + 0.0, ".15g")  # + 0.0: „-0” po zapisie wraca jako 0
    if isinstance(value, dt.datetime):
        return _to_second(value).isoformat()
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day).isoformat()
    if isinstance(value, dt.time):
        return _to_second(dt.datetime.combine(dt.date(2000, 1, 1), value)).time().isoformat()
    return str(value).replace("\r\n", "\n").replace("\r", "\n")  # XML nie zachowuje samego CR


def _digest(rows: Sequence[Sequence[object]]) -> str:
    """Skrót SHA-256 wartości komórek danych (kolejność kolumn ma znaczenie, wierszy nie: odbiorca sortuje)."""
    text = "\x1e".join(sorted("\x1f".join(_digest_cell(v) for v in row) for row in rows))
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _summary_rows(pkg: Package, data: SummaryData) -> int:
    """Numer ostatniego wiersza arkusza „Podsumowanie” — ten sam rachunek co w `_write_summary()`."""
    r = HEADER_ROW + 1 + len(data.sums)  # całość: nagłówek sekcji, liczba wierszy i sumy
    if not data.count:
        return r
    r += 2  # odstęp
    if data.groups is not None:
        r += 1 + len(data.groups[1]) + 1 + 2  # nagłówek sekcji, grupy, „Razem”, odstęp
    if not data.gaps:
        return r + 1  # zdanie „brak pustych komórek”
    r += 1 + len(data.gaps)  # nagłówek sekcji i tabela braków
    sentence = len(data.gaps) < len(pkg.columns) - len(data.masked) or bool(data.masked)
    return r + int(sentence)


def _write_meta(ws: Worksheet, pkg: Package, count: int, header_row: int, totals: bool = False,
                digest: str = "", total_cols: Sequence[int] = ()) -> None:
    """Ukryty arkusz techniczny: pozwala następnym razem porównać wersje.

    `digest` to skrót wartości z „Danych”, `total_cols` — numery kolumn (od 1) z formułą w wierszu sum;
    oba służą do odrzucenia pliku zmienionego ręcznie po eksporcie.
    """
    rows: list[tuple[Value, ...]] = [
        ("format", FORMAT_VERSION),
        ("data_sheet", tr(pkg.language, "sheet.data")),
        ("header_row", header_row),
        ("rows", count),
        ("totals_row", int(totals)),
        ("totals_cols", ",".join(map(str, total_cols))),
        ("digest", digest),
        ("key", pkg.key or ""),
        ("exported_at", pkg.exported_at.isoformat()),
        ("language", pkg.language),
    ]
    rows += [("column", c.name, h, c.kind) for c, h in zip(pkg.columns, _unique_headers(pkg.columns))]
    for r, values in enumerate(rows, start=1):
        for c, value in enumerate(values, start=1):
            _set_text(ws.cell(row=r, column=c), value)
    ws.sheet_state = "hidden"
