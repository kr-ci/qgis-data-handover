"""Zapis paczki do XLSX (openpyxl) i odczyt poprzedniej paczki do porównania.

Jedyne miejsce, które zna układ pliku. `export()` to jedyny punkt wejścia
dla okna wtyczki i testów: przygotowuje wartości raz i z tego samego wyniku
liczy zmiany i zapisuje arkusze, żeby plik nie przeczył sam sobie.
"""

from __future__ import annotations

import contextlib
import dataclasses
import datetime as dt
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
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from . import diff, privacy
from .i18n import DATE_FORMATS, count_label, format_number, tr
from .model import DICTIONARY_ITEMS, GEO_PREFIX, HEADER_ROW, MAX_ROWS, CellLook, Column, Look, Package, Sheets, Value

FORMAT_VERSION = "1"
META_SHEET = "_handover"
CELL_LIMIT = 32_767
# Excel trzyma liczby jako double: pewnych jest 15 cyfr znaczących.
EXACT_INT_LIMIT = 10**15
WIDTH_SAMPLE = 300
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


def number_format(col: Column, lang: str) -> str:
    clean_unit = col.unit.replace('"', "").strip()
    unit = f' "{clean_unit}"' if clean_unit else ""
    if col.kind == "integer":
        return ("0" if col.is_identifier else "#,##0") + unit
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
        elif col is not None and col.is_identifier:
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
        wb = load_workbook(path, read_only=True, data_only=True)
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
        width = len(columns)
        found = [
            list(values) + [None] * (width + 1 - len(values))
            for values in wb[meta["data_sheet"]].iter_rows(
                min_row=header_row, max_row=header_row + count + 1, max_col=width + 1, values_only=True
            )
        ]
        found += [[None] * (width + 1)] * (count + 2 - len(found))
        # Ręczne zmiany po eksporcie (przestawione lub dopisane kolumny, dopisane
        # lub usunięte wiersze) dałyby cicho błędny raport — odrzucamy taki plik.
        edited = found[0] != [headers[c] for c in columns] + [None]
        edited |= any(all(v is None for v in row[:width]) for row in found[1 : count + 1])
        edited |= bool(count) and any(v is not None for v in found[count + 1])
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
    if len(pkg.rows) > MAX_ROWS:
        raise PackageError(tr(lang, "error.too_many_rows", n=len(pkg.rows), limit=MAX_ROWS))
    pkg, link_notes = _drop_excess_links(pkg)
    rows, notes = prepare_rows(pkg)
    notes = link_notes + notes
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
    if sheets.dictionary:
        _write_dictionary(wb.create_sheet(), pkg, look, sheets.dictionary_items)
    if sheets.about:
        _write_about(wb.create_sheet(), pkg, look, sheets, notes, snapshot)
    if changes is not None and snapshot is not None:
        _write_changes(wb.create_sheet(), pkg, look, changes, snapshot)
    _write_meta(wb.create_sheet(META_SHEET), pkg, len(rows), look.header_row)
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


def _grid(ws: Worksheet, look: Look, header_row: int, cols: int, rows: int, table_name: str,
          stripes: bool = True) -> None:
    """Tabela Excela albo zwykły zakres z filtrami, plus pasy wierszy."""
    if not rows:
        return
    ref = f"A{header_row}:{get_column_letter(cols)}{header_row + rows}"
    if look.as_table:
        table = Table(displayName=table_name, ref=ref)
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
        if sheets.dictionary:
            subtitle += tr(lang, "data.subtitle.dictionary", name=tr(lang, "sheet.dictionary"))
        if sheets.about:
            subtitle += tr(lang, "data.subtitle.about", name=tr(lang, "sheet.about"))
        _title(ws, pkg.title, subtitle, look)
    headers = _unique_headers(pkg.columns)
    borders = _Borders(look, len(rows) + 1, len(headers))
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
    _grid(ws, look, hr, len(headers), len(rows), "HandoverData")

    _fit_columns(ws, pkg.columns, headers, rows, lang, look, hr)
    ws.sheet_view.showGridLines = look.gridlines
    _sheet_basics(ws, look, hr)
    _page_setup(ws, pkg.title, lang, hr, look.landscape)


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
                 "changes": tr(lang, "sheet.changes")}
        guide = ["rows"]
        if look.as_table or look.autofilter:
            guide.append("filters")
        if look.header_comments:
            guide.append("tips")
        if sheets.dictionary:
            guide.append("dictionary")
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


def _write_meta(ws: Worksheet, pkg: Package, count: int, header_row: int) -> None:
    """Ukryty arkusz techniczny: pozwala następnym razem porównać wersje."""
    rows: list[tuple[Value, ...]] = [
        ("format", FORMAT_VERSION),
        ("data_sheet", tr(pkg.language, "sheet.data")),
        ("header_row", header_row),
        ("rows", count),
        ("key", pkg.key or ""),
        ("exported_at", pkg.exported_at.isoformat()),
        ("language", pkg.language),
    ]
    rows += [("column", c.name, h, c.kind) for c, h in zip(pkg.columns, _unique_headers(pkg.columns))]
    for r, values in enumerate(rows, start=1):
        for c, value in enumerate(values, start=1):
            _set_text(ws.cell(row=r, column=c), value)
    ws.sheet_state = "hidden"
