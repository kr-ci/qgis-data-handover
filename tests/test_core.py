"""Testy rdzenia: zapis XLSX, porównanie wersji, prywatność. Bez QGIS."""

from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from data_handover.core import diff, privacy  # noqa: E402
from data_handover.core.i18n import LANGUAGES, TEXT, count_label  # noqa: E402
from data_handover.core.model import (  # noqa: E402
    HEADER_ROW,
    THEMES,
    CellLook,
    Column,
    Look,
    Package,
    Sheets,
    from_dict,
    look_with_theme,
    theme_from_color,
    to_dict,
)
from data_handover.core.xlsx import META_SHEET, PackageError, export, number_format, read_snapshot  # noqa: E402

LOOK = Look()
ALL = Sheets(dictionary=True, about=True)


def parcels(rows: list[list[object]] | None = None, **kw: object) -> Package:
    columns = [
        Column("nr_dz", "Nr działki", "text", is_identifier=True),
        Column("teryt", "Obręb (TERYT)", "text", is_identifier=True),
        Column("pow_m2", "Powierzchnia", "decimal", unit="m²", decimals=2, description="powierzchnia ewidencyjna"),
        Column("przezn", "Przeznaczenie", "text", represented=True, codes=[("MN", "zabudowa mieszkaniowa")]),
        Column("data_akt", "Aktualność", "date"),
        Column("wlasc", "Właściciel", "text", masked=True),
        Column("aktywna", "Aktywna", "boolean"),
    ]
    data = rows if rows is not None else [
        ["12/3", "0214011", 1523.456, "zabudowa mieszkaniowa", dt.date(2024, 5, 3), "Jan Kowalski", True],
        ["12/4", "0214011", 980.1, "usługi", dt.date(2024, 5, 3), None, False],
    ]
    pkg = Package(
        title="Działki — ul. Leśna",
        columns=columns,
        rows=data,  # type: ignore[arg-type]
        facts=[("layer", "dzialki"), ("crs", "EPSG:2180")],
        language="pl",
        key="nr_dz",
        exported_at=dt.datetime(2026, 11, 7, 14, 32),
    )
    for name, value in kw.items():
        setattr(pkg, name, value)
    return pkg


def test_sheets_layout_and_types(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(), LOOK, sheets=ALL)
    wb = load_workbook(path)
    assert wb.sheetnames == ["Dane", "Słownik", "Metryka", META_SHEET]
    assert wb[META_SHEET].sheet_state == "hidden"
    ws = wb["Dane"]
    assert ws["A1"].value == "Działki — ul. Leśna"
    assert "2 obiekty" in ws["A2"].value
    assert [c.value for c in ws[HEADER_ROW]] == [
        "Nr działki", "Obręb (TERYT)", "Powierzchnia", "Przeznaczenie", "Aktualność", "Właściciel", "Aktywna",
    ]
    row = [c for c in ws[HEADER_ROW + 1]]
    assert row[1].value == "0214011" and row[1].data_type == "s"  # zero wiodące zostaje
    assert row[2].value == pytest.approx(1523.456)
    assert row[2].number_format == '#,##0.00 "m²"'
    assert row[4].value == dt.datetime(2024, 5, 3) and row[4].number_format == "DD.MM.YYYY"
    assert row[5].value == "J*** K***"
    assert row[6].value == "tak"
    assert ws.freeze_panes == f"A{HEADER_ROW + 1}"
    table = ws.tables["HandoverData"]
    assert table.ref == f"A{HEADER_ROW}:G{HEADER_ROW + 2}"
    assert "powierzchnia ewidencyjna" in ws.cell(row=HEADER_ROW, column=3).comment.text


def test_dictionary_and_about(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(notes=["Pominięto kolumnę z danymi osobowymi: pesel"]), LOOK, sheets=ALL)
    wb = load_workbook(path)
    d = wb["Słownik"]
    rows = {d.cell(row=r, column=1).value: r for r in range(HEADER_ROW + 1, d.max_row + 1)}
    pow_row = rows["Powierzchnia"]
    assert d.cell(row=pow_row, column=3).value == "liczba dziesiętna"
    assert d.cell(row=pow_row, column=4).value == "m²"
    assert "MN = zabudowa mieszkaniowa" in d.cell(row=rows["Przeznaczenie"], column=6).value
    assert "zamaskowano" in d.cell(row=rows["Właściciel"], column=5).value
    about = [c.value for c in wb["Metryka"]["A"] if c.value] + [c.value for c in wb["Metryka"]["B"] if c.value]
    assert "EPSG:2180" in about
    assert any("Pominięto kolumnę" in str(v) for v in about)


def test_formula_like_text_stays_text(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    rows = [["=HYPERLINK(\"http://x\")", "1", 1.0, "a", None, None, None]]
    export(path, parcels(rows=rows), LOOK)
    cell = load_workbook(path)["Dane"].cell(row=HEADER_ROW + 1, column=1)
    assert cell.data_type == "s" and cell.value.startswith("=")


def test_illegal_characters_and_long_text(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    rows = [["a\x00b", "1", 1.0, "x" * 40_000, None, None, None]]
    result = export(path, parcels(rows=rows), LOOK)
    ws = load_workbook(path)["Dane"]
    assert ws.cell(row=HEADER_ROW + 1, column=1).value == "ab"
    assert len(ws.cell(row=HEADER_ROW + 1, column=4).value) == 32_767
    assert any("Obcięto 1" in n for n in result.notes)


def test_empty_scope_is_explicit(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(rows=[]), LOOK)
    ws = load_workbook(path)["Dane"]
    assert not ws.tables
    assert "Brak obiektów" in ws.cell(row=HEADER_ROW + 1, column=1).value


def test_changes_roundtrip(tmp_path: Path) -> None:
    first = str(tmp_path / "v1.xlsx")
    export(first, parcels(), LOOK)
    snap = read_snapshot(first)
    assert snap.key == "nr_dz" and len(snap.rows) == 2

    rows = [
        ["12/3", "0214011", 1523.456, "usługi", dt.date(2024, 5, 3), "Jan Kowalski", True],  # zmiana
        ["15/7", "0214011", 10.0, "usługi", None, None, False],  # nowa
    ]  # 12/4 usunięta
    second = str(tmp_path / "v2.xlsx")
    result = export(second, parcels(rows=rows), LOOK, previous=first)
    kinds = sorted((c.kind, c.key, c.column) for c in result.changes)
    assert kinds == [("added", "15/7", None), ("modified", "12/3", "przezn"), ("removed", "12/4", None)]
    wb = load_workbook(second)
    assert "Zmiany" in wb.sheetnames
    cell = wb["Dane"].cell(row=HEADER_ROW + 1, column=4)
    assert "zabudowa mieszkaniowa" in cell.comment.text
    # Podświetlenie jest formatowaniem warunkowym z pierwszeństwem przed pasami wierszy.
    rules = [(str(cf.sqref), r) for cf in wb["Dane"].conditional_formatting for r in cf.rules]
    first = min(rules, key=lambda x: x[1].priority)
    assert cell.coordinate in first[0].split() and first[1].dxf.fill.fgColor.rgb.endswith("FFF2CC")


def test_unchanged_roundtrip_has_no_changes(tmp_path: Path) -> None:
    first = str(tmp_path / "v1.xlsx")
    export(first, parcels(), LOOK)
    result = export(str(tmp_path / "v2.xlsx"), parcels(), LOOK, previous=first)
    assert result.changes == []


def test_duplicate_key_is_reported(tmp_path: Path) -> None:
    first = str(tmp_path / "v1.xlsx")
    export(first, parcels(), LOOK)
    rows = [["12/3", "1", 1.0, "", None, None, None], ["12/3", "1", 2.0, "", None, None, None]]
    with pytest.raises(PackageError, match="powtarzające"):
        export(str(tmp_path / "v2.xlsx"), parcels(rows=rows), LOOK, previous=first)


def test_foreign_file_is_rejected(tmp_path: Path) -> None:
    other = tmp_path / "obcy.xlsx"
    other.write_bytes(b"not a zip")
    with pytest.raises(PackageError, match="nie pochodzi"):
        export(str(tmp_path / "v2.xlsx"), parcels(), LOOK, previous=str(other))


def test_normalize_survives_excel() -> None:
    assert diff.normalize(5) == diff.normalize(5.0)
    assert diff.normalize(dt.date(2024, 1, 2)) == diff.normalize(dt.datetime(2024, 1, 2))
    assert diff.normalize("") is None


@pytest.mark.parametrize(
    "name, expected",
    [("nazwisko", True), ("WlascicielNazwa", True), ("tel_kontakt", True), ("hotel", False),
     ("nazwa_obiektu", False), ("kw", True), ("pow_m2", False), ("adres_zam", True)],
)
def test_privacy_by_name(name: str, expected: bool) -> None:
    assert bool(privacy.detect([(name, name)], {})) is expected


def test_privacy_by_values() -> None:
    samples = {"x": ["44051401359", "02070803628", "90090515836"]}  # poprawne PESEL
    assert privacy.detect([("x", "x")], samples)[0].reason == "wartości: PESEL"
    assert not privacy.detect([("x", "x")], {"x": ["12345678901", "11111111112", "123"]})
    assert privacy.is_pesel("44051401359") and not privacy.is_pesel("44051401358")


def test_number_formats() -> None:
    assert number_format(Column("id", "ID", "integer", is_identifier=True), "pl") == "0"
    assert number_format(Column("n", "N", "integer"), "pl") == "#,##0"
    assert number_format(Column("d", "D", "decimal", decimals=0, unit='h"a'), "pl") == '#,##0 "ha"'
    assert number_format(Column("t", "T", "datetime"), "en") == "YYYY-MM-DD HH:MM"


def test_polish_plural_and_languages_complete() -> None:
    assert count_label("pl", 1) == "1 obiekt"
    assert count_label("pl", 22) == "22 obiekty"
    assert count_label("pl", 12) == "12 obiektów"
    assert count_label("pl", 1500) == "1 500 obiektów"
    assert count_label("ru", 21) == TEXT["ru"]["count.one"].format(n=21)
    assert count_label("ru", 12) == TEXT["ru"]["count.many"].format(n=12)
    assert count_label("de", 1500) == TEXT["de"]["count.many"].format(n="1.500")


@pytest.mark.parametrize("lang", LANGUAGES)
def test_translation_complete_and_safe(lang: str) -> None:
    """Każdy język: te same klucze co angielski, te same {pola}, poprawne nazwy arkuszy."""
    import string

    def fields(text: str) -> set[str]:
        return {f for _, f, _, _ in string.Formatter().parse(text) if f}

    en, text = TEXT["en"], TEXT[lang]
    assert list(text) == list(en)
    for key, value in text.items():
        assert value.strip() and fields(value) == fields(en[key]), key
    for key in ("sheet.data", "sheet.dictionary", "sheet.about", "sheet.changes"):
        assert len(text[key]) <= 31 and not set(text[key]) & set("[]:*?/\\"), key
    assert "&P" in text["footer.page"] and "&N" in text["footer.page"]


@pytest.mark.parametrize("lang", LANGUAGES)
def test_export_in_every_language(tmp_path: Path, lang: str) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(language=lang), LOOK, sheets=ALL)
    wb = load_workbook(path)
    assert wb.sheetnames[:3] == [TEXT[lang][f"sheet.{s}"] for s in ("data", "dictionary", "about")]
    rows = [r[:] for r in parcels().rows]
    rows[0][3] = "usługi"
    result = export(str(tmp_path / "b.xlsx"), parcels(rows=rows, language=lang), LOOK, previous=path)
    assert [c.kind for c in result.changes] == ["modified"]


def test_theme_from_light_color_stays_readable() -> None:
    theme = theme_from_color("#FFE066")
    assert theme.header_font == "000000"
    assert theme.accent != "FFE066"


def test_only_data_sheet_by_default(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(), LOOK)
    wb = load_workbook(path)
    assert wb.sheetnames == ["Dane", META_SHEET]
    subtitle = wb["Dane"]["A2"].value
    assert "Słownik" not in subtitle and "Metryka" not in subtitle


def test_sheets_contain_only_chosen_items(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    sheets = Sheets(dictionary=True, dictionary_items=("kind",), about=True, about_items=("crs", "notes"))
    export(path, parcels(notes=["Uwaga testowa"]), LOOK, sheets=sheets)
    wb = load_workbook(path)
    assert [c.value for c in wb["Słownik"][HEADER_ROW] if c.value] == ["Kolumna", "Typ"]
    texts = [str(c.value) for row in wb["Metryka"].iter_rows() for c in row if c.value]
    assert "Układ współrzędnych" in texts and "EPSG:2180" in texts
    assert "Warstwa" not in texts and "Jak czytać ten plik" not in texts
    assert any("Uwaga testowa" in t for t in texts)


def test_plain_range_instead_of_table(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(), Look(as_table=False, autofilter=True))
    ws = load_workbook(path)["Dane"]
    assert not ws.tables
    assert ws.auto_filter.ref == f"A{HEADER_ROW}:G{HEADER_ROW + 2}"
    export(path, parcels(), Look(as_table=False, autofilter=False))
    assert not load_workbook(path)["Dane"].auto_filter.ref


def test_cell_look_and_borders(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    look = Look(
        header=CellLook(font="Arial", size=14, color="FF0000", italic=True, fill="000000", align="center"),
        body=CellLook(font="Verdana", size=9, bold=True, fill="FFFFEE", align="right", wrap=True),
        border_style="medium", border_color="00FF00", border_scope="horizontal",
        freeze_header=False, header_comments=False, gridlines=False, tab_color="123456",
    )
    export(path, parcels(), look)
    ws = load_workbook(path)["Dane"]
    head = ws.cell(row=HEADER_ROW, column=1)
    assert (head.font.name, head.font.sz, head.font.i, head.font.color.rgb[-6:]) == ("Arial", 14, True, "FF0000")
    assert head.fill.fgColor.rgb.endswith("000000") and head.alignment.horizontal == "center"
    assert head.comment is None
    body = ws.cell(row=HEADER_ROW + 1, column=2)
    assert (body.font.name, body.font.sz, body.font.b) == ("Verdana", 9, True)
    assert body.fill.fgColor.rgb.endswith("FFFFEE") and body.alignment.horizontal == "right"
    assert body.alignment.wrap_text
    assert body.border.top.style == "medium" and body.border.top.color.rgb.endswith("00FF00")
    assert body.border.left.style is None  # tylko linie poziome
    assert ws.freeze_panes is None and ws.sheet_view.showGridLines is False
    assert ws.sheet_properties.tabColor.rgb.endswith("123456")


def test_outline_border_only_on_edges(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(), Look(border_scope="outline"))
    ws = load_workbook(path)["Dane"]
    corner = ws.cell(row=HEADER_ROW, column=1)
    inner = ws.cell(row=HEADER_ROW + 1, column=3)
    last = ws.cell(row=HEADER_ROW + 2, column=7)
    assert corner.border.top.style == "thin" and corner.border.left.style == "thin"
    assert inner.border.left.style is None and inner.border.top.style is None
    assert last.border.bottom.style == "thin" and last.border.right.style == "thin"


def test_stripes_are_conditional_formatting(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(), Look(stripes=True, stripe_color="ABCDEF"))
    rules = [r for cf in load_workbook(path)["Dane"].conditional_formatting for r in cf.rules]
    assert len(rules) == 1 and rules[0].dxf.fill.fgColor.rgb.endswith("ABCDEF")
    export(path, parcels(), Look(stripes=False))
    assert not list(load_workbook(path)["Dane"].conditional_formatting)


def test_without_title_header_is_first_row_and_compare_works(tmp_path: Path) -> None:
    first = str(tmp_path / "v1.xlsx")
    export(first, parcels(), Look(show_title=False))
    ws = load_workbook(first)["Dane"]
    assert ws["A1"].value == "Nr działki" and ws["A2"].value == "12/3"
    assert ws.tables["HandoverData"].ref == "A1:G3"
    rows = [r[:] for r in parcels().rows]
    rows[0][3] = "usługi"
    result = export(str(tmp_path / "v2.xlsx"), parcels(rows=rows), LOOK, previous=first)
    assert [(c.kind, c.key) for c in result.changes] == [("modified", "12/3")]


def test_settings_roundtrip_and_tolerance() -> None:
    look = look_with_theme(Look(body=CellLook(size=9), border_scope="outline"), THEMES["forest"])
    assert from_dict(Look, json.loads(json.dumps(to_dict(look)))) == look
    assert look.header.fill == "2D5A3D" and look.body.size == 9
    broken = from_dict(Look, {"as_table": "tak", "header": {"size": 20, "nieznane": 1}, "obce": 2})
    assert broken.as_table is True and broken.header.size == 20 and broken.header.bold is True
    sheets = from_dict(Sheets, json.loads(json.dumps(to_dict(Sheets(about=True, about_items=("crs",))))))
    assert sheets == Sheets(about=True, about_items=("crs",))
    assert from_dict(Look, None) == Look()


def test_large_export_is_fast_enough(tmp_path: Path) -> None:
    """Pomiar w skali: 50 000 wierszy × 7 kolumn."""
    import time

    base = parcels().rows[0]
    rows = [[f"{i}/1", *base[1:]] for i in range(50_000)]
    start = time.perf_counter()
    export(str(tmp_path / "big.xlsx"), parcels(rows=rows), LOOK)
    elapsed = time.perf_counter() - start
    print(f"\n50 000 wierszy: {elapsed:.1f} s")
    assert elapsed < 60
