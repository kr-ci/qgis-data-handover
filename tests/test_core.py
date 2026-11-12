"""Testy rdzenia: zapis XLSX, porównanie wersji, prywatność. Bez QGIS."""

from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from data_handover.core import diff, privacy, xlsx  # noqa: E402
from data_handover.core.i18n import LANGUAGES, TEXT, count_label, tr  # noqa: E402
from data_handover.core.model import (  # noqa: E402
    HEADER_ROW,
    THEMES,
    CellLook,
    Column,
    GeoColumns,
    Look,
    Package,
    Sheets,
    Summary,
    from_dict,
    look_with_theme,
    map_url,
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
    for key in ("sheet.data", "sheet.dictionary", "sheet.about", "sheet.changes", "sheet.summary"):
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
    geo = GeoColumns(measure=True, area_unit="m2", link=True, map="geoportal")
    assert from_dict(GeoColumns, json.loads(json.dumps(to_dict(geo)))) == geo
    bad = from_dict(GeoColumns, {"measure": True, "area_unit": "akr", "length_unit": 5, "map": "google"})
    assert bad == GeoColumns(measure=True)


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


# --- Przegląd R4: każda naprawiona wada ma swój test ---------------------------


def _all_text(path: str) -> str:
    """Cała treść pliku widoczna dla odbiorcy: komórki i komentarze ze wszystkich arkuszy."""
    wb = load_workbook(path)
    parts = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                parts += [str(cell.value or ""), cell.comment.text if cell.comment else ""]
    return "\n".join(parts)


def test_masking_hides_plain_values_from_previous_export(tmp_path: Path) -> None:
    """Zarzut 1: poprzedni eksport był jawny, nowy maskuje — stare nazwisko nie wraca."""
    first = str(tmp_path / "v1.xlsx")
    plain = parcels()
    plain.columns[5].masked = False
    export(first, plain, LOOK)
    second = str(tmp_path / "v2.xlsx")
    result = export(second, parcels(), LOOK, sheets=ALL, previous=first)
    assert not [c for c in result.changes if c.column == "wlasc"]
    assert "Jan Kowalski" not in _all_text(second)


def test_masked_key_cannot_be_compared(tmp_path: Path) -> None:
    first = str(tmp_path / "v1.xlsx")
    export(first, parcels(), LOOK)
    pkg = parcels()
    pkg.columns[0].masked = True
    with pytest.raises(PackageError, match="bez maskowania"):
        export(str(tmp_path / "v2.xlsx"), pkg, LOOK, previous=first)


def test_masked_column_hides_value_map_labels(tmp_path: Path) -> None:
    """Zarzut 2: etykiety value map kolumny maskowanej nie trafiają do dymka ani słownika."""
    path = str(tmp_path / "a.xlsx")
    pkg = parcels()
    pkg.columns[5].codes = [("1", "Jan Kowalski")]
    export(path, pkg, LOOK, sheets=ALL)
    assert "Jan Kowalski" not in _all_text(path)


def test_filter_mentioning_personal_field() -> None:
    """Zarzut 3: filtr odwołujący się do chronionego pola jest rozpoznawany."""
    assert privacy.mentions("\"email\" = 'anna@example.com'", ["email"])
    assert privacy.mentions("EMAIL LIKE '%x%'", ["email"])
    assert not privacy.mentions("\"email_ok\" = 1", ["email"])
    assert not privacy.mentions("", ["email"])


def test_long_integers_stay_exact(tmp_path: Path) -> None:
    """Zarzut 4: liczba całkowita ponad 15 cyfr trafia jako tekst, klucze się nie zlewają."""
    assert diff.key_text(9007199254740992) != diff.key_text(9007199254740993)
    assert diff.normalize(5) == diff.normalize(5.0)
    columns = [Column("id", "ID", "integer", is_identifier=True), Column("v", "V", "integer")]
    rows: list[list[object]] = [[9007199254740993, 1], [9007199254740992, 2], [12, 3]]
    pkg = Package("t", columns, rows, key="id", language="pl")  # type: ignore[arg-type]
    first = str(tmp_path / "v1.xlsx")
    export(first, pkg, LOOK)
    ws = load_workbook(first)["Dane"]
    assert ws.cell(row=HEADER_ROW + 1, column=1).value == "9007199254740993"
    assert ws.cell(row=HEADER_ROW + 3, column=1).value == 12
    result = export(str(tmp_path / "v2.xlsx"), pkg, LOOK, previous=first)
    assert result.changes == []


def test_edited_previous_file_is_rejected(tmp_path: Path) -> None:
    """Zarzut 5: dopisany wiersz albo przestawione kolumny — błąd zamiast cichego raportu."""
    first = str(tmp_path / "v1.xlsx")
    export(first, parcels(), LOOK)
    appended = str(tmp_path / "dopisany.xlsx")
    wb = load_workbook(first)
    wb["Dane"].cell(row=HEADER_ROW + 3, column=1).value = "99/9"
    wb.save(appended)
    swapped = str(tmp_path / "przestawiony.xlsx")
    wb = load_workbook(first)
    ws = wb["Dane"]
    a, b = ws.cell(row=HEADER_ROW, column=1), ws.cell(row=HEADER_ROW, column=2)
    a.value, b.value = b.value, a.value
    wb.save(swapped)
    for edited in (appended, swapped):
        with pytest.raises(PackageError, match="zmieniono po eksporcie"):
            export(str(tmp_path / "v2.xlsx"), parcels(), LOOK, previous=edited)


def test_key_left_out_of_export_is_an_error(tmp_path: Path) -> None:
    """Zarzut 6: porównanie bez klucza nie może udawać „brak różnic”."""
    first = str(tmp_path / "v1.xlsx")
    export(first, parcels(), LOOK)
    with pytest.raises(PackageError, match="kolumny-klucza"):
        export(str(tmp_path / "v2.xlsx"), parcels(key=None), LOOK, previous=first)


def test_generated_headers_are_unique(tmp_path: Path) -> None:
    """Zarzut 7."""
    columns = [Column("a", "X"), Column("b", "X"), Column("c", "X (2)"), Column("d", "x")]
    path = str(tmp_path / "a.xlsx")
    export(path, Package("t", columns, [["1", "2", "3", "4"]], language="pl"), LOOK)
    headers = [str(c.value).lower() for c in load_workbook(path)["Dane"][HEADER_ROW]]
    assert len(set(headers)) == 4


def test_title_starting_with_equals_stays_text(tmp_path: Path) -> None:
    """Zarzut 8."""
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(title="=1+1"), LOOK)
    cell = load_workbook(path)["Dane"]["A1"]
    assert cell.value == "=1+1" and cell.data_type == "s"


def test_illegal_characters_outside_data_are_cleaned(tmp_path: Path) -> None:
    """Zarzut 10: znak sterujący w nagłówku, opisie, tytule i metryce nie przerywa eksportu."""
    path = str(tmp_path / "a.xlsx")
    pkg = parcels(title="Dział\x01ki", facts=[("note", "uwaga\x02")])
    pkg.columns[0].header = "Nr\x01 działki"
    pkg.columns[2].description = "opis\x03"
    result = export(path, pkg, LOOK, sheets=ALL)
    text = _all_text(path)
    assert "Nr działki" in text and "Działki" in text and "\x01" not in text
    assert any("4 nagłówkach" in n for n in result.notes)


def test_failed_save_keeps_previous_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Zarzut 11: błąd w trakcie zapisu nie niszczy poprzedniego pliku i nie zostawia śmieci."""
    from openpyxl import Workbook

    path = tmp_path / "a.xlsx"
    export(str(path), parcels(), LOOK)
    before = path.read_bytes()

    def broken_save(self: object, filename: str) -> None:
        Path(filename).write_bytes(b"polowa")
        raise OSError("brak miejsca na dysku")

    monkeypatch.setattr(Workbook, "save", broken_save)
    with pytest.raises(OSError):
        export(str(path), parcels(), LOOK)
    assert path.read_bytes() == before
    assert [p.name for p in tmp_path.iterdir()] == ["a.xlsx"]


def test_too_many_changes_is_a_clear_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Zarzut 13: więcej zmian niż wierszy arkusza — czytelny błąd przed zapisem."""
    from data_handover.core import xlsx

    first = str(tmp_path / "v1.xlsx")
    export(first, parcels(), LOOK)
    rows = [[r[0], "x", 1.0, "x", None, None, False] for r in parcels().rows]
    monkeypatch.setattr(xlsx, "MAX_ROWS", 3)
    with pytest.raises(PackageError, match="Porównanie znalazło"):
        export(str(tmp_path / "v2.xlsx"), parcels(rows=rows), LOOK, previous=first)


# --- R5: eksport w tle — postęp i anulowanie ----------------------------------


def test_export_reports_progress(tmp_path: Path) -> None:
    rows = [[f"{i}/1", *parcels().rows[0][1:]] for i in range(500)]
    seen: list[float] = []
    export(str(tmp_path / "p.xlsx"), parcels(rows=rows), LOOK, progress=seen.append)
    assert seen and seen == sorted(seen) and 0 <= seen[0] and seen[-1] == 100


@pytest.mark.parametrize("when", ["rows", "save"])
def test_canceled_export_keeps_previous_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, when: str) -> None:
    """Anulowanie nie zostawia uszkodzonego ani częściowego pliku — poprzednia wersja zostaje."""
    from data_handover.core import xlsx

    target = tmp_path / "plik.xlsx"
    target.write_bytes(b"poprzednia wersja")
    state = {"cancel": when == "rows"}
    real_save = xlsx.Workbook.save

    def save_then_cancel(wb: object, path: str) -> None:
        real_save(wb, path)
        state["cancel"] = True  # użytkownik anuluje, gdy plik tymczasowy jest już zapisany

    monkeypatch.setattr(xlsx.Workbook, "save", save_then_cancel)
    with pytest.raises(xlsx.CanceledError):
        export(str(target), parcels(), LOOK, canceled=lambda: state["cancel"])
    assert target.read_bytes() == b"poprzednia wersja"
    assert [p.name for p in tmp_path.iterdir()] == ["plik.xlsx"]


# --- R7: kolumny z geometrii -------------------------------------------------

OSM = "https://www.openstreetmap.org/?mlat=52.245539&mlon=21.014835#map=18/52.245539/21.014835"


def test_map_url() -> None:
    assert map_url("osm", 52.2455391, 21.0148349) == OSM
    assert map_url("geoportal", 52.245539, 21.014835) == (
        "https://mapy.geoportal.gov.pl/imapnext/imap/?mapview=52.245539,21.014835,1000s")


def geo_package(links: list[object], area: list[object] | None = None) -> Package:
    columns = [Column("nr_dz", "Nr działki", "text", is_identifier=True),
               Column("@area_ha", "Powierzchnia", "decimal", unit="ha", decimals=4),
               Column("@link", "Mapa", "link")]
    areas = area if area is not None else [0.0512] * len(links)
    rows = [[f"{i}/1", a, link] for i, (a, link) in enumerate(zip(areas, links), start=1)]
    return Package(title="Działki", columns=columns, rows=rows, language="pl", key="nr_dz")  # type: ignore[arg-type]


def test_link_column_is_clickable_text(tmp_path: Path) -> None:
    path = str(tmp_path / "geo.xlsx")
    export(path, geo_package([OSM, None]), LOOK)
    ws = load_workbook(path)["Dane"]
    cell = ws.cell(row=HEADER_ROW + 1, column=3)
    assert cell.value == tr("pl", "geo.link.label")
    assert cell.hyperlink is not None and cell.hyperlink.target == OSM
    assert cell.font.underline == "single" and cell.font.color.rgb.endswith(xlsx.LINK_COLOR)
    empty = ws.cell(row=HEADER_ROW + 2, column=3)
    assert empty.value is None and empty.hyperlink is None
    assert ws.cell(row=HEADER_ROW + 1, column=2).number_format.endswith('"ha"')


def test_link_column_not_compared(tmp_path: Path) -> None:
    first = str(tmp_path / "v1.xlsx")
    export(first, geo_package([OSM, OSM]), LOOK)
    other = OSM.replace("21.014835", "21.014900")
    result = export(str(tmp_path / "v2.xlsx"), geo_package([other, other]), LOOK, previous=first)
    assert result.changes == []


def test_rounded_measure_roundtrip_has_no_changes(tmp_path: Path) -> None:
    first = str(tmp_path / "v1.xlsx")
    export(first, geo_package([OSM, OSM], area=[0.0512, 12.3457]), LOOK)
    result = export(str(tmp_path / "v2.xlsx"), geo_package([OSM, OSM], area=[0.0512, 12.3457]), LOOK,
                    previous=first)
    assert result.changes == []


def test_too_many_links_are_dropped(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(xlsx, "HYPERLINK_LIMIT", 2)
    path = str(tmp_path / "many.xlsx")
    result = export(path, geo_package([OSM, OSM, OSM]), LOOK, sheets=ALL)
    wb = load_workbook(path)
    assert [c.value for c in wb["Dane"][HEADER_ROW]] == ["Nr działki", "Powierzchnia"]
    fields = [r[1] for r in wb["Słownik"].iter_rows(min_row=HEADER_ROW + 1, values_only=True)]
    assert fields == ["nr_dz", tr("pl", "dict.geometry_source")]
    assert [r[1] for r in wb[META_SHEET].iter_rows(values_only=True) if r[0] == "column"] == ["nr_dz", "@area_ha"]
    assert tr("pl", "note.link.dropped", limit="2") in result.notes


def test_year_column_without_thousands_separator(tmp_path: Path) -> None:
    def fmt(column: Column, values: list[object]) -> str:
        pkg = parcels(rows=[[v] for v in values])
        pkg.columns = [column]
        path = str(tmp_path / "rok.xlsx")
        export(path, pkg, LOOK)
        return str(load_workbook(path).worksheets[0].cell(HEADER_ROW + 1, 1).number_format)

    assert fmt(Column("rok", "Rok nabycia", "integer"), [2019, None, 2021]) == "0"
    assert fmt(Column("rok", "Rok nabycia", "integer"), [2019, 15000]) == "#,##0"
    assert fmt(Column("rok", "Rok nabycia", "integer", unit="szt."), [2019, 2021]) == '#,##0 "szt."'
    assert fmt(Column("id", "Id", "integer", is_identifier=True), [2019, 2021]) == "0"


def test_summary_defaults_and_from_dict() -> None:
    assert from_dict(Summary, None) == Summary()
    got = from_dict(Summary, {"row": True, "group_by": 5})
    assert got.row is True and got.group_by == "" and got.sheet is False and got.measure is True
    assert Package("t", [], []).summary == Summary()


def test_display_without_grouping() -> None:
    assert xlsx.display(2019, Column("rok", "Rok", "integer", grouping=False), "pl") == "2019"
    assert xlsx.display(2019, Column("rok", "Rok", "integer"), "pl") != "2019"


# --- wiersz sum (R8, zadanie 2) ---------------------------------------------


def _totals_pkg(**kw: object) -> Package:
    pkg = parcels(summary=Summary(row=True), **kw)
    next(c for c in pkg.columns if c.name == "pow_m2").summed = True
    return pkg


def test_totals_row_plain_range(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, _totals_pkg(), Look(as_table=False))
    ws = load_workbook(path)["Dane"]
    assert ws.cell(row=HEADER_ROW + 3, column=3).value == "=SUBTOTAL(109,C5:C6)"
    label = ws.cell(row=HEADER_ROW + 3, column=1)
    assert label.value == tr("pl", "totals.label") and label.font.bold
    assert ws.auto_filter.ref == "A4:G6"


def test_totals_row_excel_table(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, _totals_pkg(), Look())
    ws = load_workbook(path)["Dane"]
    table = ws.tables["HandoverData"]
    assert table.ref == "A4:G7" and table.totalsRowCount == 1
    assert table.autoFilter.ref == "A4:G6"
    cols = {c.name: c for c in table.tableColumns}
    assert list(cols) == [c.value for c in ws[HEADER_ROW]]
    assert cols["Powierzchnia"].totalsRowFunction == "sum"
    assert cols["Nr działki"].totalsRowLabel == tr("pl", "totals.label")
    assert ws.cell(row=HEADER_ROW + 3, column=3).value == "=SUBTOTAL(109,HandoverData[[Powierzchnia]])"


def test_totals_structured_reference_escapes(tmp_path: Path) -> None:
    assert xlsx._structured("Pow. [m²] #1 's") == "Pow. '[m²'] '#1 ''s"
    for header, expected in (
        ("Pow. [m²] #1 's", "=SUBTOTAL(109,HandoverData[[Pow. '[m²'] '#1 ''s]])"),
        ("Pow.\nm²", "=SUBTOTAL(109,C5:C6)"),
    ):
        pkg = _totals_pkg()
        next(c for c in pkg.columns if c.name == "pow_m2").header = header
        path = str(tmp_path / "a.xlsx")
        export(path, pkg, Look())
        assert load_workbook(path)["Dane"].cell(row=HEADER_ROW + 3, column=3).value == expected


def test_totals_structured_reference_escapes_backtick() -> None:
    assert xlsx._structured("a`b") == "a'`b"


def test_deleted_data_row_in_file_with_totals_row_is_rejected(tmp_path: Path) -> None:
    """Usunięty wiersz danych: wiersz sum podjeżdża w górę i nie może być wzięty za dane."""
    labelled = _totals_pkg(rows=[["12/3", "0214011", 1.0, "usługi", dt.date(2024, 5, 3), None, True],
                                 ["12/4", "0214011", 2.0, "usługi", dt.date(2024, 5, 3), None, False],
                                 ["12/5", "0214011", 3.0, "usługi", dt.date(2024, 5, 3), None, False]])
    summed_first = parcels(summary=Summary(row=True), columns=[Column("n", "Liczba", "integer", summed=True)],
                           rows=[[1], [2], [3]], key=None)
    for i, pkg in enumerate((labelled, summed_first)):
        first = str(tmp_path / f"v1_{i}.xlsx")
        export(first, pkg, Look())
        assert len(read_snapshot(first).rows) == 3
        wb = load_workbook(first)
        wb["Dane"].delete_rows(HEADER_ROW + 3)
        edited = str(tmp_path / f"usuniety_{i}.xlsx")
        wb.save(edited)
        with pytest.raises(PackageError, match="zmieniono po eksporcie"):
            export(str(tmp_path / f"v2_{i}.xlsx"), pkg, Look(), previous=edited)


def test_masked_summed_column_is_not_totalled(tmp_path: Path) -> None:
    pkg = parcels(summary=Summary(row=True, sheet=True, group_by="przezn"),
                  rows=[["12/3", "0214011", 1.0, "usługi", dt.date(2024, 5, 3), "Jan Kowalski", True],
                        ["12/4", "0214011", 2.0, "usługi", dt.date(2024, 5, 3), "Anna Nowak", False]])
    next(c for c in pkg.columns if c.name == "wlasc").summed = True
    path = str(tmp_path / "a.xlsx")
    export(path, pkg, Look())
    wb = load_workbook(path)
    ws = wb["Dane"]
    assert ws.cell(row=HEADER_ROW + 3, column=6).value is None
    summary = wb[tr("pl", "sheet.summary")]
    assert all(tr("pl", "summary.sum", header="Właściciel") != c.value for row in summary.iter_rows() for c in row)
    text = _all_text(path)
    assert "Kowalski" not in text and "Nowak" not in text


def test_no_totals_row_without_rows_or_summed_columns(tmp_path: Path) -> None:
    cases = [
        _totals_pkg(rows=[]),
        parcels(summary=Summary(row=True)),  # żadna kolumna nie jest sumowana
        parcels(),  # row=False
    ]
    next(c for c in cases[2].columns if c.name == "pow_m2").summed = True
    for i, pkg in enumerate(cases):
        path = str(tmp_path / f"{i}.xlsx")
        export(path, pkg, Look())
        wb = load_workbook(path)
        ws = wb["Dane"]
        assert not ws.tables["HandoverData"].totalsRowCount if ws.tables else True
        assert all(
            not (isinstance(c.value, str) and c.value.startswith("=")) for row in ws.iter_rows() for c in row
        )
        meta = {r[0].value: r[1].value for r in wb[META_SHEET].iter_rows() if r[0].value}
        assert meta["totals_row"] == 0


def test_totals_row_first_column_summed_has_no_label(tmp_path: Path) -> None:
    pkg = parcels(summary=Summary(row=True), columns=[Column("n", "Liczba", "integer", summed=True)],
                  rows=[[1], [2]], key=None)
    path = str(tmp_path / "a.xlsx")
    export(path, pkg, Look())
    ws = load_workbook(path)["Dane"]
    assert ws.cell(row=HEADER_ROW + 3, column=1).value == "=SUBTOTAL(109,HandoverData[[Liczba]])"
    assert ws.tables["HandoverData"].tableColumns[0].totalsRowLabel is None


def test_totals_row_survives_comparison(tmp_path: Path) -> None:
    first = str(tmp_path / "v1.xlsx")
    export(first, _totals_pkg(), Look())
    assert len(read_snapshot(first).rows) == 2
    rows = [["12/3", "0214011", 1.0, "usługi", dt.date(2024, 5, 3), None, True]]
    result = export(str(tmp_path / "v2.xlsx"), _totals_pkg(rows=rows), Look(), previous=first)
    assert {(c.kind, c.key) for c in result.changes} == {("modified", "12/3"), ("removed", "12/4")}
    wb = load_workbook(first)
    wb["Dane"].cell(row=HEADER_ROW + 4, column=1).value = "dopisek"
    edited = str(tmp_path / "edited.xlsx")
    wb.save(edited)
    with pytest.raises(PackageError, match="zmieniono po eksporcie"):
        export(str(tmp_path / "v3.xlsx"), _totals_pkg(), Look(), previous=edited)


# --- arkusz „Podsumowanie” (R8, zadanie 3) -------------------------------------


def _summary_pkg(group_by: str = "przezn", **kw: object) -> Package:
    pkg = parcels(summary=Summary(sheet=True, group_by=group_by), **kw)
    next(c for c in pkg.columns if c.name == "pow_m2").summed = True
    return pkg


def test_summarize_groups_sums_gaps() -> None:
    cols = [
        Column("gm", "Gmina", "text"),
        Column("ha", "Ha", "decimal", summed=True),
        Column("sekret", "Sekret", "text", masked=True),
        Column("pelna", "Pełna", "text"),
    ]
    gm = ["Kórnik", "Mosina", "Kórnik", None, "Mosina"]
    ha = [1.5, 2.0, None, 0.25, "x"]
    rows = [[g, h, None, "a"] for g, h in zip(gm, ha)]
    pkg = Package("T", cols, rows, language="pl", summary=Summary(sheet=True, group_by="gm"))
    data = xlsx.summarize(pkg, rows)
    assert data.count == 5
    assert [(c.name, v) for c, v in data.sums] == [("ha", pytest.approx(3.75))]
    assert data.groups is not None
    col, groups = data.groups
    assert col.name == "gm"
    assert [(g, n) for g, n, _ in groups] == [("Kórnik", 2), ("Mosina", 2), (None, 1)]
    assert [s[0] for _, _, s in groups] == [1.5, 2.0, 0.25]
    assert sum(s[0] for _, _, s in groups) == pytest.approx(data.sums[0][1])
    assert data.gaps == [("Gmina", 1), ("Ha", 1)]
    assert data.masked == ["Sekret"]


def test_summary_sheet_layout(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, _summary_pkg(), LOOK, sheets=ALL)
    wb = load_workbook(path)
    assert wb.sheetnames[:2] == ["Dane", "Podsumowanie"]
    ws = wb["Podsumowanie"]
    cells = {(c.row, c.column): c for row in ws.iter_rows() for c in row if c.value is not None}
    texts = {c.value for c in cells.values() if isinstance(c.value, str)}
    assert {tr("pl", "summary.whole"), tr("pl", "summary.count"), tr("pl", "summary.total"),
            tr("pl", "summary.gaps"), tr("pl", "summary.by", header="Przeznaczenie"),
            tr("pl", "summary.sum", header="Powierzchnia")} <= texts
    assert ws["A1"].value == tr("pl", "sheet.summary")
    assert "2 obiekty" in ws["A2"].value
    count = next(c for c in cells.values() if c.value == 2 and c.number_format == "#,##0")
    assert count.data_type == "n"
    # etykieta value map, nie kod; sumy jako liczby z formatem kolumny
    assert "zabudowa mieszkaniowa" in texts and "usługi" in texts
    total = next(c for c in cells.values() if c.value == tr("pl", "summary.total"))
    assert total.font.bold and total.border.top.style == "medium"
    area = [c for c in cells.values() if isinstance(c.value, float)]
    assert area and all(c.number_format == '#,##0.00 "m²"' for c in area)
    assert sum(c.value for c in area if c.row == total.row) == pytest.approx(1523.456 + 980.1)
    assert "Podsumowanie" in _all_text(path)


def test_summary_never_shows_masked_column(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, _summary_pkg("wlasc"), LOOK, sheets=ALL)
    ws = load_workbook(path)["Podsumowanie"]
    assert not any(isinstance(c.value, str) and c.value.startswith("Według") for row in ws.iter_rows() for c in row)
    text = "\n".join(str(c.value) for row in ws.iter_rows() for c in row if c.value is not None)
    assert "Kowalski" not in _all_text(path) and "J*** K***" not in text
    assert text.count("Właściciel") == 1 and "pominięte" in text
    assert tr("pl", "note.summary.group_missing", field="wlasc") not in _all_text(path)


def test_summary_group_column_missing(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    result = export(path, _summary_pkg("nieistnieje"), LOOK, sheets=ALL)
    note = tr("pl", "note.summary.group_missing", field="nieistnieje")
    assert note in result.notes and note in _all_text(path)
    ws = load_workbook(path)["Podsumowanie"]
    assert not any(isinstance(c.value, str) and c.value.startswith("Według") for row in ws.iter_rows() for c in row)


def test_summary_group_by_date_and_bool(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, _summary_pkg("data_akt"), LOOK)
    ws = load_workbook(path)["Podsumowanie"]
    dates = [c for row in ws.iter_rows() for c in row if c.value == dt.datetime(2024, 5, 3)]
    assert dates and dates[0].number_format == "DD.MM.YYYY"
    export(path, _summary_pkg("aktywna"), LOOK)
    ws = load_workbook(path)["Podsumowanie"]
    assert {"tak", "nie"} <= {c.value for row in ws.iter_rows() for c in row}


def test_summary_empty_package(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, _summary_pkg(rows=[]), LOOK)
    ws = load_workbook(path)["Podsumowanie"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert 0 in values and tr("pl", "summary.count") in values
    assert not any(isinstance(v, str) and (v.startswith("Według") or v == tr("pl", "summary.gaps")) for v in values)


def test_summary_off_by_default_and_gap_sentences(tmp_path: Path) -> None:
    path = str(tmp_path / "a.xlsx")
    export(path, parcels(), LOOK, sheets=ALL)
    assert load_workbook(path).sheetnames == ["Dane", "Słownik", "Metryka", META_SHEET]
    full = [["1", "0", 1.0, "a", dt.date(2024, 1, 1), None, True]]
    export(path, _summary_pkg(rows=full), LOOK)
    values = [c.value for row in load_workbook(path)["Podsumowanie"].iter_rows() for c in row]
    assert any(isinstance(v, str) and tr("pl", "summary.gaps.none") in v for v in values)
