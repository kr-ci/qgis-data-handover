"""Test integracyjny w prawdziwym QGIS (uruchamiany przez scripts/qgis-test.sh).

Sprawdza odczyt warstwy (aliasy, value map, daty, NULL, dane osobowe),
trzy zakresy eksportu, porównanie wersji, ładowanie wtyczki i okno
(zrzuty zakładek trafiają do build/screens/).
"""

from __future__ import annotations

import datetime as dt
import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from qgis.core import (  # noqa: E402
    Qgis,
    QgsEditorWidgetSetup,
    QgsFeature,
    QgsGeometry,
    QgsPointXY,
    QgsProject,
    QgsRectangle,
    QgsSettings,
    QgsVectorLayer,
)
from qgis.PyQt.QtCore import QDate  # noqa: E402
from qgis.PyQt.QtWidgets import QApplication  # noqa: E402
from qgis.testing import start_app  # noqa: E402

app = start_app()

from openpyxl import load_workbook  # noqa: E402

from data_handover import classFactory, layer_reader  # noqa: E402
from data_handover.core.i18n import tr  # noqa: E402
from data_handover.core.model import HEADER_ROW, THEMES, Look, Sheets, look_with_theme  # noqa: E402
from data_handover.core.xlsx import export  # noqa: E402
from data_handover.dialog import HandoverDialog  # noqa: E402
from data_handover.layer_reader import FieldChoice  # noqa: E402
from data_handover.plugin import UI_LANGUAGE, ui_language  # noqa: E402

VERSION = Qgis.version().split("-")[0]
SCREENS = ROOT / "build" / "screens"
SCREENS.mkdir(parents=True, exist_ok=True)
OUT = Path(tempfile.mkdtemp())


def make_layer() -> QgsVectorLayer:
    uri = ("Point?crs=EPSG:2180&field=fid:integer&field=nr_dz:string&field=teryt:string"
           "&field=pow_m2:double(12,2)&field=przezn:string&field=data_akt:date"
           "&field=nazwisko:string&field=nr_ident:string")
    layer = QgsVectorLayer(uri, "dzialki_planowane", "memory")
    assert layer.isValid()
    layer.setFieldAlias(layer.fields().indexOf("nr_dz"), "Nr działki")
    layer.setFieldAlias(layer.fields().indexOf("pow_m2"), "Powierzchnia")
    layer.setFieldAlias(layer.fields().indexOf("przezn"), "Przeznaczenie")
    value_map = {"map": [{"zabudowa mieszkaniowa jednorodzinna": "MN"}, {"usługi": "U"}, {"lasy": "ZL"}]}
    layer.setEditorWidgetSetup(layer.fields().indexOf("przezn"), QgsEditorWidgetSetup("ValueMap", value_map))
    rows = [
        (1, "12/3", "0214011", 1523.456, "MN", QDate(2024, 5, 3), "Jan Kowalski", "44051401359", (100, 100)),
        (2, "12/4", "0214011", 980.1, "U", None, "Anna Nowak", "02070803628", (200, 200)),
        (3, "15/1", "0214012", None, "ZL", QDate(2023, 1, 9), None, "90090515836", (5000, 5000)),
    ]
    features = []
    for fid, nr, teryt, area, code, date, name, ident, (x, y) in rows:
        f = QgsFeature(layer.fields())
        f.setAttributes([fid, nr, teryt, area, code, date, name, ident])
        f.setGeometry(QgsGeometry.fromPointXY(QgsPointXY(x, y)))
        features.append(f)
    assert layer.dataProvider().addFeatures(features)[0]
    layer.updateExtents()
    QgsProject.instance().addMapLayer(layer)
    return layer


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)
    print(f"  ok  {message}")


def main() -> None:
    print(f"QGIS {Qgis.version()}")
    layer = make_layer()

    personal = layer_reader.personal_data(layer)
    check(set(personal) == {"nazwisko", "nr_ident"}, f"dane osobowe wykryte: {personal}")

    choices = [FieldChoice(n) for n in layer.fields().names()]
    for c in choices:
        if c.name == "nazwisko":
            c.export = False
        if c.name == "nr_ident":
            c.protect = "mask"
    pkg = layer_reader.build_package(layer, choices, key="nr_dz", personal=personal, title="Działki testowe",
                                     language="pl")
    by_name = {c.name: c for c in pkg.columns}
    check("nazwisko" not in by_name, "pominięte pole osobowe nie trafia do paczki")
    check(by_name["nr_ident"].masked, "pole PESEL oznaczone do maskowania")
    check(by_name["przezn"].represented and ("MN", "zabudowa mieszkaniowa jednorodzinna") in by_name["przezn"].codes,
          "value map: etykiety i słownik kodów")
    check(by_name["nr_dz"].header == "Nr działki", "alias jako nagłówek")
    check(by_name["fid"].is_identifier and by_name["teryt"].is_identifier, "identyfikatory rozpoznane")
    check(by_name["pow_m2"].kind == "decimal" and by_name["data_akt"].kind == "date", "typy liczb i dat")
    pos = [c.name for c in pkg.columns]
    first = pkg.rows[0]
    check(first[pos.index("przezn")] == "zabudowa mieszkaniowa jednorodzinna", "wartość przedstawiona etykietą")
    check(first[pos.index("data_akt")] == dt.date(2024, 5, 3), "QDate → date")
    check(pkg.rows[1][pos.index("data_akt")] is None and pkg.rows[2][pos.index("pow_m2")] is None, "NULL → None")
    check(any("Pominięto" in n for n in pkg.notes), "notatka o pominięciu danych osobowych")

    v1 = str(OUT / "v1.xlsx")
    export(v1, pkg, Look(), Sheets(dictionary=True, about=True))
    wb = load_workbook(v1)
    ws = wb["Dane"]
    check(ws.cell(row=HEADER_ROW + 1, column=pos.index("teryt") + 1).value == "0214011", "TERYT z zerem w pliku")
    check(ws.cell(row=HEADER_ROW + 1, column=pos.index("nr_ident") + 1).value == "***", "PESEL zamaskowany w pliku")
    facts = {wb["Metryka"].cell(row=r, column=1).value: wb["Metryka"].cell(row=r, column=2).value
             for r in range(HEADER_ROW, HEADER_ROW + 14)}
    check("EPSG:2180" in str(facts.get("Układ współrzędnych")), "CRS w metryce")
    check(facts.get("Liczba obiektów w pliku") == "3", "liczba obiektów w metryce")

    layer.selectByIds([1])
    sel = layer_reader.build_package(layer, choices, scope="selected", personal=personal)
    check(len(sel.rows) == 1, "zakres: zaznaczone")
    ext = layer_reader.build_package(layer, choices, scope="extent", extent=QgsRectangle(0, 0, 300, 300),
                                     extent_crs=layer.crs(), personal=personal)
    check(len(ext.rows) == 2, "zakres: zasięg mapy")

    layer.startEditing()
    feature = next(layer.getFeatures())
    layer.changeAttributeValue(feature.id(), layer.fields().indexOf("przezn"), "U")
    layer.commitChanges()
    pkg2 = layer_reader.build_package(layer, choices, key="nr_dz", personal=personal, title="Działki testowe")
    result = export(str(OUT / "v2.xlsx"), pkg2, look_with_theme(Look(), THEMES["forest"]), previous=v1)
    check([(c.kind, c.key, c.column) for c in result.changes] == [("modified", "12/3", "przezn")],
          f"porównanie wersji: {result.changes}")

    # Wtyczka ładuje się przez classFactory, a okno buduje się w tej wersji Qt.
    from qgis.testing.mocked import get_iface

    iface = get_iface()
    plugin = classFactory(iface)
    plugin.initGui()
    check(plugin.action is not None and plugin.layer_action is not None, "initGui: przycisk, menu i menu warstwy")
    plugin.unload()
    check(True, "unload bez błędów")
    dialog = HandoverDialog(None, None, layer=layer)
    dialog.resize(980, 680)
    dialog.show()
    QApplication.processEvents()
    check(dialog.look() == Look() and dialog.sheets() == Sheets(), "okno startuje z domyślnym wyglądem i arkuszami")
    check(dialog.lang_combo.currentData() == "en" and dialog.lang_combo.count() == 10,
          "język pliku: 10 do wyboru, domyślnie angielski")
    window_style = dialog.styleSheet()
    dialog._apply_preset(THEMES["forest"])
    check(dialog.look().header.fill == "2D5A3D" and dialog.look().body == Look().body,
          "gotowa kolorystyka zmienia kolory, zostawia czcionki")
    check(dialog.styleSheet() == window_style, "styl okna nie zależy od stylu pliku")
    dialog.dictionary_box.setChecked(True)
    dialog.lang_combo.setCurrentIndex(dialog.lang_combo.findData("ja"))
    dialog.reset_settings()
    check(dialog.look() == Look() and dialog.sheets() == Sheets() and dialog.lang_combo.currentData() == "en",
          "przycisk „Przywróć domyślne” zeruje wygląd, arkusze i język")
    dialog._set_look(Look())
    dialog.body_editor.fill_button.setToNull()
    check(dialog.look().body.fill == "", "tło zawartości: bez koloru")
    for index, name in enumerate(("dane", "wyglad", "arkusze", "zmiany")):
        dialog.tabs.setCurrentIndex(index)
        QApplication.processEvents()
        dialog.grab().save(str(SCREENS / f"okno_{name}_qgis{VERSION}.png"))
    check(dialog.ui_lang_combo.currentData() == "en" and dialog.ui_lang_combo.count() == 10
          and dialog.windowTitle() == tr("en", "ui.title"), "okno: 10 języków, domyślnie angielski")
    dialog.ui_lang_combo.setCurrentIndex(dialog.ui_lang_combo.findData("de"))
    check(dialog.relaunch and ui_language() == "de", "zmiana języka okna: zapis i ponowne otwarcie")
    german = HandoverDialog(None, None, layer=layer)
    check(german.windowTitle() == tr("de", "ui.title") and german.ui_lang_combo.currentData() == "de"
          and german.tabs.tabText(1) == tr("de", "ui.tab.look"), "okno po niemiecku")
    german.resize(980, 680)
    german.show()
    QApplication.processEvents()
    german.grab().save(str(SCREENS / f"okno_de_qgis{VERSION}.png"))
    german.close()
    QgsSettings().remove(UI_LANGUAGE)
    personal_rows = [r for r in range(dialog.table.rowCount())
                     if dialog.table.item(r, 0).text() in ("nazwisko", "nr_ident")]
    check(all(dialog.table.item(r, 0).checkState() != dialog.table.item(0, 0).checkState()
              or dialog.table.item(r, 0).text() == "fid" for r in personal_rows), "pola osobowe domyślnie odznaczone")
    check(len(dialog.choices()) == len(layer.fields()), "okno pokazuje wszystkie pola")
    dialog.close()

    # Przegląd R4, zarzut 3: filtr odwołujący się do pominiętego pola osobowego nie trafia do metryki.
    layer.setSubsetString("\"nazwisko\" = 'Jan Kowalski'")
    hidden = layer_reader.build_package(layer, choices, personal=personal, language="pl")
    check(dict(hidden.facts)["filter"] == tr("pl", "filter.hidden"), "filtr z danymi osobowymi ukryty w metryce")
    layer.setSubsetString("")

    # Zarzuty 9 i 12: podstawa porównania zmienia się tylko po udanym zapisie, a uwagi widać w komunikacie.
    import data_handover.dialog as dialog_module

    layer.setCustomProperty(dialog_module.PREVIOUS, "")
    box = HandoverDialog(iface, None, layer=layer)
    warnings: list[str] = []
    box._warn = warnings.append  # type: ignore[method-assign]
    target = {"path": str(OUT / "brak_katalogu" / "plik.xlsx")}
    dialog_module.QFileDialog.getSaveFileName = lambda *args, **kwargs: (target["path"], "")
    box.export()
    check(bool(warnings) and layer.customProperty(dialog_module.PREVIOUS, "") == "",
          "nieudany zapis nie zmienia podstawy porównania")
    target["path"] = str(OUT / "v3.xlsx")
    box = HandoverDialog(iface, None, layer=layer)
    box.export()
    check(layer.customProperty(dialog_module.PREVIOUS, "") == target["path"],
          "udany zapis staje się podstawą porównania")
    bar = iface.messageBar()
    text = bar.createMessage.call_args[0][1]
    level = bar.pushWidget.call_args[0][1]
    check("•" in text and level == Qgis.MessageLevel.Warning, "uwagi z eksportu widoczne w komunikacie")

    # Plik do obejrzenia (zrzut arkuszy robi scripts/render-xlsx.sh).
    export(str(ROOT / "build" / f"przyklad_qgis{VERSION}.xlsx"), pkg2, Look(),
           Sheets(dictionary=True, about=True), previous=v1)
    print("OK")


if __name__ == "__main__":
    main()
    sys.stdout.flush()
    os._exit(0)  # bez sprzątania QGIS przy wyjściu (znane awarie przy exit w trybie testowym)
