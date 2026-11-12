"""Test integracyjny w prawdziwym QGIS (uruchamiany przez scripts/qgis-test.sh).

Sprawdza odczyt warstwy (aliasy, value map, daty, NULL, dane osobowe),
trzy zakresy eksportu, porównanie wersji, ładowanie wtyczki i okno
(zrzuty zakładek trafiają do build/screens/) oraz algorytm Processing
(także przez `qgis_process`).
"""

from __future__ import annotations

import datetime as dt
import glob
import json
import math
import os
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.append("/usr/share/qgis/python/plugins")  # moduł processing

from qgis.core import (  # noqa: E402
    Qgis,
    QgsAggregateCalculator,
    QgsApplication,
    QgsEditorWidgetSetup,
    QgsFeature,
    QgsField,
    QgsGeometry,
    QgsPointXY,
    QgsProcessingAlgRunnerTask,
    QgsProcessingContext,
    QgsProcessingFeedback,
    QgsProject,
    QgsRectangle,
    QgsSettings,
    QgsVectorFileWriter,
    QgsVectorLayer,
)
from qgis.PyQt.QtCore import QDate, QMetaType, QSettings, Qt, QTimer  # noqa: E402
from qgis.PyQt.QtWidgets import QApplication  # noqa: E402
from qgis.testing import start_app  # noqa: E402

app = start_app()

from openpyxl import load_workbook  # noqa: E402

from data_handover import classFactory, layer_reader  # noqa: E402
from data_handover.core.i18n import LANGUAGES, tr  # noqa: E402
from data_handover.core.model import (  # noqa: E402
    HEADER_ROW,
    THEMES,
    GeoColumns,
    Look,
    Sheets,
    Summary,
    look_with_theme,
    map_url,
    to_dict,
)
from data_handover.core.xlsx import export  # noqa: E402
from data_handover.dialog import HandoverDialog  # noqa: E402
from data_handover.export_task import ExportTask  # noqa: E402
from data_handover.export_task import start as start_task  # noqa: E402
from data_handover.layer_reader import FieldChoice  # noqa: E402
from data_handover.plugin import UI_LANGUAGE, ui_language  # noqa: E402

VERSION = Qgis.version().split("-")[0]
SCREENS = ROOT / "build" / "screens"
SCREENS.mkdir(parents=True, exist_ok=True)
OUT = Path(tempfile.mkdtemp())


def make_layer() -> QgsVectorLayer:
    uri = ("Point?crs=EPSG:2180&field=fid:integer&field=nr_dz:string&field=teryt:string"
           "&field=pow_m2:double(12,2)&field=przezn:string&field=data_akt:date"
           "&field=nazwisko:string&field=nr_ident:string&field=uwagi:string")
    layer = QgsVectorLayer(uri, "dzialki_planowane", "memory")
    assert layer.isValid()
    layer.setFieldAlias(layer.fields().indexOf("nr_dz"), "Nr działki")
    layer.setFieldAlias(layer.fields().indexOf("pow_m2"), "Powierzchnia")
    layer.setFieldAlias(layer.fields().indexOf("przezn"), "Przeznaczenie")
    value_map = {"map": [{"zabudowa mieszkaniowa jednorodzinna": "MN"}, {"usługi": "U"}, {"lasy": "ZL"}]}
    layer.setEditorWidgetSetup(layer.fields().indexOf("przezn"), QgsEditorWidgetSetup("ValueMap", value_map))
    rows = [
        (1, "12/3", "0214011", 1523.456, "MN", QDate(2024, 5, 3), "Jan Kowalski", "44051401359",
         "oględziny 2024-05-03 10:00", (100, 100)),
        (2, "12/4", "0214011", 980.1, "U", None, "Anna Nowak", "02070803628", None, (200, 200)),
        (3, "15/1", "0214012", None, "ZL", QDate(2023, 1, 9), None, "90090515836", "03.02.2023 08:15", (5000, 5000)),
    ]
    features = []
    for fid, nr, teryt, area, code, date, name, ident, note, (x, y) in rows:
        f = QgsFeature(layer.fields())
        f.setAttributes([fid, nr, teryt, area, code, date, name, ident, note])
        f.setGeometry(QgsGeometry.fromPointXY(QgsPointXY(x, y)))
        features.append(f)
    assert layer.dataProvider().addFeatures(features)[0]
    layer.updateExtents()
    QgsProject.instance().addMapLayer(layer)
    return layer


def wait(task: ExportTask, seconds: float = 120) -> None:
    """Czeka na koniec zadania, obsługując w tym czasie zdarzenia jak okno QGIS."""
    deadline = time.monotonic() + seconds
    while not task.done:
        if time.monotonic() > deadline:
            raise AssertionError("zadanie w tle nie skończyło się na czas")
        QApplication.processEvents()
        time.sleep(0.005)


def background(small: QgsVectorLayer) -> None:
    """R5: eksport w tle nie zamraża pętli zdarzeń, a anulowanie nie zostawia pliku."""
    uri = "Point?crs=EPSG:2180&field=fid:integer&field=nr_dz:string&field=pow_m2:double&field=przezn:string"
    layer = QgsVectorLayer(uri, "duza", "memory")
    layer.setEditorWidgetSetup(3, small.editorWidgetSetup(small.fields().indexOf("przezn")))  # value map
    features = []
    for i in range(30_000):
        f = QgsFeature(layer.fields())
        f.setAttributes([i, f"{i}/1", 1000.0 + i, "U"])
        features.append(f)
    assert layer.dataProvider().addFeatures(features)[0]
    choices = [FieldChoice(n) for n in layer.fields().names()]

    ticks: list[float] = []
    timer = QTimer()
    timer.timeout.connect(lambda: ticks.append(time.monotonic()))
    timer.start(10)
    path = str(OUT / "w_tle.xlsx")
    task = ExportTask("test", layer_reader.plan_package(layer, choices, personal={}), path)
    progress: list[float] = []
    task.progressChanged.connect(progress.append)
    start_task(task)
    wait(task)
    timer.stop()
    check(task.result is not None and task.result.rows == 30_000 and os.path.isfile(path),
          f"zadanie w tle zapisuje plik ({task.error})")
    check(len(ticks) >= 5 and len(progress) >= 10, f"pętla zdarzeń działa w trakcie eksportu "
          f"({len(ticks)} tyknięć zegara, {len(progress)} sygnałów postępu)")
    check(load_workbook(path, read_only=True).worksheets[0].max_row > 30_000,
          "plik z zadania w tle ma wszystkie wiersze")

    # Anulowanie w trakcie: poprzednia wersja pliku zostaje, plików tymczasowych brak.
    before = Path(path).read_bytes()
    task = ExportTask("test", layer_reader.plan_package(layer, choices, personal={}), path)
    task.progressChanged.connect(lambda value: task.cancel() if value > 5 else None)
    start_task(task)
    wait(task)
    check(task.result is None and task.error is None and Path(path).read_bytes() == before
          and sorted(p.name for p in OUT.glob(".~*")) == [], "anulowanie zostawia poprzedni plik bez zmian")

    # Przegląd R5, zarzut 1: relacja bez pamięci podręcznej — w tle bez formattera, wartość jako tekst.
    broken = QgsVectorLayer("None?field=ref:integer", "relacja", "memory")
    broken.setEditorWidgetSetup(0, QgsEditorWidgetSetup("RelationReference", {"Relation": "nie_istnieje"}))
    f = QgsFeature(broken.fields())
    f.setAttributes([7])
    broken.dataProvider().addFeatures([f])
    rel = layer_reader.plan_package(broken, [FieldChoice("ref")], personal={})
    check(rel.read().rows == [["7"]] and rel.converters[0].__name__ == "<lambda>",  # nie formatter
          "relacja bez pamięci podręcznej: wartość jako tekst, bez formattera w tle")

    # Przegląd R6, zarzut 2: pole typu, którego warstwa w pamięci nie zna, nie może sprawić,
    # że formatter w tle dostanie oryginalną warstwę zamiast kopii pól.
    from qgis.core import QgsField, QgsFields
    from qgis.PyQt.QtCore import QMetaType

    odd = QgsFields()
    for field in (QgsField("a", QMetaType.Type.Int), QgsField("znak", QMetaType.Type.QChar),
                  QgsField("b", QMetaType.Type.QString)):
        odd.append(field)
    naive = QgsVectorLayer("None", "proba", "memory")
    naive.dataProvider().addAttributes(odd.toList())
    naive.updateFields()

    class OddLayer:
        def fields(self) -> QgsFields:
            return odd

    copy = layer_reader.fields_only(OddLayer())  # type: ignore[arg-type]
    check(naive.fields().names() == ["a", "b"] and isinstance(copy, QgsVectorLayer)
          and copy.fields().names() == ["a", "znak", "b"],
          "pole nieznanego typu: kopia pól zachowuje numery, bez oryginału w tle")

    # Warstwa usunięta z projektu w trakcie eksportu: zadanie czyta własną kopię źródła.
    QgsProject.instance().addMapLayer(layer)
    task = ExportTask("test", layer_reader.plan_package(layer, choices, personal={}), str(OUT / "usunieta.xlsx"))
    start_task(task)
    QgsProject.instance().removeMapLayer(layer.id())
    wait(task)
    check(task.result is not None and task.result.rows == 30_000, "usunięcie warstwy w trakcie nie psuje eksportu")
    sheet = load_workbook(str(OUT / "usunieta.xlsx"), read_only=True).worksheets[0]
    check(list(sheet.iter_rows(min_row=HEADER_ROW + 1, max_row=HEADER_ROW + 1, values_only=True))[0][3] == "usługi",
          "etykieta value map z zadania w tle")


def workbook_cells(path: str) -> dict[str, object]:
    """Treść i wygląd wszystkich arkuszy — do porównania dwóch plików.

    Różni się tylko data eksportu: zastępujemy dokładnie ją (odczytaną z arkusza
    technicznego tego pliku), nie każdy tekst wyglądający na datę.
    """
    from data_handover.core.xlsx import META_SHEET, display

    wb = load_workbook(path)
    meta = {str(row[0]): row[1] for row in wb[META_SHEET].iter_rows(values_only=True)}
    stamp = dt.datetime.fromisoformat(str(meta["exported_at"]))
    shown = display(stamp, None, str(meta["language"]))
    sheets: dict[str, object] = {}
    for ws in wb.worksheets:
        cells = []
        for row in ws.iter_rows():
            for c in row:
                value = c.value
                if isinstance(value, str):
                    value = value.replace(stamp.isoformat(), "<data>").replace(shown, "<data>")
                color = c.font.color.rgb if c.font.color is not None else None
                cells.append((c.coordinate, value, c.number_format, c.font.name, c.font.sz, c.font.b, c.font.i,
                              c.font.u, color, c.fill.fgColor.rgb, repr(c.border), c.alignment.horizontal,
                              c.alignment.vertical, c.alignment.wrap_text,
                              c.comment.text if c.comment is not None else None,
                              c.hyperlink.target if c.hyperlink is not None else None))
        rules = [(str(cf.sqref), r.type, r.formula, repr(r.dxf)) for cf in ws.conditional_formatting for r in cf.rules]
        widths = {k: d.width for k, d in ws.column_dimensions.items()}
        page = (ws.page_setup.orientation, ws.page_setup.fitToWidth, ws.sheet_properties.pageSetUpPr.fitToPage
                if ws.sheet_properties.pageSetUpPr else None, ws.print_title_rows, ws.oddFooter.center.text,
                ws.oddHeader.center.text, ws.sheet_view.showGridLines)
        tables = sorted((t.name, t.ref, repr(t.tableStyleInfo)) for t in ws.tables.values())
        sheets[ws.title] = (ws.sheet_state, cells, tables, ws.freeze_panes, rules, widths, page,
                            sorted(str(m) for m in ws.merged_cells.ranges), ws.auto_filter.ref,
                            ws.sheet_properties.tabColor.rgb if ws.sheet_properties.tabColor else None)
    return sheets


def gpkg_layer(name: str) -> QgsVectorLayer:
    """Warstwa testowa zapisana do GeoPackage (`qgis_process` czyta ją z pliku projektu)."""
    memory = make_layer()
    path = str(OUT / f"{name}.gpkg")
    options = QgsVectorFileWriter.SaveVectorOptions()
    options.driverName = "GPKG"
    options.layerName = name
    error = QgsVectorFileWriter.writeAsVectorFormatV3(memory, path, QgsProject.instance().transformContext(),
                                                      options)
    assert error[0] == QgsVectorFileWriter.WriterError.NoError, error
    layer = QgsVectorLayer(path, name, "ogr")
    assert layer.isValid() and layer.featureCount() == 3
    for i, field in enumerate(memory.fields()):
        index = layer.fields().indexOf(field.name())
        layer.setFieldAlias(index, memory.attributeAlias(i))
        layer.setEditorWidgetSetup(index, memory.editorWidgetSetup(i))
    QgsProject.instance().removeMapLayer(memory.id())
    QgsProject.instance().addMapLayer(layer)
    return layer


def window_export(iface: object, layer: QgsVectorLayer, path: str, previous: str = "") -> None:
    """Eksport z okna z zapisanymi ustawieniami (okno dostaje ścieżkę zamiast okna zapisu).

    Bez `previous` — bez arkusza „Zmiany”; z nim okno porównuje z tym plikiem
    według klucza zapisanego przy warstwie.
    """
    import data_handover.dialog as dialog_module

    dialog_module.QFileDialog.getSaveFileName = lambda *args, **kwargs: (path, "")
    layer.setCustomProperty(dialog_module.PREVIOUS, previous)
    box = HandoverDialog(iface, None, layer=layer)
    box.export()
    wait(box.task)
    assert box.last_result is not None, box.task.error


def processing_algorithm(iface: object) -> None:
    """R6: algorytm Processing daje plik identyczny z eksportem z okna przy tych samych ustawieniach."""
    import processing  # noqa: E402

    plugin = classFactory(iface)
    plugin.initGui()
    alg = QgsApplication.processingRegistry().algorithmById("dataHandover:package")
    check(alg is not None, "initGui rejestruje algorytm dataHandover:package")

    layer = gpkg_layer("dzialki_r6")
    # Ustawienia pól jak po pracy w oknie: kolejność, nagłówek, jednostka, opis, pominięta kolumna, klucz.
    names = layer.fields().names()
    choices = [FieldChoice(n, export=n not in ("nazwisko", "data_akt")) for n in ["teryt"] + [
        n for n in names if n != "teryt"]]
    for c in choices:
        if c.name == "pow_m2":
            c.header, c.unit, c.description = "Pow. [m²]", "m²", "Powierzchnia z ewidencji"
        if c.name == "nr_ident":
            c.protect = "mask"
    # R7: kolumny z geometrii zapisane przy warstwie — algorytm liczy je tak samo jak okno.
    layer_reader.save_choices(layer, choices, "nr_dz", GeoColumns(xy=True, wgs84=True, link=True, map="geoportal"))
    look = look_with_theme(Look(as_table=False), THEMES["forest"])
    settings = {"look": json.dumps(to_dict(look)),
                "sheets": json.dumps(to_dict(Sheets(dictionary=True, about=True))),
                "author": "Wydział Geodezji", "file_language": "pl"}
    for key, value in settings.items():
        QgsSettings().setValue("data_handover/" + key, value)

    from_window = str(OUT / "r6_okno.xlsx")
    window_export(iface, layer, from_window)
    from_alg = str(OUT / "r6_algorytm.xlsx")
    result = processing.run("dataHandover:package", {"INPUT": layer, "OUTPUT": from_alg})
    check(result["OUTPUT"] == from_alg and workbook_cells(from_alg) == workbook_cells(from_window),
          "processing.run: plik identyczny z eksportem z okna (arkusze, wartości, wygląd)")
    meta = [r[1] for r in load_workbook(from_alg)["_handover"].iter_rows(values_only=True) if r[0] == "column"]
    check([m.split(":")[0] for m in meta[-5:]] == ["@x", "@y", "@lat", "@lon", "@link"],
          f"algorytm: kolumny geometrii z ustawień warstwy {meta}")
    data = load_workbook(from_alg)["Dane"]
    check(load_workbook(from_alg).sheetnames[:3] == ["Dane", "Słownik", "Metryka"] and not data.tables
          and data.cell(row=HEADER_ROW, column=1).fill.fgColor.rgb.endswith(THEMES["forest"].header_fill)
          and data.cell(row=HEADER_ROW, column=1).value == "teryt",
          "arkusze, język, wygląd i kolejność kolumn z ustawień okna")
    english = str(OUT / "r6_en.xlsx")
    processing.run("dataHandover:package", {"INPUT": layer, "LANGUAGE": 1 + LANGUAGES.index("en"), "OUTPUT": english})
    check(load_workbook(english).sheetnames[0] == tr("en", "sheet.data"), "język pliku wybrany w parametrze")

    # Przybornik uruchamia algorytm w tle: przygotowanie w wątku okna, zapis w wątku tła.
    in_background = str(OUT / "r6_w_tle.xlsx")
    context, feedback = QgsProcessingContext(), QgsProcessingFeedback()
    context.setProject(QgsProject.instance())
    task = QgsProcessingAlgRunnerTask(alg, {"INPUT": layer.id(), "OUTPUT": in_background}, context, feedback)
    finished: list[bool] = []
    task.executed.connect(lambda ok, _results: finished.append(ok))
    QgsApplication.taskManager().addTask(task)
    deadline = time.monotonic() + 60
    while not finished and time.monotonic() < deadline:
        QApplication.processEvents()
        time.sleep(0.005)
    check(finished == [True] and workbook_cells(in_background) == workbook_cells(from_window),
          "algorytm w tle (jak z Przybornika) daje ten sam plik")

    model_thread(layer, from_window)

    themed = str(OUT / "r6_motyw.xlsx")
    processing.run("dataHandover:package", {"INPUT": layer, "THEME": 1 + list(THEMES).index("plum"),
                                            "OUTPUT": themed})
    header = load_workbook(themed)["Dane"].cell(row=HEADER_ROW, column=1)
    check(header.fill.fgColor.rgb.endswith(THEMES["plum"].header_fill) and header.font.name == look.header.font,
          "motyw zmienia kolory zapisanego wyglądu, układ zostaje")

    # Ta sama podstawa porównania w oknie i w algorytmie; zmiana jednej wartości po pierwszym eksporcie.
    layer.startEditing()
    first = next(layer.getFeatures())
    layer.changeAttributeValue(first.id(), layer.fields().indexOf("pow_m2"), 1600.0)
    layer.commitChanges()
    compared_window = str(OUT / "r6_zmiany_okno.xlsx")
    window_export(iface, layer, compared_window, previous=from_window)
    compared = str(OUT / "r6_zmiany.xlsx")
    processing.run("dataHandover:package", {"INPUT": layer, "PREVIOUS": from_window, "OUTPUT": compared})
    changes = load_workbook(compared)["Zmiany"]
    check(workbook_cells(compared) == workbook_cells(compared_window)
          and any(c.value == "12/3" for row in changes.iter_rows() for c in row),
          "poprzedni plik bez klucza: klucz zapisany w oknie, arkusz „Zmiany” jak z okna")
    layer_reader.save_choices(layer, choices, None)
    class Errors(QgsProcessingFeedback):
        """Komunikaty błędów, które użytkownik widzi w oknie algorytmu."""

        def __init__(self) -> None:
            super().__init__()
            self.messages: list[str] = []

        def reportError(self, error: str, fatalError: bool = False) -> None:  # noqa: N802, N803
            self.messages.append(error)

    errors = Errors()
    for params in ({"PREVIOUS": from_window}, {"PREVIOUS": str(OUT / "nie_ma.xlsx"), "KEY": "nr_dz"},
                   {"SELECTED_ONLY": True}):
        try:
            processing.run("dataHandover:package", {"INPUT": layer, "OUTPUT": str(OUT / "zly.xlsx"), **params},
                           feedback=errors)
        except Exception:  # noqa: S110 - QgsProcessingException; treść sprawdzamy w `errors`
            pass
    lang = ui_language()
    expected = [tr(lang, k) for k in ("ui.error.no_key", "ui.error.no_previous", "alg.error.no_selection")]
    check(all(e in errors.messages for e in expected) and not os.path.exists(OUT / "zly.xlsx"),
          "brak klucza, brak poprzedniego pliku, brak zaznaczenia: czytelny błąd, bez pliku")
    layer.selectByIds([next(layer.getFeatures()).id()])
    selected = str(OUT / "r6_zaznaczone.xlsx")
    processing.run("dataHandover:package", {"INPUT": layer, "SELECTED_ONLY": True, "OUTPUT": selected})
    check(load_workbook(selected)["Dane"].max_row == HEADER_ROW + 1, "tylko zaznaczone")
    layer.removeSelection()
    layer_reader.save_choices(layer, choices, "nr_dz")

    # Przegląd R6, zarzut 1: po zawężeniu filtra do dwóch obiektów PESEL-e nie są już wykrywane
    # po wartościach (za mała próbka), a zapisane maskowanie nadal obowiązuje — w oknie i w algorytmie.
    layer.setSubsetString('"fid" <= 2')
    narrowed = layer_reader.personal_data(layer)
    masked = str(OUT / "r6_mala_probka.xlsx")
    processing.run("dataHandover:package", {"INPUT": layer, "OUTPUT": masked})
    values = [c.value for row in load_workbook(masked)["Dane"].iter_rows(min_row=HEADER_ROW + 1) for c in row]
    check("nr_ident" in narrowed and "44051401359" not in values and "***" in values,
          "zapisane maskowanie obowiązuje, choć mała próbka nie wykrywa pola")
    layer.setSubsetString("")
    qgis_process(iface, layer, settings)
    for key in settings:
        QgsSettings().remove("data_handover/" + key)
    plugin.unload()
    check(QgsApplication.processingRegistry().algorithmById("dataHandover:package") is None,
          "unload wyrejestrowuje algorytm")


def model_thread(layer: QgsVectorLayer, from_window: str) -> None:
    """W modelu graficznym (uruchomionym w tle) QGIS woła prepareAlgorithm w wątku okna."""
    from qgis.core import (
        QgsProcessingModelAlgorithm,
        QgsProcessingModelChildAlgorithm,
        QgsProcessingModelChildParameterSource,
        QgsProcessingModelOutput,
        QgsProcessingModelParameter,
        QgsProcessingParameterVectorLayer,
    )
    from qgis.PyQt.QtCore import QThread

    from data_handover import processing_provider

    on_main: list[bool] = []
    original = processing_provider.PackageAlgorithm.prepareAlgorithm

    def spy(self: object, *args: object) -> bool:
        on_main.append(QThread.currentThread() == QApplication.instance().thread())
        return original(self, *args)  # type: ignore[arg-type]

    processing_provider.PackageAlgorithm.prepareAlgorithm = spy  # type: ignore[method-assign]
    model = QgsProcessingModelAlgorithm("r6", "testy")
    model.addModelParameter(QgsProcessingParameterVectorLayer("warstwa", "warstwa"),
                            QgsProcessingModelParameter("warstwa"))
    child = QgsProcessingModelChildAlgorithm("dataHandover:package")
    child.setChildId("paczka")
    child.addParameterSources("INPUT", [QgsProcessingModelChildParameterSource.fromModelParameter("warstwa")])
    output = QgsProcessingModelOutput("plik", "plik")
    output.setChildId("paczka")
    output.setChildOutputName("OUTPUT")
    child.setModelOutputs({"plik": output})
    model.addChildAlgorithm(child)
    model.updateDestinationParameters()
    path = str(OUT / "r6_model.xlsx")
    context, feedback = QgsProcessingContext(), QgsProcessingFeedback()
    context.setProject(QgsProject.instance())
    task = QgsProcessingAlgRunnerTask(model, {"warstwa": layer.id(), "paczka:plik": path}, context, feedback)
    finished: list[bool] = []
    task.executed.connect(lambda ok, _results: finished.append(ok))
    QgsApplication.taskManager().addTask(task)
    deadline = time.monotonic() + 60
    while not finished and time.monotonic() < deadline:
        QApplication.processEvents()
        time.sleep(0.005)
    processing_provider.PackageAlgorithm.prepareAlgorithm = original  # type: ignore[method-assign]
    check(finished == [True] and on_main == [True] and workbook_cells(path) == workbook_cells(from_window),
          f"model graficzny w tle: przygotowanie w wątku okna, ten sam plik ({finished}, {on_main})")


def qgis_process(iface: object, layer: QgsVectorLayer, settings: dict[str, str]) -> None:
    """Kryterium R6: `qgis_process run dataHandover:package` daje ten sam plik co okno."""
    project = str(OUT / "r6.qgz")
    check(QgsProject.instance().write(project), "projekt z ustawieniami pól zapisany")
    # Jednorazowy katalog domowy: test nie może zmienić prawdziwego profilu QGIS,
    # nawet uruchomiony poza kontenerem.
    home = OUT / "home"
    home.mkdir(exist_ok=True)
    env = dict(os.environ, QGIS_PLUGINPATH=str(ROOT), QT_QPA_PLATFORM="offscreen", HOME=str(home))
    for name in ("QGIS_CUSTOM_CONFIG_PATH", "XDG_DATA_HOME", "XDG_CONFIG_HOME"):
        env.pop(name, None)

    def run(*args: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(["qgis_process", *args], env=env, capture_output=True, text=True, timeout=180)

    enabled = run("plugins", "enable", "data_handover")
    assert enabled.returncode == 0, enabled.stdout + enabled.stderr
    # W prawdziwej instalacji okno i qgis_process czytają ten sam profil — test ma osobny,
    # więc ustawienia okna kopiujemy do pliku ustawień qgis_process.
    inis = [p for p in glob.glob(str(home / ".local/share/QGIS/*/profiles/default/QGIS/*.ini"))
            if "data_handover=true" in Path(p).read_text(errors="replace")]
    assert len(inis) == 1, inis
    ini = QSettings(inis[0], QSettings.Format.IniFormat)
    for key, value in settings.items():
        ini.setValue("data_handover/" + key, value)
    ini.sync()

    from_window = str(OUT / "r6_okno_qgis_process.xlsx")
    window_export(iface, layer, from_window)
    from_cli = str(OUT / "r6_qgis_process.xlsx")
    done = run("run", "dataHandover:package", f"--PROJECT_PATH={project}", "--",
               f"INPUT={layer.id()}", f"LANGUAGE={1 + LANGUAGES.index('pl')}", f"OUTPUT={from_cli}")
    check(done.returncode == 0 and os.path.isfile(from_cli), f"qgis_process run dataHandover:package "
          f"(kod {done.returncode}){'' if done.returncode == 0 else ': ' + done.stdout[-800:] + done.stderr[-800:]}")
    check(workbook_cells(from_cli) == workbook_cells(from_window),
          "qgis_process: plik identyczny z eksportem z okna")


PARCEL = ("POLYGON((500000 480000, 500120 480000, 500120 480050, 500040 480050, 500040 480130, "
          "500000 480130, 500000 480000))")


def geo_layer(wkt_type: str, crs: str, geometries: list[str | None]) -> QgsVectorLayer:
    layer = QgsVectorLayer(f"{wkt_type}?crs={crs}&field=nr_dz:string", f"geo_{wkt_type}", "memory")
    features = []
    for i, wkt in enumerate(geometries, start=1):
        f = QgsFeature(layer.fields())
        f.setAttributes([f"{i}/1"])
        if wkt is not None:
            f.setGeometry(QgsGeometry.fromWkt(wkt))
        features.append(f)
    assert layer.dataProvider().addFeatures(features)[0]
    QgsProject.instance().addMapLayer(layer)
    return layer


def qgis_measure(layer: QgsVectorLayer, expression: str, feature: QgsFeature | None = None) -> float:
    """Wartość `$area`/`$length` tak, jak liczy ją kalkulator pól QGIS (elipsoida i jednostki projektu)."""
    from qgis.core import QgsDistanceArea, QgsExpression, QgsExpressionContext, QgsExpressionContextUtils

    project = QgsProject.instance()
    calc = QgsDistanceArea()
    calc.setSourceCrs(layer.crs(), project.transformContext())
    calc.setEllipsoid(project.ellipsoid())
    exp = QgsExpression(expression)
    exp.setGeomCalculator(calc)
    exp.setDistanceUnits(project.distanceUnits())
    exp.setAreaUnits(project.areaUnits())
    context = QgsExpressionContext(QgsExpressionContextUtils.globalProjectLayerScopes(layer))
    context.setFeature(feature if feature is not None else next(layer.getFeatures()))
    return float(exp.evaluate(context))


def geometry() -> None:
    """R7: kolumny z geometrii — zgodność z $area/$length, punkt wewnętrzny, WGS84, link, zapis ustawień."""
    from qgis.core import QgsCoordinateReferenceSystem, QgsCoordinateTransform

    project = QgsProject.instance()
    project.setCrs(QgsCoordinateReferenceSystem("EPSG:2180"))  # bez CRS projekt zgłasza elipsoidę „NONE”
    project.setEllipsoid("EPSG:7019")  # GRS 1980
    check(project.ellipsoid() == "EPSG:7019", "projekt testowy z elipsoidą GRS 1980")
    layer = geo_layer("Polygon", "EPSG:2180", [PARCEL, None])
    check(layer_reader.geometry_kind(layer) == "polygon", "typ geometrii: poligon")
    all_on = GeoColumns(measure=True, area_unit="m2", xy=True, wgs84=True, link=True, map="geoportal")
    pkg = layer_reader.build_package(layer, [FieldChoice("nr_dz")], geo=all_on, ellipsoid=project.ellipsoid(),
                                     personal={}, language="pl")
    names = [c.name for c in pkg.columns]
    check(names == ["nr_dz", "@area_m2:EPSG:2180:EPSG:7019", "@x:EPSG:2180", "@y:EPSG:2180", "@lat", "@lon", "@link"],
          f"kolumny geometrii (jednostka i metoda w nazwie): {names}")
    row = {n.split(":")[0]: v for n, v in zip(names, pkg.rows[0])}
    expected = qgis_measure(layer, "$area")
    check(abs(row["@area_m2"] - expected) <= 0.01, f"powierzchnia {row['@area_m2']} = $area {expected:.4f} (±0,01 m²)")
    planar = QgsGeometry.fromWkt(PARCEL).area()
    check(abs(row["@area_m2"] / planar - 1) < 0.003 and row["@area_m2"] != round(planar, 2),
          "powierzchnia elipsoidalna bliska płaskiej w EPSG:2180, ale nie płaska")
    inside = QgsGeometry.fromPointXY(QgsPointXY(row["@x"], row["@y"]))
    check(QgsGeometry.fromWkt(PARCEL).contains(inside), "punkt wewnętrzny leży w działce w kształcie „L”")
    to_wgs = QgsCoordinateTransform(layer.crs(), QgsCoordinateReferenceSystem("EPSG:4326"), project)
    wgs = to_wgs.transform(QgsPointXY(row["@x"], row["@y"]))
    check(abs(row["@lat"] - wgs.y()) < 2e-6 and abs(row["@lon"] - wgs.x()) < 2e-6, "szerokość i długość WGS84")
    check(row["@link"] == map_url("geoportal", row["@lat"], row["@lon"]), "link do Geoportalu")
    check(all(v is None for v in pkg.rows[1][1:]), "obiekt bez geometrii: puste komórki")
    area = pkg.columns[1]
    check("EPSG:2180" in area.description and "ewidencyjn" in area.description and area.unit == "m²",
          f"opis metody: {area.description}")

    ha = layer_reader.build_package(layer, [FieldChoice("nr_dz")], geo=GeoColumns(measure=True),
                                    ellipsoid=project.ellipsoid(), personal={})
    check(ha.rows[0][1] == round(row["@area_m2"] / 10_000, 4) and ha.columns[1].unit == "ha", "powierzchnia w ha")

    line = geo_layer("LineString", "EPSG:2180", ["LINESTRING(500000 480000, 500300 480400, 501000 480400)"])
    lpkg = layer_reader.build_package(line, [FieldChoice("nr_dz")], geo=GeoColumns(measure=True, length_unit="km"),
                                      ellipsoid=project.ellipsoid(), personal={})
    check(lpkg.columns[1].name.startswith("@length_km:")
          and abs(lpkg.rows[0][1] - round(qgis_measure(line, "$length") / 1000, 3)) < 1e-9, "długość linii = $length")

    # Projekt bez elipsoidy, warstwa w stopniach: elipsoida układu warstwy, nie stopnie kwadratowe.
    project.setEllipsoid("NONE")
    degrees = geo_layer("Polygon", "EPSG:4326", ["POLYGON((21 52, 21.001 52, 21.001 52.001, 21 52.001, 21 52))"])
    dpkg = layer_reader.build_package(degrees, [FieldChoice("nr_dz")], geo=GeoColumns(measure=True, area_unit="m2"),
                                      ellipsoid=project.ellipsoid(), personal={})
    check(7_000 < dpkg.rows[0][1] < 8_000 and "WGS" in dpkg.columns[1].description,
          f"bez elipsoidy projektu: elipsoida układu ({dpkg.rows[0][1]} m²)")
    project.setEllipsoid("EPSG:7019")

    table = QgsVectorLayer("None?field=nr_dz:string", "tabela", "memory")
    tpkg = layer_reader.build_package(table, [FieldChoice("nr_dz")], geo=all_on, personal={})
    check(layer_reader.geometry_kind(table) == "none" and [c.name for c in tpkg.columns] == ["nr_dz"],
          "warstwa bez geometrii: bez kolumn geometrii")

    # Ustawienia geometrii zapisane przy warstwie; zapis samych pól ich nie kasuje.
    layer_reader.save_choices(layer, [FieldChoice("nr_dz")], None, all_on)
    layer_reader.save_choices(layer, [FieldChoice("nr_dz", header="Nr")], None)
    choices, _, geo, _ = layer_reader.saved_choices(layer, {})
    check(geo == all_on and choices[0].header == "Nr", "ustawienia geometrii przy warstwie przeżywają zapis pól")

    # W tle (QgsTask) te same wartości co w wątku okna.
    path = str(OUT / "r7_geometria.xlsx")
    task = ExportTask("geo", layer_reader.plan_package(layer, [FieldChoice("nr_dz")], geo=all_on,
                                                       ellipsoid=project.ellipsoid(), personal={},
                                                       language="pl"), path)
    start_task(task)
    wait(task)
    data = load_workbook(path)["Dane"]
    check(task.result is not None and data.cell(row=HEADER_ROW + 1, column=2).value == row["@area_m2"]
          and data.cell(row=HEADER_ROW + 1, column=7).hyperlink.target == row["@link"],
          "eksport w tle liczy geometrię tak samo")
    sample = layer_reader.build_package(layer, [FieldChoice("nr_dz", header="Nr działki")],
                                        geo=GeoColumns(measure=True, xy=True, wgs84=True, link=True, map="geoportal"),
                                        ellipsoid=project.ellipsoid(), personal={}, language="pl", title="Działki")
    export(str(ROOT / "build" / f"przyklad_geometria_qgis{VERSION}.xlsx"), sample, Look(), Sheets(dictionary=True))

    # Przegląd R7, zarzut 1: warstwa bez CRS — bez udawanych stopni WGS84 i linku, z uwagą.
    nocrs = geo_layer("Polygon", "", [PARCEL])
    npkg = layer_reader.build_package(nocrs, [FieldChoice("nr_dz")], geo=all_on, ellipsoid=project.ellipsoid(),
                                      personal={}, language="pl")
    nnames = [c.name.split(":")[0] for c in npkg.columns]
    check(not nocrs.crs().isValid() and nnames == ["nr_dz", "@x", "@y"]
          and tr("pl", "note.geo.wgs_skipped", crs="?") in npkg.notes
          and tr("pl", "note.geo.measure_skipped", crs="?") in npkg.notes,
          f"warstwa bez CRS: tylko X/Y, uwagi o pominięciu ({nnames})")

    # Zarzut 2: kontekst transformacji z algorytmu, nie z globalnego projektu.
    from qgis.core import QgsCoordinateTransformContext

    custom = QgsCoordinateTransformContext()
    operation = ("+proj=pipeline +step +inv +proj=tmerc +lat_0=0 +lon_0=19 +k=0.9993 +x_0=500000 "
                 "+y_0=-5300000 +ellps=GRS80 +step +proj=unitconvert +xy_in=rad +xy_out=deg")
    custom.addCoordinateOperation(layer.crs(), QgsCoordinateReferenceSystem("EPSG:4326"), operation)
    cplan = layer_reader.plan_package(layer, [FieldChoice("nr_dz")], geo=GeoColumns(wgs84=True), personal={},
                                      transform_context=custom)
    check(cplan.geo.context.coordinateOperations() == custom.coordinateOperations()
          and cplan.geo.context != project.transformContext(), "plan używa podanego kontekstu transformacji")

    # Zarzut 3: zmiana układu warstwy to dla porównania inne kolumny, nie zmiana każdego wiersza.
    to_2177 = QgsCoordinateTransform(layer.crs(), QgsCoordinateReferenceSystem("EPSG:2177"), project)
    moved = QgsGeometry.fromWkt(PARCEL)
    moved.transform(to_2177)
    other = geo_layer("Polygon", "EPSG:2177", [moved.asWkt()])
    xy = GeoColumns(measure=True, xy=True)
    first = str(OUT / "r7_2180.xlsx")
    export(first, layer_reader.build_package(layer, [FieldChoice("nr_dz")], geo=xy, key="nr_dz", personal={},
                                             ellipsoid=project.ellipsoid(), language="pl"))
    second = export(str(OUT / "r7_2177.xlsx"), layer_reader.build_package(
        other, [FieldChoice("nr_dz")], geo=xy, key="nr_dz", personal={}, ellipsoid=project.ellipsoid(),
        language="pl"), previous=first)
    check(not [c for c in second.changes if c.kind == "modified"]
          and any(n.startswith(tr("pl", "note.columns.added", names="")[:20]) for n in second.notes),
          f"zmiana układu: kolumny dodane/usunięte, bez fałszywych zmian ({second.changes}, {second.notes})")

    # Zarzut 4: pole warstwy o nazwie kolumny geometrii — nazwy unikalne, plik nadaje się do porównania.
    from qgis.core import QgsField
    from qgis.PyQt.QtCore import QMetaType

    clash = geo_layer("Point", "EPSG:2180", ["POINT(500000 480000)"])
    clash.dataProvider().addAttributes([QgsField("@lat", QMetaType.Type.QString)])
    clash.updateFields()
    cpkg = layer_reader.build_package(clash, [FieldChoice("nr_dz"), FieldChoice("@lat")], geo=GeoColumns(wgs84=True),
                                      key="nr_dz", personal={}, language="pl")
    cnames = [c.name for c in cpkg.columns]
    v1 = str(OUT / "r7_kolizja.xlsx")
    export(v1, cpkg)
    again_pkg = layer_reader.build_package(clash, [FieldChoice("nr_dz"), FieldChoice("@lat")],
                                           geo=GeoColumns(wgs84=True), key="nr_dz", personal={}, language="pl")
    check(len(set(cnames)) == len(cnames) and "@lat~2" in cnames
          and export(str(OUT / "r7_kolizja2.xlsx"), again_pkg, previous=v1).changes == [],
          f"kolizja nazwy z polem warstwy: unikalne nazwy, ponowne porównanie działa ({cnames})")

    # Zarzut 6: obiekt, którego nie da się przeliczyć, to uwaga z liczbą, nie cichy brak.
    bad = geo_layer("Point", "EPSG:4326", ["POINT(10 95)", "POINT(21 52)"])
    bpkg = layer_reader.build_package(bad, [FieldChoice("nr_dz")], geo=GeoColumns(wgs84=True, link=True),
                                      personal={}, language="pl")
    check(bpkg.rows[0][1:] == [None, None, None] and bpkg.rows[1][1] == 52.0
          and tr("pl", "note.geo.failed", n=1) in bpkg.notes, "nieudane przeliczenie: puste komórki i uwaga z liczbą")
    for gone in (nocrs, other, clash, bad):
        project.removeMapLayer(gone.id())

    # Okno (wariant A): grupa pod listą pól, pomiar zależny od typu geometrii, zapis przy warstwie.
    layer_reader.save_choices(layer, [FieldChoice("nr_dz")], None, GeoColumns())
    box = HandoverDialog(None, None, layer=layer)
    check(box.geo_box.isEnabled() and box.measure_check.text() == tr(ui_language(), "ui.geo.area")
          and box.geo_columns() == GeoColumns(), "okno: grupa geometrii, domyślnie nic nie zaznaczone")
    box.measure_check.setChecked(True)
    box.unit_combo.setCurrentIndex(box.unit_combo.findData("m2"))
    box.link_check.setChecked(True)
    box.map_combo.setCurrentIndex(box.map_combo.findData("geoportal"))
    chosen = box.geo_columns()
    box._save_state()
    box.close()
    again = HandoverDialog(None, None, layer=layer)
    check(chosen == GeoColumns(measure=True, area_unit="m2", link=True, map="geoportal")
          and again.geo_columns() == chosen, "okno: wybór geometrii zapamiętany przy warstwie")
    again.close()
    lbox = HandoverDialog(None, None, layer=line)
    check(lbox.measure_check.text() == tr(ui_language(), "ui.geo.length")
          and [lbox.unit_combo.itemData(i) for i in range(lbox.unit_combo.count())] == ["m", "km"],
          "okno: na warstwie liniowej długość w m albo km")
    lbox.close()
    project.addMapLayer(table)
    tbox = HandoverDialog(None, None, layer=table)
    check(not tbox.geo_box.isEnabled(), "okno: warstwa bez geometrii — grupa wyłączona")
    tbox.close()
    QgsSettings().setValue(UI_LANGUAGE, "pl")
    for lang in ("pl", "en"):
        QgsSettings().setValue(UI_LANGUAGE, lang)
        shot = HandoverDialog(None, None, layer=layer)
        shot.resize(980, 700)
        shot.show()
        QApplication.processEvents()
        shot.grab().save(str(SCREENS / f"okno_geometria_{lang}_qgis{VERSION}.png"))
        shot.close()
    QgsSettings().remove(UI_LANGUAGE)
    for name in (layer, line, degrees, table):
        project.removeMapLayer(name.id())


def labelled(path: str, sheet: str, label: str) -> object:
    """Wartość z komórki obok pierwszej komórki o treści `label` (para etykieta–wartość)."""
    for row in load_workbook(path)[sheet].iter_rows():
        for i, c in enumerate(row[:-1]):
            if c.value == label:
                return row[i + 1].value
    raise AssertionError(f"brak etykiety „{label}” w arkuszu {sheet}")


def summaries(iface: object) -> None:
    """R8: które kolumny są sumowane, sumy w „Podsumowaniu” zgodne z QGIS, zapis ustawień, Processing."""
    import processing  # noqa: E402
    from qgis.core import QgsCoordinateReferenceSystem

    project = QgsProject.instance()
    project.setCrs(QgsCoordinateReferenceSystem("EPSG:2180"))
    project.setEllipsoid("EPSG:7019")
    plugin = classFactory(iface)
    plugin.initGui()

    # Warstwa z make_layer() + pole dziesiętne i pole „rok” (całkowite).
    layer = make_layer()
    layer.dataProvider().addAttributes([QgsField("wartosc", QMetaType.Type.Double),
                                        QgsField("rok", QMetaType.Type.Int)])
    layer.updateFields()
    layer.startEditing()
    for feature, (value, year) in zip(layer.getFeatures(), [(10.5, 2019), (20.25, 2020), (None, 2021)]):
        layer.changeAttributeValue(feature.id(), layer.fields().indexOf("wartosc"), value)
        layer.changeAttributeValue(feature.id(), layer.fields().indexOf("rok"), year)
    layer.commitChanges()
    personal = layer_reader.personal_data(layer)
    choices = [FieldChoice(n, export=n != "nazwisko", protect="mask" if n == "nr_ident" else "")
               for n in layer.fields().names()]
    plan = layer_reader.plan_package(layer, choices, personal=personal, language="pl",
                                     summary=Summary(sheet=True, row=True, group_by="teryt"))
    by_name = {c.name: c for c in plan.columns}
    check(plan.summary == Summary(sheet=True, row=True, group_by="teryt"), "plan niesie ustawienia podsumowania")
    check(by_name["pow_m2"].summed and by_name["wartosc"].summed, "pola dziesiętne domyślnie sumowane")
    check(not by_name["rok"].summed, "pole całkowite domyślnie nie jest sumowane")
    check(not by_name["nr_dz"].summed and not by_name["fid"].summed and not by_name["teryt"].summed,
          "identyfikatory nie są sumowane")
    check(not by_name["nr_ident"].summed and not by_name["przezn"].summed and not by_name["uwagi"].summed,
          "kolumny maskowane, z etykietami i tekstowe nie są sumowane")
    forcing = ("rok", "nr_ident", "przezn", "fid")
    on = [FieldChoice(c.name, export=c.export, protect=c.protect, sum=True if c.name in forcing else None)
          for c in choices]
    forced = {c.name: c for c in layer_reader.plan_package(layer, on, personal=personal).columns}
    check(forced["rok"].summed and not forced["nr_ident"].summed and not forced["przezn"].summed
          and not forced["fid"].summed, "sum=True włącza liczbowe, ale nie identyfikator, maskę ani etykiety")
    off = [FieldChoice(c.name, export=c.export, protect=c.protect, sum=False if c.name == "pow_m2" else None)
           for c in choices]
    check(not {c.name: c for c in layer_reader.plan_package(layer, off, personal=personal).columns}[
          "pow_m2"].summed, "sum=False wyłącza pole dziesiętne")

    path = str(OUT / "r8_podsumowanie.xlsx")
    export(path, plan.read(), Look(), Sheets())
    sheet = tr("pl", "sheet.summary")
    total = labelled(path, sheet, tr("pl", "summary.sum", header="Powierzchnia"))
    expected = layer.aggregate(QgsAggregateCalculator.Aggregate.Sum, "pow_m2")[0]
    check(isinstance(total, float) and math.isclose(total, expected, rel_tol=1e-12),
          f"suma pow_m2 w „Podsumowaniu” {total} == aggregate(Sum) {expected}")
    params = QgsAggregateCalculator.AggregateParameters()
    params.filter = "\"teryt\" = '0214011'"
    in_group = layer.aggregate(QgsAggregateCalculator.Aggregate.Count, "fid", params)[0]
    check(labelled(path, sheet, "0214011") == in_group == 2, f"liczba w grupie 0214011 == aggregate(Count) {in_group}")
    check(labelled(path, sheet, tr("pl", "summary.count")) == layer.featureCount(), "liczba obiektów w całości")

    # Zapis i odczyt ustawień.
    settings = Summary(sheet=True, row=True, group_by="teryt", measure=False)
    layer_reader.save_choices(layer, [FieldChoice("pow_m2", sum=True), FieldChoice("rok")], None, summary=settings)
    saved = layer_reader.saved_choices(layer, {})
    check(saved[3] == settings, "Summary zapisane przy warstwie i odczytane")
    check({c.name: c.sum for c in saved[0] if c.name in ("pow_m2", "rok")} == {"pow_m2": True, "rok": None},
          "„sum” przeżywa zapis; brak wyboru = None")
    layer_reader.save_choices(layer, [FieldChoice("pow_m2", sum=True)], None)
    check(layer_reader.saved_choices(layer, {})[3] == settings, "zapis bez summary go nie kasuje")
    check(layer_reader.load_choices(layer)["fields"]["pow_m2"]["sum"] is True, "w JSON „sum”: true")
    layer_reader.save_choices(layer, [FieldChoice("rok")], None)
    check("sum" not in layer_reader.load_choices(layer)["fields"]["rok"], "w JSON brak klucza „sum” bez wyboru")
    check(layer_reader.saved_choices(make_layer(), {})[3] == Summary(), "bez zapisu: domyślne Summary")

    # Geometria: pomiar sumowany jak widać w komórkach (zaokrąglone wartości).
    shifted = PARCEL.replace("500000", "510000").replace("500120", "510100").replace("500040", "510030")
    poly = geo_layer("Polygon", "EPSG:2180", [PARCEL, shifted, PARCEL.replace("130", "200")])
    geo = GeoColumns(measure=True, area_unit="ha", xy=True)
    plan = layer_reader.plan_package(poly, [FieldChoice("nr_dz")], geo=geo, ellipsoid=project.ellipsoid(),
                                     personal={}, language="pl", summary=Summary(sheet=True))
    area = next(c for c in plan.columns if c.name.startswith("@area_"))
    check(area.summed and not any(c.summed for c in plan.columns if c.name.startswith("@x")),
          "pomiar z geometrii sumowany, X/Y nie")
    path = str(OUT / "r8_geometria.xlsx")
    export(path, plan.read(), Look(), Sheets())
    each = [qgis_measure(poly, "$area", f) for f in poly.getFeatures()]
    rounded = math.fsum(round(a / 10000, 4) for a in each)
    got = labelled(path, sheet, tr("pl", "summary.sum", header=tr("pl", "geo.col.area")))
    exact = poly.aggregate(QgsAggregateCalculator.Aggregate.Sum, "$area")[0] / 10000
    check(got == rounded and abs(got - exact) <= len(each) * 0.5e-4,
          f"suma pomiaru {got} = suma zaokrąglonych $area {rounded}, od aggregate {exact:.6f} o <= n·0,5e-4")
    plan = layer_reader.plan_package(poly, [FieldChoice("nr_dz")], geo=geo, ellipsoid=project.ellipsoid(),
                                     personal={}, summary=Summary(measure=False))
    check(not any(c.summed for c in plan.columns), "Summary.measure=False wyłącza sumę pomiaru")

    # Processing: ustawienia z warstwy trafiają do pliku tak samo jak z okna.
    gp = gpkg_layer("dzialki_r8")
    r8 = [FieldChoice(n, export=n not in ("nazwisko", "nr_ident")) for n in gp.fields().names()]
    layer_reader.save_choices(gp, r8, None, summary=Summary(sheet=True, row=True, group_by="teryt"))
    for key, value in {"look": json.dumps(to_dict(Look())), "sheets": json.dumps(to_dict(Sheets())),
                       "author": "", "file_language": "pl"}.items():
        QgsSettings().setValue("data_handover/" + key, value)
    from_window = str(OUT / "r8_okno.xlsx")
    window_export(iface, gp, from_window)
    from_alg = str(OUT / "r8_algorytm.xlsx")
    processing.run("dataHandover:package", {"INPUT": gp, "OUTPUT": from_alg})
    check(sheet in load_workbook(from_alg).sheetnames and workbook_cells(from_alg) == workbook_cells(from_window),
          "Processing: „Podsumowanie” i wiersz sum jak w eksporcie z okna")
    summary_window(iface, layer, poly)
    for item in (layer, poly, gp):
        project.removeMapLayer(item.id())


def summary_window(iface: object, layer: QgsVectorLayer, poly: QgsVectorLayer) -> None:
    """R8, okno: grupa „Sumy i podsumowanie” — listy, zapis, eksport, „sum” zgodne z Processing."""
    import processing  # noqa: E402

    def states(combo: object) -> dict[str, bool]:
        return {str(combo.itemData(i)): combo.itemCheckState(i) == Qt.CheckState.Checked  # type: ignore[attr-defined]
                for i in range(combo.count())}  # type: ignore[attr-defined]

    def datas(combo: object) -> list[str]:
        return [str(combo.itemData(i)) for i in range(combo.count())]  # type: ignore[attr-defined]

    layer.removeCustomProperty(layer_reader.PROPERTY)
    box = HandoverDialog(iface, None, layer=layer)
    check(states(box.sum_combo) == {"pow_m2": True, "wartosc": True, "rok": False},
          f"okno: sumowane domyślnie dziesiętne, całkowite nie; bez identyfikatorów i tekstu ({states(box.sum_combo)})")
    check(not box.totals_check.isChecked() and not box.summary_check.isChecked() and box.summary() == Summary(),
          "okno: wiersz sum i arkusz domyślnie wyłączone")
    check("nr_dz" in datas(box.group_combo) and "pow_m2" in datas(box.group_combo)
          and datas(box.group_combo)[0] == "", "okno: grupowanie — pozycja „bez podziału” i kolumny eksportowane")
    for r in range(box.table.rowCount()):  # eksport maskowanego pola: ani w sumach, ani w grupowaniu
        if box.table.item(r, 0).text() == "nr_ident":
            box.table.item(r, 0).setCheckState(Qt.CheckState.Checked)
    check("nr_ident" not in states(box.sum_combo) and "nr_ident" not in datas(box.group_combo),
          "okno: pole maskowane nie jest na liście sum ani grupowania")
    for r in range(box.table.rowCount()):
        if box.table.item(r, 0).text() == "wartosc":
            box.table.item(r, 0).setCheckState(Qt.CheckState.Unchecked)
    check("wartosc" not in states(box.sum_combo) and "wartosc" not in datas(box.group_combo),
          "okno: odznaczone pole znika z obu list")
    box.close()

    # Pomiar z geometrii: na liście, gdy zaznaczony w grupie geometrii.
    pbox = HandoverDialog(iface, None, layer=poly)
    check("@measure" not in datas(pbox.sum_combo), "okno: bez pomiaru w geometrii nie ma go na liście sum")
    pbox.measure_check.setChecked(True)
    check(states(pbox.sum_combo).get("@measure") is True
          and pbox.sum_combo.itemText(pbox.sum_combo.findData("@measure"))
          == pbox.t("ui.summary.measure", header=pbox.measure_check.text()),
          "okno: pomiar z geometrii na liście sum, domyślnie zaznaczony")
    pbox.sum_combo.setItemCheckState(pbox.sum_combo.findData("@measure"), Qt.CheckState.Unchecked)
    pbox.measure_check.setChecked(False)
    pbox.measure_check.setChecked(True)
    check(states(pbox.sum_combo)["@measure"] is False and pbox.summary().measure is False,
          "okno: wybór pomiaru przeżywa odświeżenie listy")
    pbox.close()

    # Zapisany wybór „sum”: okno go pokazuje, zachowuje przy zapisie i stosuje przy eksporcie.
    saved = Summary(sheet=True, group_by="teryt")
    layer_reader.save_choices(layer, [FieldChoice("rok", sum=True), FieldChoice("pow_m2", sum=False)], None,
                              summary=saved)
    sbox = HandoverDialog(iface, None, layer=layer)
    check(states(sbox.sum_combo) == {"pow_m2": False, "wartosc": True, "rok": True}
          and sbox.summary_check.isChecked() and not sbox.totals_check.isChecked()
          and sbox.group_combo.currentData() == "teryt" and sbox.summary() == saved,
          f"okno: zapisane „sum” i podsumowanie widoczne po otwarciu ({states(sbox.sum_combo)})")
    def set_export(win: object, name: str, on: bool) -> None:
        for r in range(win.table.rowCount()):  # type: ignore[attr-defined]
            if win.table.item(r, 0).text() == name:  # type: ignore[attr-defined]
                win.table.item(r, 0).setCheckState(  # type: ignore[attr-defined]
                    Qt.CheckState.Checked if on else Qt.CheckState.Unchecked)

    set_export(sbox, "teryt", False)  # zapisane pole grupowania przestaje być eksportowane
    check(sbox.group_combo.currentData() == "" and sbox.summary().group_by == "",
          f"okno: lista pokazuje „bez podziału”, więc do pliku nie idzie stare pole ({sbox.summary().group_by!r})")
    set_export(sbox, "teryt", True)
    check(sbox.group_combo.currentData() == "teryt" and sbox.summary() == saved,
          "okno: ponowne zaznaczenie pola przywraca grupowanie")
    sbox.group_combo.setCurrentIndex(0)
    check(sbox.summary().group_by == "", "okno: wybór „bez podziału” daje pusty podział")
    sbox.group_combo.setCurrentIndex(sbox.group_combo.findData("teryt"))
    sbox._save_state()
    sbox.close()
    ybox = HandoverDialog(iface, None, layer=layer)
    for name in layer.fields().names():
        set_export(ybox, name, name == "rok")
    preview = re.sub(r"<[^>]+>", "", ybox.preview.text())
    check("2019" in preview and not re.search(r"2\W019", preview),
          f"okno: podgląd pokazuje rok bez separatora ({preview!r})")
    ybox.close()
    kept = {c.name: c.sum for c in layer_reader.saved_choices(layer, {})[0]}
    check(kept["rok"] is True and kept["pow_m2"] is False and kept["wartosc"] is True and kept["fid"] is None
          and kept["nr_dz"] is None and layer_reader.saved_choices(layer, {})[3] == saved,
          f"okno: zapis nie kasuje „sum” ({kept})")
    path = str(OUT / "r8_okno_rok.xlsx")
    window_export(iface, layer, path)
    sheet = tr("pl", "sheet.summary")
    check(labelled(path, sheet, tr("pl", "summary.sum", header="rok")) == 2019 + 2020 + 2021,
          "okno: eksport sumuje pole całkowite zapisane z sum=True")
    check(all(c.value != tr("pl", "summary.sum", header="Powierzchnia")
              for row in load_workbook(path)[sheet].iter_rows() for c in row),
          "okno: eksport nie sumuje pola z sum=False")

    # To samo, co robi Processing, na zapisanych ustawieniach warstwy.
    gp = gpkg_layer("dzialki_r8_sum")
    fields = [FieldChoice(n, export=n not in ("nazwisko", "nr_ident"), sum=False if n == "pow_m2" else None)
              for n in gp.fields().names()]
    layer_reader.save_choices(gp, fields, None, summary=Summary(sheet=True, row=True, group_by="teryt"))
    from_window = str(OUT / "r8_okno_sum.xlsx")
    window_export(iface, gp, from_window)
    from_alg = str(OUT / "r8_algorytm_sum.xlsx")
    processing.run("dataHandover:package", {"INPUT": gp, "OUTPUT": from_alg})
    check(workbook_cells(from_alg) == workbook_cells(from_window)
          and layer_reader.load_choices(gp)["fields"]["pow_m2"]["sum"] is False,
          "okno i Processing: ten sam plik przy zapisanym sum=False")

    # Okno ustawia, co widać; eksport z okna zapisuje to przy warstwie.
    gp_box = HandoverDialog(iface, None, layer=gp)
    gp_box.totals_check.setChecked(True)
    gp_box.summary_check.setChecked(False)
    gp_box.group_combo.setCurrentIndex(gp_box.group_combo.findData("nr_dz"))
    gp_box._save_state()
    now = layer_reader.saved_choices(gp, {})[3]
    gp_box.close()
    check(now == Summary(row=True, sheet=False, group_by="nr_dz"), f"okno: zapis ustawień podsumowania ({now})")
    again = HandoverDialog(iface, None, layer=gp)
    again.reset_settings()
    check(again.summary() == now, "„Przywróć domyślne” nie zmienia ustawień sum warstwy")
    again.close()
    project = QgsProject.instance()
    project.removeMapLayer(gp.id())

    layer_reader.save_choices(layer, [FieldChoice("rok", sum=True)], None, summary=Summary(row=True, sheet=True))
    for lang in ("pl", "en"):
        QgsSettings().setValue(UI_LANGUAGE, lang)
        shot = HandoverDialog(None, None, layer=layer)
        shot.resize(980, 760)
        shot.show()
        QApplication.processEvents()
        shot.grab().save(str(SCREENS / f"okno_sumy_{lang}_qgis{VERSION}.png"))
        shot.close()
    QgsSettings().remove(UI_LANGUAGE)


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
    wait(box.task)
    check(bool(warnings) and layer.customProperty(dialog_module.PREVIOUS, "") == "",
          "nieudany zapis nie zmienia podstawy porównania")
    target["path"] = str(OUT / "v3.xlsx")
    box = HandoverDialog(iface, None, layer=layer)
    box.export()
    check(not box.isVisible() and box.task is not None and iface.messageBar().pushWidget.call_args[0][1]
          == Qgis.MessageLevel.Info, "eksport z okna idzie w tle, z paskiem postępu w komunikatach")
    wait(box.task)
    check(layer.customProperty(dialog_module.PREVIOUS, "") == target["path"],
          "udany zapis staje się podstawą porównania")
    bar = iface.messageBar()
    text = bar.createMessage.call_args[0][1]
    level = bar.pushWidget.call_args[0][1]
    check("•" in text and level == Qgis.MessageLevel.Warning, "uwagi z eksportu widoczne w komunikacie")

    # Przegląd R5, zarzut 2: drugi eksport nie startuje, dopóki trwa pierwszy.
    first = HandoverDialog(iface, None, layer=layer)
    target["path"] = str(OUT / "pierwszy.xlsx")
    first.export()
    second = HandoverDialog(iface, None, layer=layer)
    busy: list[str] = []
    second._warn = busy.append  # type: ignore[method-assign]
    second.export()
    check(busy == [tr(ui_language(), "ui.error.busy")] and second.task is None, "drugi eksport czeka na pierwszy")
    wait(first.task)

    # Zarzut 3: warstwa usunięta przed końcem zadania nie dostaje podstawy porównania (i nic się nie sypie).
    gone = make_layer()
    box = HandoverDialog(iface, None, layer=gone)
    target["path"] = str(OUT / "usunieta_z_okna.xlsx")
    box.export()
    QgsProject.instance().removeMapLayer(gone.id())
    wait(box.task)
    check(box.layer_gone and box.last_result is not None, "warstwa usunięta w trakcie: plik jest, podstawy brak")

    geometry()
    processing_algorithm(iface)
    summaries(iface)
    background(layer)

    # Plik do obejrzenia (zrzut arkuszy robi scripts/render-xlsx.sh).
    export(str(ROOT / "build" / f"przyklad_qgis{VERSION}.xlsx"), pkg2, Look(),
           Sheets(dictionary=True, about=True), previous=v1)
    print("OK")


if __name__ == "__main__":
    main()
    sys.stdout.flush()
    os._exit(0)  # bez sprzątania QGIS przy wyjściu (znane awarie przy exit w trybie testowym)
