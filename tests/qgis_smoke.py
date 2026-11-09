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
import os
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
    QgsApplication,
    QgsEditorWidgetSetup,
    QgsFeature,
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
from qgis.PyQt.QtCore import QDate, QSettings, QTimer  # noqa: E402
from qgis.PyQt.QtWidgets import QApplication  # noqa: E402
from qgis.testing import start_app  # noqa: E402

app = start_app()

from openpyxl import load_workbook  # noqa: E402

from data_handover import classFactory, layer_reader  # noqa: E402
from data_handover.core.i18n import LANGUAGES, tr  # noqa: E402
from data_handover.core.model import HEADER_ROW, THEMES, Look, Sheets, look_with_theme, to_dict  # noqa: E402
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
                              c.comment.text if c.comment is not None else None))
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
    layer_reader.save_choices(layer, choices, "nr_dz")
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

    processing_algorithm(iface)
    background(layer)

    # Plik do obejrzenia (zrzut arkuszy robi scripts/render-xlsx.sh).
    export(str(ROOT / "build" / f"przyklad_qgis{VERSION}.xlsx"), pkg2, Look(),
           Sheets(dictionary=True, about=True), previous=v1)
    print("OK")


if __name__ == "__main__":
    main()
    sys.stdout.flush()
    os._exit(0)  # bez sprzątania QGIS przy wyjściu (znane awarie przy exit w trybie testowym)
