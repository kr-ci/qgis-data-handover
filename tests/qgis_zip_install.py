"""Paczka ZIP działa jak po „Zainstaluj z pliku ZIP” (uruchamiany przez scripts/qgis-test.sh).

Rozpakowuje build/data_handover-<wersja>.zip do pustego katalogu wtyczek,
sprawdza metadane i zgodność wersji tak jak instalator QGIS, a potem wczytuje
i uruchamia wtyczkę przez qgis.utils — z paczki, nie z katalogu źródeł.
scripts/qgis-test.sh uruchamia go dwa razy i mówi wprost, czego oczekuje (zmienna
OPENPYXL): `bundled` — w systemie nie ma openpyxl, eksport musi użyć kopii z paczki
(data_handover/ext); `system` — openpyxl jest w systemie i ma pierwszeństwo.
"""

from __future__ import annotations

import importlib.util
import os
import stat
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.append("/usr/share/qgis/python")  # pyplugin_installer

from qgis.core import Qgis  # noqa: E402
from qgis.testing import start_app  # noqa: E402
from qgis.testing.mocked import get_iface  # noqa: E402

app = start_app()

import qgis.utils  # noqa: E402
from pyplugin_installer.version_compare import isCompatible  # noqa: E402

REQUIRED = ("name", "qgisMinimumVersion", "description", "about", "version", "author", "email",
            "repository", "tracker", "homepage", "changelog")


def check(condition: bool, message: str) -> None:
    print(("  ok  " if condition else "  BŁĄD ") + message)
    if not condition:
        raise AssertionError(message)


def main() -> None:
    version = next(line.split("=", 1)[1].strip()
                   for line in (ROOT / "data_handover" / "metadata.txt").read_text(encoding="utf-8").splitlines()
                   if line.startswith("version="))
    package = ROOT / "build" / f"data_handover-{version}.zip"
    with zipfile.ZipFile(package) as archive:
        names = archive.namelist()
        modes = {info.filename: info.external_attr >> 16 for info in archive.infolist()}
        check(all(n.startswith("data_handover/") for n in names), "ZIP: jeden katalog najwyższego poziomu")
        check({"data_handover/__init__.py", "data_handover/metadata.txt", "data_handover/LICENSE"} <= set(names),
              "ZIP: obowiązkowe __init__.py, metadata.txt, LICENSE")
        check(not any("__pycache__" in n or n.endswith(".pyc") or "/." in n for n in names),
              "ZIP: bez __pycache__, .pyc i plików ukrytych")
        check(not any(mode & (stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH) for mode in modes.values()),
              "ZIP: bez plików wykonywalnych")
        plugins = tempfile.mkdtemp()
        archive.extractall(plugins)

    source = os.environ.get("OPENPYXL", "")  # skąd ma pochodzić openpyxl w tym przebiegu
    check(source in ("bundled", "system"), f"OPENPYXL=bundled|system ustawione przez skrypt ({source!r})")
    system_openpyxl = importlib.util.find_spec("openpyxl") is not None  # przed dodaniem paczki do ścieżki
    check(system_openpyxl == (source == "system"),
          f"środowisko testu: openpyxl w systemie {'jest' if system_openpyxl else 'brak'} (oczekiwano {source})")
    sys.path.insert(0, plugins)
    qgis.utils.plugin_paths = [plugins]
    qgis.utils.iface = get_iface()
    qgis.utils.updateAvailablePlugins()
    check("data_handover" in qgis.utils.available_plugins, "QGIS widzi wtyczkę w katalogu wtyczek")
    meta = {key: qgis.utils.pluginMetadata("data_handover", key) for key in REQUIRED}
    missing = [key for key, value in meta.items() if value in ("", "__error__")]
    check(not missing, f"metadane: wymagane pola obecne ({missing})")
    check(meta["version"] == version and qgis.utils.pluginMetadata("data_handover", "experimental") == "False",
          f"metadane: wersja {version}, nie eksperymentalna")
    qgis_version = Qgis.version().split("-")[0]
    check(isCompatible(qgis_version, meta["qgisMinimumVersion"],
                       qgis.utils.pluginMetadata("data_handover", "qgisMaximumVersion")),
          f"instalator QGIS {qgis_version} uznaje wtyczkę za zgodną")
    check(qgis.utils.loadPlugin("data_handover") and qgis.utils.startPlugin("data_handover"),
          "wtyczka z paczki wczytana i uruchomiona (initGui)")
    module = sys.modules["data_handover"]
    check(os.path.realpath(module.__file__).startswith(os.path.realpath(plugins)),
          "wczytana z rozpakowanej paczki, nie ze źródeł")
    from data_handover.core.i18n import LANGUAGES, TEXT
    from data_handover.dialog import HandoverDialog

    window = HandoverDialog(None, None)
    check(window.ui_lang_combo.count() == len(LANGUAGES) == 10
          and all(len(TEXT[lang]) == len(TEXT["en"]) for lang in LANGUAGES),
          "okno z paczki otwiera się, wszystkie 10 języków w ZIP")
    window.close()

    # Eksport kodem z paczki: wszystkie arkusze, z tytułem, sumami i porównaniem z poprzednim plikiem.
    from openpyxl import load_workbook
    from qgis.core import QgsFeature, QgsGeometry, QgsPointXY, QgsVectorLayer

    from data_handover import layer_reader
    from data_handover.core.i18n import tr
    from data_handover.core.model import Look, Sheets, Summary
    from data_handover.core.xlsx import export

    layer = QgsVectorLayer("Point?crs=EPSG:2180&field=nr:string&field=pow:double(10,2)", "zip", "memory")
    for nr, area in (("1/1", 10.5), ("1/2", 20.25)):
        feature = QgsFeature(layer.fields())
        feature.setAttributes([nr, area])
        feature.setGeometry(QgsGeometry.fromPointXY(QgsPointXY(500000, 500000)))
        layer.dataProvider().addFeature(feature)
    choices = [layer_reader.FieldChoice(n) for n in layer.fields().names()]
    out = tempfile.mkdtemp()
    first, second = os.path.join(out, "v1.xlsx"), os.path.join(out, "v2.xlsx")
    for path, previous in ((first, None), (second, first)):
        pkg = layer_reader.build_package(layer, choices, key="nr", title="ZIP", language="pl",
                                         summary=Summary(row=True, sheet=True))
        export(path, pkg, Look(), Sheets(dictionary=True, about=True), previous=previous)
    expected = [tr("pl", f"sheet.{s}") for s in ("data", "summary", "dictionary", "about", "changes")]
    names = load_workbook(second).sheetnames
    check(names[:5] == expected, f"eksport kodem z paczki: wszystkie arkusze ({names})")
    import openpyxl

    from_package = os.path.realpath(openpyxl.__file__).startswith(os.path.realpath(plugins))
    origin = "dołączona do paczki" if source == "bundled" else "z systemu (ma pierwszeństwo)"
    check(from_package == (source == "bundled"), f"openpyxl {origin}: {openpyxl.__file__}")
    check(qgis.utils.unloadPlugin("data_handover"), "wyłączenie wtyczki bez błędów")
    print("OK")


if __name__ == "__main__":
    main()
    sys.stdout.flush()
    os._exit(0)  # bez sprzątania QGIS przy wyjściu (jak w tests/qgis_smoke.py)
