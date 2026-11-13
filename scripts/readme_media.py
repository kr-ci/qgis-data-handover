"""Zrzuty okna i przykładowy plik do README (uruchamia scripts/readme-media.sh w QGIS w Dockerze).

Wynik w build/readme/: okno_<zakładka>.png, parcels_v1.xlsx i parcels.xlsx
(drugi z arkuszem „Changes” względem pierwszego).
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from qgis.core import (  # noqa: E402
    QgsEditorWidgetSetup,
    QgsFeature,
    QgsGeometry,
    QgsProject,
    QgsVectorFileWriter,
    QgsVectorLayer,
)
from qgis.PyQt.QtCore import QDate  # noqa: E402
from qgis.PyQt.QtWidgets import QApplication  # noqa: E402
from qgis.testing import start_app  # noqa: E402
from qgis.testing.mocked import get_iface  # noqa: E402

app = start_app()

import data_handover.dialog as dialog_module  # noqa: E402
from data_handover import layer_reader  # noqa: E402
from data_handover.core.model import Summary  # noqa: E402
from data_handover.dialog import HandoverDialog  # noqa: E402
from data_handover.layer_reader import FieldChoice  # noqa: E402

OUT = ROOT / "build" / "readme"
OUT.mkdir(parents=True, exist_ok=True)

# Działki w Warszawie (EPSG:2180): numer, obręb, użytek, data, właściciel, e-mail, szerokość × wysokość [m].
PARCELS = [
    ("40/1", "0401", "MN", QDate(2024, 5, 3), "Jan Kowalski", "jan.kowalski@example.com", 24, 35),
    ("40/2", "0401", "MN", QDate(2024, 5, 3), "Anna Nowak", "anna.nowak@example.com", 22, 34),
    ("41", "0401", "U", QDate(2023, 11, 20), "Piotr Wiśniewski", "p.wisniewski@example.com", 48, 60),
    ("42/5", "0401", "ZL", None, None, None, 120, 210),
    ("43", "0402", "R", QDate(2022, 3, 14), "Maria Wójcik", "maria.wojcik@example.com", 95, 140),
    ("44/1", "0402", "MN", QDate(2025, 1, 8), "Tomasz Kamiński", "t.kaminski@example.com", 26, 38),
    ("44/2", "0402", "MN", QDate(2025, 1, 8), "Ewa Lewandowska", "ewa.lewandowska@example.com", 25, 36),
    ("45", "0402", "U", QDate(2021, 9, 30), "Krzysztof Zieliński", "k.zielinski@example.com", 60, 45),
    ("46/3", "0403", "R", QDate(2020, 6, 1), "Agnieszka Szymańska", "a.szymanska@example.com", 110, 180),
    ("47", "0403", "ZL", None, None, None, 150, 160),
]
ALIASES = (("parcel_id", "Parcel ID"), ("parcel_no", "Parcel no."), ("precinct", "Precinct"),
           ("land_use", "Land use"), ("updated", "Last updated"), ("owner", "Owner"), ("email", "E-mail"))
LAND_USE = {"map": [{"Residential": "MN"}, {"Services": "U"}, {"Forest": "ZL"}, {"Agricultural": "R"}]}


def make_layer() -> QgsVectorLayer:
    """Warstwa z GeoPackage (/data/parcels.gpkg w kontenerze), żeby „Source” w „About” wyglądało jak u autora."""
    uri = ("Polygon?crs=EPSG:2180&field=parcel_id:string&field=parcel_no:string&field=precinct:string"
           "&field=land_use:string&field=updated:date&field=owner:string&field=email:string")
    memory = QgsVectorLayer(uri, "parcels", "memory")
    x = 637000.0
    features = []
    for no, precinct, use, date, owner, email, w, h in PARCELS:
        f = QgsFeature(memory.fields())
        f.setAttributes([f"146510_8.{precinct}.{no}", no, precinct, use, date, owner, email])
        f.setGeometry(QgsGeometry.fromWkt(
            f"POLYGON(({x} 486000,{x + w} 486000,{x + w} {486000 + h},{x} {486000 + h},{x} 486000))"))
        features.append(f)
        x += w + 5
    assert memory.dataProvider().addFeatures(features)[0]
    Path("/data").mkdir(exist_ok=True)
    options = QgsVectorFileWriter.SaveVectorOptions()
    options.driverName, options.layerName = "GPKG", "parcels"
    error = QgsVectorFileWriter.writeAsVectorFormatV3(memory, "/data/parcels.gpkg",
                                                      QgsProject.instance().transformContext(), options)
    assert error[0] == QgsVectorFileWriter.WriterError.NoError, error
    layer = QgsVectorLayer("/data/parcels.gpkg|layername=parcels", "parcels", "ogr")
    assert layer.isValid() and layer.featureCount() == len(PARCELS)
    for name, alias in ALIASES:
        layer.setFieldAlias(layer.fields().indexOf(name), alias)
    layer.setEditorWidgetSetup(layer.fields().indexOf("land_use"), QgsEditorWidgetSetup("ValueMap", LAND_USE))
    QgsProject.instance().addMapLayer(layer)
    descriptions = {
        "parcel_id": "Cadastral parcel identifier (TERYT unit, precinct, number)",
        "land_use": "Zoning designation from the local plan",
    }
    choices = [FieldChoice(n, export=n not in ("fid", "owner"), description=descriptions.get(n, ""),
                           protect="mask" if n == "email" else "")
               for n in layer.fields().names()[1:] + ["fid"]]  # fid z GeoPackage na koniec listy
    layer_reader.save_choices(layer, choices, "parcel_id",
                              summary=Summary(row=True, sheet=True, group_by="land_use"))
    return layer


def export(iface: object, layer: QgsVectorLayer, path: Path, previous: str = "", shots: bool = False) -> None:
    dialog_module.QFileDialog.getSaveFileName = lambda *args, **kwargs: (str(path), "")
    box = HandoverDialog(iface, None, layer=layer)
    box.title_edit.setText("Parcels — Ochota precincts 0401–0403")
    box.author_edit.setText("GIS Department")
    box.measure_check.setChecked(True)
    box.wgs84_check.setChecked(True)
    box.link_check.setChecked(True)
    box.dictionary_box.setChecked(True)
    box.about_box.setChecked(True)
    if previous:
        box.changes_check.setChecked(True)
        box.key_combo.setField("parcel_id")
        box.previous_widget.setFilePath(previous)
    if shots:
        box.resize(1000, 800)
        box.show()
        for index, name in enumerate(("data", "look", "sheets", "changes")):
            box.tabs.setCurrentIndex(index)
            QApplication.processEvents()
            box.grab().save(str(OUT / f"window_{name}.png"))
        box.tabs.setCurrentIndex(0)
    box.export()
    deadline = time.monotonic() + 60
    while box.task is not None and box.last_result is None and time.monotonic() < deadline:
        QApplication.processEvents()
        time.sleep(0.05)
    assert box.last_result is not None, box.task.error if box.task is not None else "brak zadania"


def main() -> None:
    iface = get_iface()
    layer = make_layer()
    first = OUT / "parcels_v1.xlsx"
    export(iface, layer, first)
    # Druga wersja: zmiana przeznaczenia jednej działki i nowa data — do arkusza „Changes”.
    layer.startEditing()
    feature = next(f for f in layer.getFeatures() if f["parcel_no"] == "43")
    layer.changeAttributeValue(feature.id(), layer.fields().indexOf("land_use"), "MN")
    layer.changeAttributeValue(feature.id(), layer.fields().indexOf("updated"), QDate(2026, 10, 2))
    layer.commitChanges()
    export(iface, layer, OUT / "parcels.xlsx", previous=str(first), shots=True)
    print("tryb=readme")


if __name__ == "__main__":
    main()
    sys.stdout.flush()
