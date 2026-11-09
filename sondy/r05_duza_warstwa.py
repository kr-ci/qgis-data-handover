"""Sonda R5: eksport 200 000 obiektów — czas odczytu, zapisu i najdłuższe zamrożenie okna.

Uruchomienie (w obrazie QGIS, z katalogu projektu):
    scripts/qgis-probe.sh 4.2.3 sondy/r05_duza_warstwa.py [sync|task]

Tryb `sync` to eksport tak jak przed R5 (w wątku okna), `task` — przez
zadanie w tle. „Zamrożenie” to najdłuższa przerwa między kolejnymi
tyknięciami zegara co 20 ms: tyle czasu QGIS nie obsługiwał zdarzeń.
"""

from __future__ import annotations

import os
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from qgis.core import QgsFeature, QgsGeometry, QgsPointXY, QgsProject, QgsVectorLayer  # noqa: E402
from qgis.PyQt.QtCore import QTimer  # noqa: E402
from qgis.testing import start_app  # noqa: E402

app = start_app()

from data_handover import layer_reader  # noqa: E402
from data_handover.core import xlsx  # noqa: E402
from data_handover.layer_reader import FieldChoice  # noqa: E402

N = int(os.environ.get("N", "200000"))
MODE = sys.argv[1] if len(sys.argv) > 1 else "sync"


def make_layer() -> QgsVectorLayer:
    uri = ("Point?crs=EPSG:2180&field=fid:integer&field=nr_dz:string&field=teryt:string"
           "&field=pow_m2:double(12,2)&field=przezn:string&field=data_akt:date&field=opis:string")
    layer = QgsVectorLayer(uri, "duza", "memory")
    features = []
    for i in range(N):
        f = QgsFeature(layer.fields())
        f.setAttributes([i, f"{i}/1", "0214011", 1000.0 + i, "MN", None, "działka testowa"])
        f.setGeometry(QgsGeometry.fromPointXY(QgsPointXY(i % 1000, i // 1000)))
        features.append(f)
    layer.dataProvider().addFeatures(features)
    QgsProject.instance().addMapLayer(layer)
    return layer


def main() -> None:
    layer = make_layer()
    choices = [FieldChoice(n) for n in layer.fields().names()]
    path = os.path.join(tempfile.mkdtemp(), "duza.xlsx")
    ticks: list[float] = []
    timer = QTimer()
    timer.timeout.connect(lambda: ticks.append(time.perf_counter()))
    timer.start(20)
    app.processEvents()
    start = time.perf_counter()
    ticks.append(start)
    if MODE == "sync":
        pkg = layer_reader.build_package(layer, choices, personal={})
        read = time.perf_counter() - start
        xlsx.export(path, pkg)
    else:
        from data_handover.export_task import ExportTask

        task = ExportTask("sonda", layer_reader.plan_package(layer, choices, personal={}), path)
        from qgis.core import QgsApplication

        QgsApplication.taskManager().addTask(task)
        while not task.done:
            app.processEvents()
            time.sleep(0.005)
        read = task.read_seconds
        assert task.result is not None, task.error
    total = time.perf_counter() - start
    app.processEvents()
    ticks.append(time.perf_counter())
    freeze = max(b - a for a, b in zip(ticks, ticks[1:]))
    size = os.path.getsize(path) / 1e6
    print(f"tryb={MODE} obiekty={N} odczyt={read:.1f}s razem={total:.1f}s "
          f"najdluzsze_zamrozenie={freeze:.2f}s plik={size:.1f}MB")


if __name__ == "__main__":
    main()
    sys.stdout.flush()
    os._exit(0)
