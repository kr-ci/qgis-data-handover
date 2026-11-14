"""Sonda R11: czy okno mieści się na ekranie (rozmiar, minimum, szerokości zakładek) w każdym języku.

Uruchomienie z wybranym ekranem i skalowaniem, np.:
  docker run --rm -v "$PWD:/src" -w /src qgis/qgis:4.2.3 bash -c \
    'apt-get update -qq && apt-get install -y -qq python3-openpyxl >/dev/null;
     QT_SCALE_FACTOR=1.25 xvfb-run -a -s "-screen 0 1366x768x24" python3 -u sondy/r11_okno.py'
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from qgis.testing import start_app  # noqa: E402

app = start_app()

from qgis.core import QgsSettings  # noqa: E402
from qgis.PyQt.QtWidgets import QApplication  # noqa: E402

from data_handover.core.i18n import LANGUAGES  # noqa: E402
from data_handover.dialog import HandoverDialog  # noqa: E402
from data_handover.plugin import UI_LANGUAGE  # noqa: E402
from tests.qgis_smoke import make_layer  # noqa: E402

layer = make_layer()
area = QApplication.primaryScreen().availableGeometry()
print(f"ekran (logiczne px): {area.width()}×{area.height()}, skala {QApplication.primaryScreen().devicePixelRatio()}")
for lang in LANGUAGES:
    QgsSettings().setValue(UI_LANGUAGE, lang)
    box = HandoverDialog(None, None, layer=layer)
    box.show()
    QApplication.processEvents()
    hint = box.minimumSizeHint()
    pages = [box.tabs.widget(i).minimumSizeHint() for i in range(box.tabs.count())]
    print(f"tryb={lang} okno={box.width()}×{box.height()} minimum={max(hint.width(), box.minimumWidth())}×"
          f"{max(hint.height(), box.minimumHeight())} zakładki=" + " ".join(f"{p.width()}×{p.height()}" for p in pages))
    box.close()
sys.stdout.flush()
os._exit(0)
