"""Test w prawdziwym QGIS: wtyczka bez biblioteki openpyxl.

QGIS wywołuje initGui() przy starcie, zanim użytkownik cokolwiek kliknie.
Jeśli initGui() importuje (choćby pośrednio) openpyxl, wtyczka w ogóle się
nie wczytuje i użytkownik widzi traceback zamiast instrukcji instalacji.
Ten test blokuje import openpyxl i sprawdza, że wtyczka się wczytuje, a po
kliknięciu pokazuje komunikat z instrukcją.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.modules["openpyxl"] = None  # type: ignore[assignment]  # każdy import openpyxl rzuci ImportError

from qgis.PyQt.QtWidgets import QMessageBox  # noqa: E402
from qgis.testing import start_app  # noqa: E402
from qgis.testing.mocked import get_iface  # noqa: E402

app = start_app()

from data_handover import classFactory  # noqa: E402

shown: list[str] = []
QMessageBox.critical = lambda parent, title, text, *args: shown.append(text)  # type: ignore[method-assign]

plugin = classFactory(get_iface())
plugin.initGui()
assert plugin.action is not None, "initGui nie utworzył przycisku"
print("ok  initGui bez openpyxl")
plugin.run()
assert len(shown) == 1 and "openpyxl" in shown[0], f"brak komunikatu o openpyxl: {shown}"
print("ok  run() pokazuje instrukcję instalacji openpyxl")
plugin.unload()
print("ok  unload bez openpyxl")
sys.stdout.flush()
os._exit(0)
