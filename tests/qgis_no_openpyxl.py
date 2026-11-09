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

from qgis.core import QgsApplication, QgsProcessingContext, QgsProcessingFeedback  # noqa: E402
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

# Algorytm Processing: zarejestrowany, a uruchomiony — ta sama instrukcja zamiast tracebacku.
alg = QgsApplication.processingRegistry().algorithmById("dataHandover:package")
assert alg is not None, "initGui nie zarejestrował algorytmu"
errors: list[str] = []


class Feedback(QgsProcessingFeedback):
    def reportError(self, error: str, fatalError: bool = False) -> None:  # noqa: N802, N803
        errors.append(error)


ok = alg.create().prepare({"INPUT": None, "OUTPUT": "/tmp/x.xlsx"}, QgsProcessingContext(), Feedback())
assert not ok and any("openpyxl" in e for e in errors), f"brak instrukcji o openpyxl: {errors}"
print("ok  algorytm Processing bez openpyxl: zarejestrowany, nie startuje")
plugin.unload()
print("ok  unload bez openpyxl")
sys.stdout.flush()
os._exit(0)
