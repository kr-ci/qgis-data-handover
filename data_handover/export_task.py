"""Eksport w tle: zadanie QGIS (`QgsTask`) z postępem i anulowaniem.

Pasek postępu i przycisk anulowania pokazuje QGIS w pasku stanu (menedżer
zadań), a okno wtyczki dodatkowo w pasku komunikatów. Wynik obsługuje
`on_finished` — w wątku okna, bo tam QGIS wywołuje `finished()`.
"""

from __future__ import annotations

import time
import traceback
from typing import Callable, Optional

from qgis.core import QgsApplication, QgsTask

from .core import xlsx
from .core.model import Look, Sheets
from .layer_reader import PackagePlan

# QGIS nie trzyma pythonowego obiektu zadania: bez tej referencji
# odśmiecacz usunąłby je w trakcie pracy.
_RUNNING: set[ExportTask] = set()


class ExportTask(QgsTask):
    def __init__(self, description: str, plan: PackagePlan, path: str, look: Optional[Look] = None,
                 sheets: Optional[Sheets] = None, previous: Optional[str] = None) -> None:
        super().__init__(description)
        self.plan, self.path, self.look, self.sheets, self.previous = plan, path, look, sheets, previous
        self.result: Optional[xlsx.ExportResult] = None
        self.error: Optional[Exception] = None
        self.trace = ""
        self.done = False
        self.read_seconds = 0.0
        self.on_finished: Optional[Callable[[ExportTask], None]] = None

    def run(self) -> bool:
        """Wątek tła: tu nie wolno dotykać warstwy ani widżetów."""
        try:
            start = time.perf_counter()
            # 10% paska na odczyt, 90% na zapis: w pomiarze R5 (200 000 obiektów)
            # odczyt trwał 2,5 s, a cały eksport ok. 50 s (sondy/r05_duza_warstwa.py).
            pkg = self.plan.read(lambda p: self.setProgress(p / 10), self.isCanceled)
            self.read_seconds = time.perf_counter() - start
            self.result = xlsx.export(self.path, pkg, self.look, self.sheets, previous=self.previous,
                                      progress=lambda p: self.setProgress(10 + p * 0.9), canceled=self.isCanceled)
            return True
        except xlsx.CanceledError:
            return False
        except Exception as exc:  # zapis pliku: brak uprawnień, plik otwarty w Excelu, błąd danych
            self.error, self.trace = exc, traceback.format_exc()
            return False

    def finished(self, result: bool) -> None:  # noqa: ARG002 - podpis wymagany przez QGIS
        self.done = True
        _RUNNING.discard(self)
        if self.on_finished is not None:
            self.on_finished(self)


def busy() -> bool:
    """Trwa eksport. Drugi naraz mógłby skończyć się wcześniej i zostać nadpisany starszym."""
    return bool(_RUNNING)


def start(task: ExportTask) -> None:
    _RUNNING.add(task)
    QgsApplication.taskManager().addTask(task)
