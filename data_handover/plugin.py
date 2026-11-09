"""Rejestracja wtyczki w QGIS: przycisk, menu Wtyczki i menu kontekstowe warstwy."""

from __future__ import annotations

import json
import os
from typing import Any, Optional

from qgis.core import Qgis, QgsApplication, QgsSettings
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QMessageBox

try:  # Qt6: QAction w QtGui; Qt5: w QtWidgets
    from qgis.PyQt.QtGui import QAction
except ImportError:  # pragma: no cover - zależne od wersji Qt
    from qgis.PyQt.QtWidgets import QAction

from .core.i18n import supported, tr

MENU = "&Data Handover"


# Ustawienia okna w profilu użytkownika QGIS. Tu, nie w dialog.py: initGui()
# i algorytm Processing nie mogą importować niczego, co ciągnie openpyxl,
# bo bez niego wtyczka nie wczytałaby się wcale.
SETTINGS = "data_handover/"
UI_LANGUAGE = SETTINGS + "ui_language"


def ui_language() -> str:
    """Język okna wybrany w stopce okna; domyślnie angielski."""
    return supported(str(QgsSettings().value(UI_LANGUAGE, "en") or "en"))


def saved_json(name: str) -> Any:
    """Ustawienie okna zapisane jako JSON (wygląd, arkusze); None, gdy brak albo uszkodzone."""
    try:
        return json.loads(str(QgsSettings().value(SETTINGS + name, "") or "null"))
    except ValueError:
        return None


class DataHandoverPlugin:
    def __init__(self, iface: Any) -> None:
        self.iface = iface
        self.action: Optional[QAction] = None
        self.layer_action: Optional[QAction] = None
        self.provider: Any = None

    def initProcessing(self) -> None:  # noqa: N802 - nazwa wymagana przez QGIS (także qgis_process)
        from .processing_provider import Provider

        self.provider = Provider()
        QgsApplication.processingRegistry().addProvider(self.provider)

    def initGui(self) -> None:  # noqa: N802 - nazwa wymagana przez QGIS
        self.initProcessing()
        lang = ui_language()
        icon = QIcon(os.path.join(os.path.dirname(__file__), "icon.svg"))
        window = self.iface.mainWindow()
        self.action = QAction(icon, tr(lang, "ui.menu"), window)
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu(MENU, self.action)
        self.layer_action = QAction(icon, tr(lang, "ui.menu"), window)
        self.layer_action.triggered.connect(self.run)
        self.iface.addCustomActionForLayerType(self.layer_action, "", Qgis.LayerType.Vector, True)

    def unload(self) -> None:
        if self.provider is not None:
            QgsApplication.processingRegistry().removeProvider(self.provider)
        if self.action is not None:
            self.iface.removePluginMenu(MENU, self.action)
            self.iface.removeToolBarIcon(self.action)
        if self.layer_action is not None:
            self.iface.removeCustomActionForLayerType(self.layer_action)

    def run(self) -> None:
        try:
            import openpyxl  # noqa: F401
        except ImportError:
            lang = ui_language()
            QMessageBox.critical(self.iface.mainWindow(), "Data Handover", tr(lang, "ui.error.openpyxl"))
            return
        from .dialog import HandoverDialog

        layer, tab, geometry = self.iface.activeLayer(), 0, None
        while True:  # zmiana języka okna = to samo okno zbudowane od nowa w nowym języku
            dialog = HandoverDialog(self.iface, self.iface.mainWindow(), layer=layer)
            dialog.tabs.setCurrentIndex(tab)
            if geometry is not None:
                dialog.restoreGeometry(geometry)
            dialog.exec()
            if not dialog.relaunch:
                break
            layer, tab, geometry = dialog.layer, dialog.tabs.currentIndex(), dialog.saveGeometry()
        for action in (self.action, self.layer_action):
            if action is not None:
                action.setText(tr(ui_language(), "ui.menu"))
