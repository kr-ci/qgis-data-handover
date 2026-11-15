"""Sonda: prawdziwy menedżer wtyczek QGIS (np. 3.44.7 z qgis.org dla Ubuntu 24.04, z QtWebKit).

Otwiera okno Wtyczki, wybiera kolejne wtyczki z listy i zmienia szerokość okna jak na
zrzutach właściciela (1150, 1040, 940 px). Dla każdej szerokości: zrzut okna i to, czy
panel opisu ma poziomy suwak (opis szerszy niż panel — ikona po prawej znika z widoku).

Uruchamia QGIS: qgis --nologo --code sondy/menedzer_qgis.py (pełne polecenie Docker
w sondy/wyniki/menedzer_opis.txt). Wtyczki do porównania: zmienna DH_PLUGINS, nazwy
z metadata.txt rozdzielone „;”.
"""

import os
import traceback

from qgis.PyQt.QtCore import QEventLoop, Qt, QThread, QTimer
from qgis.PyQt.QtWidgets import QApplication, QLineEdit, QListView, QTextBrowser
from qgis.utils import iface

OUT = "/src/build/sonda_menedzer"
WIDTHS = (1150, 1040, 940)
PLUGINS = os.environ.get("DH_PLUGINS", "Data Handover").split(";")


def settle(steps=40):
    for _ in range(steps):
        QApplication.processEvents(QEventLoop.ProcessEventsFlag.AllEvents, 50)
        QThread.msleep(25)


def manager():
    for widget in QApplication.topLevelWidgets():
        plugins = widget.findChild(QListView, "vwPlugins")
        if plugins is not None and plugins.isVisible():
            return plugins.window()  # okno Wtyczki, nie główne okno QGIS (jego rodzic)
    return None


def details_view(dialog):
    """Panel opisu: QWebView (QGIS z QtWebKit) albo QTextBrowser (bez QtWebKit)."""
    try:
        from qgis.PyQt.QtWebKitWidgets import QWebView

        view = dialog.findChild(QWebView, "wvDetails")
        if view is not None:
            return view
    except ImportError:
        pass
    return dialog.findChild(QTextBrowser, "wvDetails")


def horizontal_overflow(details):
    """Ile pikseli opisu jest poza panelem w poziomie."""
    if isinstance(details, QTextBrowser):
        return details.horizontalScrollBar().maximum()
    return details.page().mainFrame().scrollBarMaximum(Qt.Orientation.Horizontal)


def run():
    dialog = manager()
    if dialog is None:
        QTimer.singleShot(1000, run)
        return
    try:
        measure(dialog)
    except Exception:  # QGIS połyka wyjątki z zegara — wypisujemy je sami
        traceback.print_exc()
    os._exit(0)


def measure(dialog):
    os.makedirs(OUT, exist_ok=True)
    view = dialog.findChild(QListView, "vwPlugins")
    details = details_view(dialog)
    search = dialog.findChild(QLineEdit, "leFilter")
    print(f"silnik opisu: {type(details).__name__}", flush=True)
    for name in PLUGINS:
        search.setText(name)
        settle()
        model = view.model()
        rows = [model.index(r, 0) for r in range(model.rowCount())]
        match = [i for i in rows if i.data() == name]
        if not match:
            print(f"brak wtyczki {name!r} na liście: {[i.data() for i in rows]}", flush=True)
            continue
        view.setCurrentIndex(match[0])
        settle()
        if not isinstance(details, QTextBrowser):  # prawdziwy HTML opisu do dalszej analizy
            with open(f"{OUT}/qgis_{name.replace(' ', '_')}.html", "w", encoding="utf-8") as out:
                out.write(details.page().mainFrame().toHtml())
        for width in WIDTHS:
            dialog.resize(width, 830)
            settle()
            tag = name.replace(" ", "_")
            dialog.grab().save(f"{OUT}/qgis_{tag}_{width}.png")
            print(f"tryb=menedzer_qgis wtyczka={name!r} okno={width} opis={details.width()} px "
                  f"poza_panelem={horizontal_overflow(details)} px", flush=True)


QTimer.singleShot(3000, run)
iface.actionManagePlugins().trigger()  # okno modalne: run() działa w jego pętli zdarzeń
