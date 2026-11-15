"""Sonda: opis wtyczki w menedżerze wtyczek QGIS — gdzie ląduje ikona przy zmianie szerokości panelu.

QGIS bez QtWebKit (np. obrazy qgis/qgis, pakiety dla Ubuntu) pokazuje opis w QTextBrowser
(QgsWebView). HTML jest odtworzony z QgsPluginManager::showPluginDetails (QGIS 3.44,
src/app/pluginmanager/qgspluginmanager.cpp): ikona to <img style="float:right; max-width:64px;
max-height:64px;"> w pierwszej komórce tabeli, przed tytułem i opisem.

QGIS 3.x z QtWebKit (pakiety qgis.org dla Ubuntu) rysuje ten sam HTML silnikiem WebKit:
  docker run --rm -v "$PWD:/src" -w /src ubuntu:24.04 bash -c 'apt-get update -qq &&
    apt-get install -y -qq python3-pyqt5 python3-pyqt5.qtwebkit xvfb xauth fonts-dejavu >/dev/null &&
    timeout 300 xvfb-run -a -s "-screen 0 1600x1200x24" python3 -u sondy/menedzer_opis.py webkit'

Qt5 (QGIS 3.x) bez Dockera:
  QT_QPA_PLATFORM=offscreen uv run --no-project --python 3.12 --with PyQt5 python sondy/menedzer_opis.py qt5

Uruchomienie w QGIS 4.2.3 (zrzuty w build/sonda_menedzer/):
  docker run --rm -v "$PWD:/src" -w /src qgis/qgis:4.2.3 \
    timeout 300 xvfb-run -a -s "-screen 0 1600x1200x24" python3 -u sondy/menedzer_opis.py
"""

from __future__ import annotations

import configparser
import html
import os
import sys
from pathlib import Path

try:
    from qgis.PyQt.QtCore import QThread
    from qgis.PyQt.QtWidgets import QApplication, QTextBrowser, QWidget
except ImportError:  # bez QGIS: sam Qt5 (pip PyQt5 albo pakiety Ubuntu)
    from PyQt5.QtCore import QThread  # type: ignore[no-redef]
    from PyQt5.QtWidgets import QApplication, QTextBrowser, QWidget  # type: ignore[no-redef]

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "build" / "sonda_menedzer"
STYLE = (
    "<style> body { color:#000000; background-color:#ffffff; } body, table { padding:0px; margin:0px;"
    " font-family:Verdana, Sans-serif; font-size:10pt; } td { vertical-align:top; }"
    " td.key { font-weight: bold; white-space:nowrap; padding-right:10px; text-align:right; } </style>"
)


def details(meta: dict[str, str], icon: str, changelog: bool) -> str:
    """Jak showPluginDetails dla zainstalowanej wtyczki (bez głosowania, bez komunikatów)."""
    page = STYLE + "<body><table cellspacing='4' width='100%'><tr><td colspan='2'>"
    page += f"<img src=\"file://{icon}\" style=\"float:right;max-width:64px;max-height:64px;\">"
    page += f"<h1>{meta['name']}</h1><h3>{meta['description']}</h3>"
    page += meta["about"].replace("\n", "<br/>") + "<br/><br/>"
    page += "</td></tr><tr><td width='1%'> </td><td width='99%'> </td></tr>"
    page += f"<tr><td class='key'>Category </td><td>{meta['category']}</td></tr>"
    tags = ", ".join(f"<a href='rpc2://search.tag/{t.strip()}/'>{t.strip()}</a>" for t in meta["tags"].split(","))
    page += f"<tr><td class='key'>Tags </td><td>{tags}</td></tr>"
    page += (f"<tr><td class='key'>More info </td><td><a href='{meta['homepage']}'>homepage</a> &nbsp; "
             f"<a href='{meta['tracker']}'>bug tracker</a> &nbsp; <a href='{meta['repository']}'>code repository</a>"
             "</td></tr>")
    page += f"<tr><td class='key'>Author </td><td>{meta['author']}</td></tr>"
    page += f"<tr><td class='key'>Installed version </td><td>{meta['version']}</td></tr>"
    if changelog and meta.get("changelog"):
        page += f"<tr><td class='key'>Changelog </td><td>{meta['changelog'].strip()}</td></tr>"
    return page + "</table></body>"


def settle(app: QApplication) -> None:
    for _ in range(20):
        app.processEvents()
        QThread.msleep(25)


def make_view(variant: str, app: QApplication, width: int) -> QWidget:
    if variant.startswith("webkit"):
        from PyQt5.QtWebKitWidgets import QWebView

        view: QWidget = QWebView()
    else:
        view = QTextBrowser()
    view.resize(width, 900)
    view.show()
    return view


def show(view: QWidget, page: str, app: QApplication) -> None:
    loaded: list[bool] = []
    if hasattr(view, "loadFinished"):  # WebKit wczytuje stronę asynchronicznie, czasem już w setHtml
        view.loadFinished.connect(loaded.append)
    view.setHtml(page)
    for _ in range(200):
        if loaded or not hasattr(view, "loadFinished"):
            break
        settle(app)
    settle(app)


def main() -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    OUT.mkdir(parents=True, exist_ok=True)
    parser = configparser.ConfigParser(interpolation=None)
    parser.read(ROOT / "data_handover" / "metadata.txt", encoding="utf-8")
    meta = {k: html.escape(v, quote=False) for k, v in parser["general"].items()}
    icon = str(ROOT / "data_handover" / "icon.svg")
    variant = sys.argv[1] if len(sys.argv) > 1 else "obecny"

    page = details(meta, icon, changelog=True)
    view = make_view(variant, app, 700)
    show(view, page, app)
    view.grab().save(str(OUT / f"{variant}_1_otwarte_700.png"))
    # Zmiana szerokości już otwartego panelu (przeciągnięcie okna albo rozdzielacza).
    for step, width in ((2, 1100), (3, 450), (4, 700)):
        view.resize(width, 900)
        settle(app)
        view.grab().save(str(OUT / f"{variant}_{step}_po_zmianie_{width}.png"))
    # Dla porównania: panel od razu w szerokości 450 i 1100.
    for width in (450, 1100):
        fresh = make_view(variant, app, width)
        show(fresh, page, app)
        fresh.grab().save(str(OUT / f"{variant}_5_od_razu_{width}.png"))
    print(f"tryb=menedzer wariant={variant} zrzuty={OUT}")
    sys.stdout.flush()
    os._exit(0)


if __name__ == "__main__":
    main()
