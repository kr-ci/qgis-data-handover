"""Okno na każdym monitorze (R11), w prawdziwym QGIS.

`scripts/qgis-test.sh` uruchamia ten test kilka razy, na ekranach Xvfb o różnej
wielkości i ze skalowaniem (QT_SCALE_FACTOR). Sprawdza, że okno mieści się na
ekranie, przewija się zamiast ściskać, zapamiętuje rozmiar, rozdzielacz
i powiększenie podglądu, a „Przywróć domyślne” je zeruje. Zrzuty całego
ekranu trafiają do build/screens/.
"""

from __future__ import annotations

import os
import re
import sys

from qgis.core import QgsSettings
from qgis.PyQt.QtCore import QSize
from qgis.PyQt.QtWidgets import QApplication
from qgis_smoke import SCREENS, VERSION, check, make_layer  # uruchamia też QGIS (start_app)

from data_handover.core.i18n import LANGUAGES
from data_handover.dialog import COL_DESC, HandoverDialog
from data_handover.plugin import SETTINGS, UI_LANGUAGE

# Mały ekran: laptop 1366×768 ze skalowaniem 125 % to 1093×614 px logicznych;
# zostawiamy miejsce na pasek zadań i pasek tytułu okna.
SMALL = QSize(1000, 520)


def fresh() -> None:
    for name in ("window", "look_splitter", "preview_zoom"):
        QgsSettings().remove(SETTINGS + name)


def opened(layer: object) -> HandoverDialog:
    box = HandoverDialog(None, None, layer=layer)
    box.show()
    QApplication.processEvents()
    return box


def sizes(text: str) -> list[float]:
    return [float(v) for v in re.findall(r"font-size:([\d.]+)pt", text)]


def main() -> None:
    screen = QApplication.primaryScreen()
    area = screen.availableGeometry()
    ratio = screen.devicePixelRatio()
    tag = f"{round(area.width() * ratio)}x{round(area.height() * ratio)}@{ratio:g}"
    print(f"QGIS {VERSION}, ekran {tag} = {area.width()}×{area.height()} px logicznych")
    layer = make_layer()
    fresh()

    for lang in LANGUAGES:
        QgsSettings().setValue(UI_LANGUAGE, lang)
        box = opened(layer)
        minimum = box.minimumSizeHint().expandedTo(box.minimumSize())
        check(minimum.width() <= SMALL.width() and minimum.height() <= SMALL.height(),
              f"{lang}: minimum {minimum.width()}×{minimum.height()} mieści się na małym ekranie")
        check(area.contains(box.frameGeometry()), f"{lang}: okno {box.width()}×{box.height()} na ekranie {tag}")
        if lang == "en":
            for index, name in enumerate(("dane", "wyglad", "arkusze", "zmiany")):
                box.tabs.setCurrentIndex(index)
                QApplication.processEvents()
                screen.grabWindow(0).save(str(SCREENS / f"ekran_{tag}_{name}_qgis{VERSION}.png"))
        box.close()
    QgsSettings().remove(UI_LANGUAGE)
    fresh()  # domyślny rozdzielacz zależy od języka — dalej porównujemy z angielskim

    # Tabela pól: w szerokim oknie kolumny wypełniają tabelę (opis bierze resztę),
    # w najmniejszym tabela przewija się w poziomie, a opis nie jest ściśnięty.
    box = opened(layer)
    default_size = box.size()
    box.resize(min(area.width(), 1800), box.height())
    QApplication.processEvents()
    table = box.table
    check(table.horizontalHeader().length() >= table.viewport().width() - 1
          and table.horizontalScrollBar().maximum() == 0,
          f"szerokie okno: kolumny wypełniają tabelę pól (opis {table.columnWidth(COL_DESC)} px)")
    box.resize(default_size)
    box.tabs.setCurrentIndex(1)
    QApplication.processEvents()
    default_split = box.look_splitter.sizes()
    box.resize(box.minimumSizeHint())
    box.tabs.setCurrentIndex(0)
    QApplication.processEvents()
    check(table.horizontalScrollBar().maximum() > 0 and table.columnWidth(COL_DESC) >= 300,
          "najmniejsze okno: tabela pól ma poziomy suwak, opis nie jest ściśnięty")

    # Podgląd: + i − zmieniają wielkość pisma, w granicach; powiększony przewija się w poziomie.
    box.tabs.setCurrentIndex(1)
    QApplication.processEvents()
    base = sizes(box.preview.text())
    while box.zoom_in_button.isEnabled():
        box.zoom_in_button.click()
    QApplication.processEvents()
    big = sizes(box.preview.text())
    check(box.zoom == 200 and len(big) == len(base) > 0
          and all(abs(b - 2 * a) < 0.06 for a, b in zip(base, big)),
          f"powiększenie 200 %: pismo w podglądzie dwa razy większe ({base[:2]} → {big[:2]})")
    check(box.preview_scroll.horizontalScrollBar().maximum() > 0, "powiększony podgląd ma poziomy suwak")
    while box.zoom_out_button.isEnabled():
        box.zoom_out_button.click()
    check(box.zoom == 50 and box.zoom_in_button.isEnabled(), "pomniejszenie kończy się na 50 %")
    box.close()

    # Zapamiętane: rozmiar okna, rozdzielacz i powiększenie (także po zmianie języka okna).
    box = opened(layer)
    box.resize(900, 540)
    box.tabs.setCurrentIndex(1)
    QApplication.processEvents()
    box.look_splitter.setSizes([260, 600])
    box.zoom_in_button.click()
    box.zoom_in_button.click()
    zoom, split = box.zoom, box.look_splitter.sizes()
    box.close()
    again = opened(layer)
    again.tabs.setCurrentIndex(1)
    QApplication.processEvents()
    check(again.size() == QSize(900, 540) and again.look_splitter.sizes() == split and again.zoom == zoom != 100,
          f"okno pamięta rozmiar, rozdzielacz {split} i powiększenie {zoom} %")
    again.ui_lang_combo.setCurrentIndex(again.ui_lang_combo.findData("de"))
    german = opened(layer)
    check(german.size() == QSize(900, 540) and german.zoom == zoom, "po zmianie języka okno ma ten sam rozmiar")
    german.close()
    QgsSettings().remove(UI_LANGUAGE)

    # Rozmiar z większego monitora albo położenie poza ekranem: okno wraca na ekran.
    box = opened(layer)
    box.setGeometry(area.right() + 400, area.bottom() + 300, area.width() + 900, area.height() + 700)
    QApplication.processEvents()
    box.close()
    back = opened(layer)
    check(area.contains(back.frameGeometry()),
          f"zapisany rozmiar większy niż ekran i poza nim: okno {back.frameGeometry().getRect()} na ekranie")

    # „Przywróć domyślne”: rozmiar, rozdzielacz i powiększenie jak przy pierwszym otwarciu.
    back.tabs.setCurrentIndex(1)
    back.zoom_in_button.click()
    back.look_splitter.setSizes([200, 700])
    back.reset_settings()
    QApplication.processEvents()
    check(back.size() == default_size and back.look_splitter.sizes() == default_split and back.zoom == 100
          and area.contains(back.frameGeometry()),
          f"„Przywróć domyślne”: okno {back.width()}×{back.height()}, rozdzielacz i powiększenie domyślne")
    back.close()
    fresh()
    print("OK")


if __name__ == "__main__":
    main()
    sys.stdout.flush()
    os._exit(0)  # jak w qgis_smoke.py: bez sprzątania QGIS przy wyjściu
