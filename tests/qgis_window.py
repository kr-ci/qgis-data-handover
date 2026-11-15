"""Okno na każdym monitorze (R11), w prawdziwym QGIS.

`scripts/qgis-test.sh` uruchamia ten test kilka razy, na ekranach Xvfb o różnej
wielkości i ze skalowaniem (QT_SCALE_FACTOR). Sprawdza, że okno mieści się na
ekranie, przewija się zamiast ściskać, zapamiętuje rozmiar i rozdzielacz,
a „Przywróć domyślne” je zeruje. Podgląd pokazuje czcionki w rozmiarze
z ustawień. Zrzuty całego ekranu trafiają do build/screens/.
"""

from __future__ import annotations

import os
import re
import sys

from qgis.core import QgsSettings
from qgis.PyQt.QtCore import QByteArray, QSize
from qgis.PyQt.QtWidgets import QApplication
from qgis_smoke import SCREENS, VERSION, check, make_layer  # uruchamia też QGIS (start_app)

from data_handover.core.i18n import LANGUAGES
from data_handover.dialog import COL_DESC, HandoverDialog
from data_handover.plugin import SETTINGS, UI_LANGUAGE

# Mały ekran: laptop 1366×768 ze skalowaniem 125 % to 1093×614 px logicznych;
# zostawiamy miejsce na pasek zadań i pasek tytułu okna.
SMALL = QSize(1000, 520)


def fresh() -> None:
    for name in ("window", "look_splitter"):
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

    # Uszkodzone albo obce ustawienia nie blokują okna (przegląd R11, zarzut 1).
    for saved in ("abc", QByteArray(b"junk")):
        for name in ("window", "look_splitter"):
            QgsSettings().setValue(SETTINGS + name, saved)
        box = opened(layer)
        check(area.contains(box.frameGeometry()), f"uszkodzone ustawienia ({saved!r}): okno otwiera się na ekranie")
        box.close()
    fresh()

    # Podgląd pokazuje czcionki w rozmiarze z ustawień — bez powiększenia (decyzja właściciela
    # 2026-11-15), także gdy w profilu został wpis powiększenia z wersji roboczej.
    QgsSettings().setValue(SETTINGS + "preview_zoom", 200)
    box = opened(layer)
    look, found = box.look(), sizes(box.preview.text())
    check(look.body.size in found
          and all(s in (look.header.size, look.body.size, min(look.title_size, 20)) for s in found),
          f"podgląd: pismo w rozmiarze z ustawień ({sorted(set(found))} pt)")
    box.close()
    QgsSettings().remove(SETTINGS + "preview_zoom")

    # Każdy język: okno na ekranie, także gdy ekran jest węższy niż stopka w jednym wierszu.
    for lang in LANGUAGES:
        QgsSettings().setValue(UI_LANGUAGE, lang)
        box = opened(layer)
        minimum = box.minimumSizeHint().expandedTo(box.minimumSize())
        check(minimum.width() <= SMALL.width() and minimum.height() <= SMALL.height(),
              f"{lang}: minimum {minimum.width()}×{minimum.height()} mieści się na małym ekranie")
        check(area.contains(box.frameGeometry()), f"{lang}: okno {box.width()}×{box.height()} na ekranie {tag}")
        shots = {"en": ("dane", "wyglad", "arkusze", "zmiany"), "es": ("dane",)}.get(lang, ())
        for index, name in enumerate(shots):
            box.tabs.setCurrentIndex(index)
            QApplication.processEvents()
            screen.grabWindow(0).save(str(SCREENS / f"ekran_{tag}_{name}_{lang}_qgis{VERSION}.png"))
        box.close()
    QgsSettings().remove(UI_LANGUAGE)
    fresh()  # domyślny rozdzielacz zależy od języka — dalej porównujemy z angielskim

    # Tabela pól: w szerokim oknie kolumny wypełniają tabelę (opis bierze resztę),
    # w najmniejszym tabela przewija się w poziomie, a opis nie jest ściśnięty.
    box = opened(layer)
    default_size = box.size()
    box.resize(1800, box.height())
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
    box.tabs.setCurrentIndex(1)
    QApplication.processEvents()
    check(box.preview_scroll.horizontalScrollBar().maximum() > 0,
          "najmniejsze okno: podgląd przewija się w poziomie, zamiast ściskać tabelę")
    box.close()

    # Zapamiętane: rozmiar okna i rozdzielacz (także po zmianie języka okna).
    box = opened(layer)
    box.resize(area.width() * 4 // 5, (area.height() - 40) * 4 // 5)
    box.tabs.setCurrentIndex(1)
    QApplication.processEvents()
    wanted = box.size()
    box.look_splitter.setSizes([260, 600])
    split = box.look_splitter.sizes()
    box.close()
    again = opened(layer)
    again.tabs.setCurrentIndex(1)
    QApplication.processEvents()
    check(again.size() == wanted and again.look_splitter.sizes() == split,
          f"okno pamięta rozmiar i rozdzielacz {split}")
    again.ui_lang_combo.setCurrentIndex(again.ui_lang_combo.findData("de"))
    german = opened(layer)
    check(german.size() == wanted.expandedTo(german.minimumSizeHint()),
          "po zmianie języka okno ma ten sam rozmiar (chyba że minimum w nowym języku jest większe)")
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

    # „Przywróć domyślne”: rozmiar i rozdzielacz jak przy pierwszym otwarciu.
    back.tabs.setCurrentIndex(1)
    back.look_splitter.setSizes([200, 700])
    back.reset_settings()
    QApplication.processEvents()
    # Gdy okno jest węższe niż ustawienia i podgląd razem, Qt rozdziela brak miejsca zależnie
    # od kolejności zmian rozmiaru — po resecie rozdzielacz może stać kilka pikseli dalej.
    split_ok = all(abs(a - b) <= 10 for a, b in zip(back.look_splitter.sizes(), default_split))
    check(back.size() == default_size and split_ok and area.contains(back.frameGeometry()),
          f"„Przywróć domyślne”: okno {back.width()}×{back.height()}, rozdzielacz {back.look_splitter.sizes()}"
          f" (domyślny {default_split})")
    back.close()
    fresh()
    print("OK")


if __name__ == "__main__":
    main()
    sys.stdout.flush()
    os._exit(0)  # jak w qgis_smoke.py: bez sprzątania QGIS przy wyjściu
