"""Okno wtyczki. Kod zgodny z PyQt5 i PyQt6: tylko `qgis.PyQt` i pełne nazwy enumów."""

from __future__ import annotations

import contextlib
import datetime as dt
import html
import json
import os
import re
from typing import Any, Optional

from qgis.core import Qgis, QgsMessageLog, QgsSettings, QgsVectorLayer
from qgis.gui import QgsColorButton, QgsFieldComboBox, QgsFileWidget, QgsMapLayerComboBox
from qgis.PyQt.QtCore import Qt, QUrl, pyqtSignal
from qgis.PyQt.QtGui import QColor, QDesktopServices, QIcon, QPainter, QPixmap
from qgis.PyQt.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFontComboBox,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from . import export_task, layer_reader
from .core import privacy, xlsx
from .core.i18n import LANGUAGE_NAMES, LANGUAGES, count_label, tr
from .core.model import (
    ABOUT_ITEMS,
    ALIGNMENTS,
    BORDER_SCOPES,
    BORDER_STYLES,
    DICTIONARY_ITEMS,
    THEMES,
    CellLook,
    Look,
    Sheets,
    Theme,
    from_dict,
    look_with_theme,
    theme_from_color,
    to_dict,
)
from .layer_reader import FieldChoice
from .plugin import UI_LANGUAGE, ui_language

SETTINGS = "data_handover/"
PREVIOUS = "data_handover/previous"
DEFAULT_LANGUAGE = "en"  # język pliku, dopóki użytkownik nie wybierze innego
# Stały wygląd okna — niezależny od stylu pliku, który użytkownik ustawia.
ACCENT, ACCENT_TEXT = "1F3A5F", "FFFFFF"
COL_FIELD, COL_KIND, COL_HEADER, COL_UNIT, COL_DESC, COL_PRIVACY = range(6)
NAME_ROLE = Qt.ItemDataRole.UserRole
KIND_ROLE = Qt.ItemDataRole.UserRole + 1
PERSONAL_COLOR = QColor("#9C2B23")
# Grubość i rodzaj linii w podglądzie HTML odpowiadające liniom Excela.
CSS_BORDER = {"hair": "1px dotted", "thin": "1px solid", "medium": "2px solid", "thick": "3px solid",
              "dashed": "1px dashed", "dotted": "1px dotted", "double": "3px double"}


def safe_filename(text: str) -> str:
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", text).strip().replace(" ", "_")
    return text[:80] or "dane"


def swatch(theme: Theme, size: int = 26) -> QIcon:
    pixmap = QPixmap(size, size)
    pixmap.fill(QColor("#" + theme.soft))
    painter = QPainter(pixmap)
    painter.fillRect(0, 0, size, size * 6 // 10, QColor("#" + theme.header_fill))
    painter.end()
    return QIcon(pixmap)


def rgb(button: QgsColorButton) -> str:
    """Kolor z przycisku w zapisie RRGGBB; pusty, gdy wybrano „bez koloru”."""
    if button.isNull():
        return ""
    return button.color().name()[1:].upper()


def set_rgb(button: QgsColorButton, value: str) -> None:
    if value:
        button.setColor(QColor("#" + value))
    else:
        button.setToNull()


def color_button(title: str, allow_none: str = "") -> QgsColorButton:
    button = QgsColorButton()
    button.setAllowOpacity(False)
    button.setColorDialogTitle(title)
    button.setMinimumWidth(70)
    if allow_none:
        button.setShowNull(True, allow_none)
    return button


class CellLookEditor(QGroupBox):
    """Ustawienia czcionki, kolorów i wyrównania jednej grupy komórek."""

    changed = pyqtSignal()

    def __init__(self, title: str, lang: str) -> None:
        super().__init__(title)
        t = lambda text_id: tr(lang, text_id)  # noqa: E731
        grid = QGridLayout(self)
        self.font_combo = QFontComboBox()
        self.font_combo.setToolTip(t("ui.look.font.hint"))
        self.size_spin = QSpinBox()
        self.size_spin.setRange(6, 48)
        self.size_spin.setSuffix(" pt")
        self.style_buttons: dict[str, QToolButton] = {}
        styles = QHBoxLayout()
        for key, letter in (("bold", "B"), ("italic", "I"), ("underline", "U")):
            button = QToolButton()
            button.setText(letter)
            button.setCheckable(True)
            button.setToolTip(t(f"ui.look.{key}"))
            font = button.font()
            font.setBold(key == "bold")
            font.setItalic(key == "italic")
            font.setUnderline(key == "underline")
            button.setFont(font)
            button.toggled.connect(self.changed)
            self.style_buttons[key] = button
            styles.addWidget(button)
        styles.addStretch()
        self.color_button = color_button(t("ui.look.color"))
        self.fill_button = color_button(t("ui.look.fill"), t("ui.look.no_fill"))
        self.align_combo = QComboBox()
        for key in ALIGNMENTS:
            self.align_combo.addItem(t(f"ui.look.align.{key}"), key)
        self.wrap_check = QCheckBox(t("ui.look.wrap"))

        grid.addWidget(QLabel(t("ui.look.font")), 0, 0)
        grid.addWidget(self.font_combo, 0, 1)
        grid.addWidget(self.size_spin, 0, 2)
        grid.addLayout(styles, 0, 3)
        grid.addWidget(QLabel(t("ui.look.color")), 1, 0)
        grid.addWidget(self.color_button, 1, 1)
        grid.addWidget(QLabel(t("ui.look.fill")), 1, 2)
        grid.addWidget(self.fill_button, 1, 3)
        grid.addWidget(QLabel(t("ui.look.align")), 2, 0)
        grid.addWidget(self.align_combo, 2, 1)
        grid.addWidget(self.wrap_check, 2, 2, 1, 2)
        grid.setColumnStretch(1, 1)
        self.extra_row = 3  # miejsce na dodatkowe pola (pasy wierszy w „Zawartości”)

        self.font_combo.editTextChanged.connect(self.changed)
        self.size_spin.valueChanged.connect(self.changed)
        self.color_button.colorChanged.connect(self.changed)
        self.fill_button.colorChanged.connect(self.changed)
        self.align_combo.currentIndexChanged.connect(self.changed)
        self.wrap_check.toggled.connect(self.changed)

    def value(self) -> CellLook:
        return CellLook(
            # Nazwa jako tekst: czcionki może nie być w tym systemie (np. Calibri na Linuksie),
            # a odbiorca w Excelu ją ma — nie podmieniamy jej na zastępczą.
            font=self.font_combo.currentText().strip() or "Calibri",
            size=self.size_spin.value(),
            color=rgb(self.color_button) or "000000",
            bold=self.style_buttons["bold"].isChecked(),
            italic=self.style_buttons["italic"].isChecked(),
            underline=self.style_buttons["underline"].isChecked(),
            fill=rgb(self.fill_button),
            align=str(self.align_combo.currentData()),
            wrap=self.wrap_check.isChecked(),
        )

    def set_value(self, look: CellLook) -> None:
        self.font_combo.setEditText(look.font)
        self.size_spin.setValue(look.size)
        set_rgb(self.color_button, look.color)
        for key, button in self.style_buttons.items():
            button.setChecked(bool(getattr(look, key)))
        set_rgb(self.fill_button, look.fill)
        self.align_combo.setCurrentIndex(max(0, self.align_combo.findData(look.align)))
        self.wrap_check.setChecked(look.wrap)


class HandoverDialog(QDialog):
    def __init__(self, iface: Any = None, parent: Optional[QWidget] = None,
                 layer: Optional[QgsVectorLayer] = None) -> None:
        super().__init__(parent)
        self.iface = iface
        self.lang = ui_language()
        self.settings = QgsSettings()
        self.personal: dict[str, str] = {}
        self.samples: dict[str, list[Any]] = {}
        self.layer: Optional[QgsVectorLayer] = None
        self.custom_color = str(self.settings.value(SETTINGS + "custom_color", "#2F6F8F"))
        self._filling = False
        self.relaunch = False  # True: wtyczka otworzy okno ponownie w nowym języku
        self.task: Optional[export_task.ExportTask] = None
        self.progress_item: Any = None
        self.layer_gone = False
        self.last_result: Optional[xlsx.ExportResult] = None

        self.setWindowTitle(self.t("ui.title"))
        self.setMinimumSize(980, 700)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 12)
        root.addWidget(self._banner())

        self.tabs = QTabWidget()
        self.tabs.addTab(self._data_tab(), self.t("ui.tab.data"))
        self.tabs.addTab(self._look_tab(), self.t("ui.tab.look"))
        self.tabs.addTab(self._sheets_tab(), self.t("ui.tab.sheets"))
        self.tabs.addTab(self._changes_tab(), self.t("ui.tab.changes"))
        body = QVBoxLayout()
        body.setContentsMargins(12, 8, 12, 0)
        body.addWidget(self.tabs)
        root.addLayout(body)

        buttons = QDialogButtonBox()
        reset = QPushButton(self.t("ui.reset"))
        reset.setToolTip(self.t("ui.reset.tip"))
        reset.clicked.connect(self.reset_settings)
        buttons.addButton(reset, QDialogButtonBox.ButtonRole.ResetRole)
        self.export_button = QPushButton(self.t("ui.export"))
        self.export_button.setDefault(True)
        self.export_button.setObjectName("exportButton")
        buttons.addButton(self.export_button, QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton(QPushButton(self.t("ui.close")), QDialogButtonBox.ButtonRole.RejectRole)
        self.export_button.clicked.connect(self.export)
        buttons.rejected.connect(self.reject)
        self.ui_lang_combo = QComboBox()
        for code in LANGUAGES:
            self.ui_lang_combo.addItem(LANGUAGE_NAMES[code], code)
        self.ui_lang_combo.setCurrentIndex(max(0, self.ui_lang_combo.findData(self.lang)))
        self.ui_lang_combo.currentIndexChanged.connect(self._switch_language)
        footer = QHBoxLayout()
        footer.setContentsMargins(12, 0, 12, 0)
        footer.addWidget(QLabel(self.t("ui.language")))
        footer.addWidget(self.ui_lang_combo)
        footer.addSpacing(12)
        footer.addWidget(buttons)
        root.addLayout(footer)

        self._apply_window_style()
        self._set_look(from_dict(Look, self._saved_json("look")))
        self._set_sheets(from_dict(Sheets, self._saved_json("sheets")))
        self.layer_combo.layerChanged.connect(self._load_layer)
        if layer is not None:
            self.layer_combo.setLayer(layer)
        self._load_layer(self.layer_combo.currentLayer())

    def t(self, text_id: str, **kwargs: object) -> str:
        return tr(self.lang, text_id, **kwargs)

    def _saved_json(self, name: str) -> Any:
        try:
            return json.loads(str(self.settings.value(SETTINGS + name, "") or "null"))
        except ValueError:
            return None

    # ------------------------------------------------------------------ UI

    def _banner(self) -> QFrame:
        self.banner = QFrame()
        self.banner.setObjectName("banner")
        layout = QVBoxLayout(self.banner)
        layout.setContentsMargins(18, 14, 18, 14)
        title = QLabel("Data Handover")
        title.setObjectName("bannerTitle")
        tagline = QLabel(self.t("ui.tagline"))
        tagline.setObjectName("bannerTagline")
        layout.addWidget(title)
        layout.addWidget(tagline)
        return self.banner

    def _data_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        form = QFormLayout()
        self.layer_combo = QgsMapLayerComboBox()
        self.layer_combo.setFilters(Qgis.LayerFilter.VectorLayer)
        form.addRow(self.t("ui.layer"), self.layer_combo)
        scope_row = QHBoxLayout()
        self.scope_group = QButtonGroup(self)
        self.scope_buttons: dict[str, QRadioButton] = {}
        for scope in ("all", "selected", "extent"):
            button = QRadioButton()
            self.scope_buttons[scope] = button
            self.scope_group.addButton(button)
            scope_row.addWidget(button)
            button.toggled.connect(self._update_preview)
        scope_row.addStretch()
        self.scope_buttons["all"].setChecked(True)
        form.addRow(self.t("ui.scope"), scope_row)
        layout.addLayout(form)

        hint = QLabel(self.t("ui.fields.hint"))
        hint.setWordWrap(True)
        hint.setObjectName("hint")
        layout.addWidget(hint)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels([
            self.t("ui.col.field"), self.t("ui.col.kind"), self.t("ui.col.header"),
            self.t("ui.col.unit"), self.t("ui.col.description"), self.t("ui.col.privacy"),
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        for col, mode in ((COL_FIELD, QHeaderView.ResizeMode.ResizeToContents),
                          (COL_KIND, QHeaderView.ResizeMode.ResizeToContents),
                          (COL_HEADER, QHeaderView.ResizeMode.Interactive),
                          (COL_UNIT, QHeaderView.ResizeMode.ResizeToContents),
                          (COL_DESC, QHeaderView.ResizeMode.Stretch),
                          (COL_PRIVACY, QHeaderView.ResizeMode.ResizeToContents)):
            header.setSectionResizeMode(col, mode)
        self.table.setColumnWidth(COL_HEADER, 180)
        self.table.itemChanged.connect(self._update_preview)
        layout.addWidget(self.table)

        row = QHBoxLayout()
        for text_id, slot in (("ui.select.all", lambda: self._check_all(True)),
                              ("ui.select.none", lambda: self._check_all(False))):
            button = QPushButton(self.t(text_id))
            button.clicked.connect(slot)
            row.addWidget(button)
        row.addStretch()
        for text_id, step in (("ui.move.up", -1), ("ui.move.down", 1)):
            button = QPushButton(self.t(text_id))
            button.clicked.connect(lambda _checked=False, s=step: self._move(s))
            row.addWidget(button)
        layout.addLayout(row)
        return tab

    def _look_tab(self) -> QWidget:
        tab = QWidget()
        outer = QHBoxLayout(tab)
        content = QWidget()
        column = QVBoxLayout(content)

        presets = QHBoxLayout()
        presets.addWidget(QLabel(self.t("ui.file.theme")))
        for name, theme in THEMES.items():
            button = QToolButton()
            button.setIcon(swatch(theme))
            button.setIconSize(button.iconSize() * 1.6)
            button.setToolTip(self.t(f"ui.theme.{name}"))
            button.clicked.connect(lambda _checked=False, th=theme: self._apply_preset(th))
            presets.addWidget(button)
        pick = QPushButton(self.t("ui.file.custom"))
        pick.clicked.connect(self._pick_color)
        presets.addWidget(pick)
        presets.addStretch()
        column.addLayout(presets)

        layout_box = QGroupBox(self.t("ui.look.layout"))
        grid = QGridLayout(layout_box)
        self.as_table_check = QCheckBox(self.t("ui.look.as_table"))
        self.as_table_check.setToolTip(self.t("ui.look.as_table.tip"))
        font = self.as_table_check.font()
        font.setBold(True)
        self.as_table_check.setFont(font)
        self.autofilter_check = QCheckBox(self.t("ui.look.autofilter"))
        self.title_check = QCheckBox(self.t("ui.look.show_title"))
        self.freeze_check = QCheckBox(self.t("ui.look.freeze"))
        self.comments_check = QCheckBox(self.t("ui.look.comments"))
        self.gridlines_check = QCheckBox(self.t("ui.look.gridlines"))
        grid.addWidget(self.as_table_check, 0, 0)
        grid.addWidget(self.autofilter_check, 0, 1)
        grid.addWidget(self.title_check, 1, 0)
        grid.addWidget(self.freeze_check, 1, 1)
        grid.addWidget(self.comments_check, 2, 0)
        grid.addWidget(self.gridlines_check, 2, 1)

        form = QFormLayout()
        title_row = QHBoxLayout()
        self.title_color = color_button(self.t("ui.look.title"))
        self.title_size = QSpinBox()
        self.title_size.setRange(10, 36)
        self.title_size.setSuffix(" pt")
        title_row.addWidget(self.title_color)
        title_row.addWidget(self.title_size)
        title_row.addStretch()
        form.addRow(self.t("ui.look.title"), title_row)
        self.max_width = QSpinBox()
        self.max_width.setRange(10, 150)
        self.max_width.setSuffix(self.t("ui.look.max_width.suffix"))
        form.addRow(self.t("ui.look.max_width"), self.max_width)
        other_row = QHBoxLayout()
        self.orientation = QComboBox()
        self.orientation.addItem(self.t("ui.look.landscape"), True)
        self.orientation.addItem(self.t("ui.look.portrait"), False)
        self.tab_color = color_button(self.t("ui.look.tab_color"))
        other_row.addWidget(self.orientation)
        other_row.addSpacing(16)
        other_row.addWidget(QLabel(self.t("ui.look.tab_color")))
        other_row.addWidget(self.tab_color)
        other_row.addStretch()
        form.addRow(self.t("ui.look.print"), other_row)
        grid.addLayout(form, 3, 0, 1, 2)
        column.addWidget(layout_box)

        self.header_editor = CellLookEditor(self.t("ui.look.header"), self.lang)
        self.body_editor = CellLookEditor(self.t("ui.look.body"), self.lang)
        stripes = QHBoxLayout()
        self.stripes_check = QCheckBox(self.t("ui.look.stripes"))
        self.stripe_color = color_button(self.t("ui.look.stripes"))
        stripes.addWidget(self.stripes_check)
        stripes.addWidget(self.stripe_color)
        stripes.addStretch()
        body_grid = self.body_editor.layout()
        assert isinstance(body_grid, QGridLayout)
        body_grid.addLayout(stripes, self.body_editor.extra_row, 0, 1, 4)
        column.addWidget(self.header_editor)
        column.addWidget(self.body_editor)

        border_box = QGroupBox(self.t("ui.look.borders"))
        border_row = QHBoxLayout(border_box)
        self.border_style = QComboBox()
        for key in BORDER_STYLES:
            self.border_style.addItem(self.t(f"ui.look.border.{key}"), key)
        self.border_color = color_button(self.t("ui.look.border.color"))
        self.border_scope = QComboBox()
        for key in BORDER_SCOPES:
            self.border_scope.addItem(self.t(f"ui.look.border.{key}"), key)
        for label, widget in (("ui.look.border.style", self.border_style),
                              ("ui.look.border.color", self.border_color),
                              ("ui.look.border.scope", self.border_scope)):
            border_row.addWidget(QLabel(self.t(label)))
            border_row.addWidget(widget)
            border_row.addSpacing(10)
        border_row.addStretch()
        column.addWidget(border_box)
        column.addStretch()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setMinimumWidth(content.sizeHint().width() + 20)
        scroll.setWidget(content)
        outer.addWidget(scroll, 3)

        preview_box = QVBoxLayout()
        caption = QLabel(self.t("ui.file.preview"))
        caption.setObjectName("hint")
        self.preview = QLabel()
        self.preview.setTextFormat(Qt.TextFormat.RichText)
        self.preview.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.preview.setObjectName("preview")
        self.preview.setMinimumWidth(330)
        self.preview.setWordWrap(True)
        preview_box.addWidget(caption)
        preview_box.addWidget(self.preview, 1)
        outer.addLayout(preview_box, 2)

        for check in (self.as_table_check, self.autofilter_check, self.title_check, self.freeze_check,
                      self.comments_check, self.gridlines_check, self.stripes_check):
            check.toggled.connect(self._look_changed)
        for spin in (self.title_size, self.max_width):
            spin.valueChanged.connect(self._look_changed)
        for combo in (self.orientation, self.border_style, self.border_scope):
            combo.currentIndexChanged.connect(self._look_changed)
        for button in (self.title_color, self.tab_color, self.stripe_color, self.border_color):
            button.colorChanged.connect(self._look_changed)
        self.header_editor.changed.connect(self._look_changed)
        self.body_editor.changed.connect(self._look_changed)
        return tab

    def _sheets_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        form = QFormLayout()
        self.title_edit = QLineEdit()
        self.title_edit.textChanged.connect(self._update_preview)
        self.lang_combo = QComboBox()
        for code in LANGUAGES:
            self.lang_combo.addItem(LANGUAGE_NAMES[code], code)
        saved_lang = str(self.settings.value(SETTINGS + "file_language", DEFAULT_LANGUAGE))
        self.lang_combo.setCurrentIndex(max(0, self.lang_combo.findData(saved_lang)))
        self.lang_combo.currentIndexChanged.connect(self._update_preview)
        form.addRow(self.t("ui.file.title"), self.title_edit)
        form.addRow(self.t("ui.file.language"), self.lang_combo)
        layout.addLayout(form)
        hint = QLabel(self.t("ui.sheets.data"))
        hint.setObjectName("hint")
        layout.addWidget(hint)

        self.dictionary_box = QGroupBox(self.t("ui.sheets.dictionary"))
        self.dictionary_box.setCheckable(True)
        dict_layout = QVBoxLayout(self.dictionary_box)
        dict_layout.addWidget(QLabel(self.t("ui.sheets.dictionary.hint")))
        row = QHBoxLayout()
        self.dictionary_checks: dict[str, QCheckBox] = {}
        for key in DICTIONARY_ITEMS:
            check = QCheckBox(self.t(f"dict.col.{key}"))
            self.dictionary_checks[key] = check
            row.addWidget(check)
        row.addStretch()
        dict_layout.addLayout(row)
        layout.addWidget(self.dictionary_box)

        self.about_box = QGroupBox(self.t("ui.sheets.about"))
        self.about_box.setCheckable(True)
        about_layout = QVBoxLayout(self.about_box)
        about_layout.addWidget(QLabel(self.t("ui.sheets.about.hint")))
        grid = QGridLayout()
        self.about_checks: dict[str, QCheckBox] = {}
        for i, key in enumerate(ABOUT_ITEMS):
            label = {"guide": "about.guide", "notes": "about.notes"}.get(key, f"fact.{key}")
            check = QCheckBox(self.t(label))
            self.about_checks[key] = check
            grid.addWidget(check, i // 4, i % 4)
        about_layout.addLayout(grid)
        about_form = QFormLayout()
        self.author_edit = QLineEdit(str(self.settings.value(SETTINGS + "author", "")))
        self.note_edit = QPlainTextEdit()
        self.note_edit.setPlaceholderText(self.t("ui.file.note.placeholder"))
        self.note_edit.setMaximumHeight(70)
        about_form.addRow(self.t("ui.file.author"), self.author_edit)
        about_form.addRow(self.t("ui.file.note"), self.note_edit)
        about_layout.addLayout(about_form)
        layout.addWidget(self.about_box)
        layout.addStretch()

        self.dictionary_box.toggled.connect(self._update_preview)
        self.about_box.toggled.connect(self._update_preview)
        self.about_checks["author"].toggled.connect(self.author_edit.setEnabled)
        self.about_checks["note"].toggled.connect(self.note_edit.setEnabled)
        return tab

    def _changes_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.changes_check = QCheckBox(self.t("ui.changes.enable"))
        layout.addWidget(self.changes_check)
        form = QFormLayout()
        self.previous_widget = QgsFileWidget()
        self.previous_widget.setFilter(self.t("ui.save.filter"))
        self.key_combo = QgsFieldComboBox()
        self.key_combo.setAllowEmptyFieldName(True)
        form.addRow(self.t("ui.changes.file"), self.previous_widget)
        form.addRow(self.t("ui.changes.key"), self.key_combo)
        layout.addLayout(form)
        hint = QLabel(self.t("ui.changes.hint"))
        hint.setWordWrap(True)
        hint.setObjectName("hint")
        layout.addWidget(hint)
        layout.addStretch()
        self.changes_check.toggled.connect(self._toggle_changes)
        self._toggle_changes(False)
        return tab

    # ---------------------------------------------------------- look state

    def look(self) -> Look:
        return Look(
            as_table=self.as_table_check.isChecked(),
            autofilter=self.autofilter_check.isChecked(),
            show_title=self.title_check.isChecked(),
            title_color=rgb(self.title_color) or "000000",
            title_size=self.title_size.value(),
            header=self.header_editor.value(),
            body=self.body_editor.value(),
            stripes=self.stripes_check.isChecked(),
            stripe_color=rgb(self.stripe_color) or "F2F2F2",
            border_style=str(self.border_style.currentData()),
            border_color=rgb(self.border_color) or "000000",
            border_scope=str(self.border_scope.currentData()),
            freeze_header=self.freeze_check.isChecked(),
            header_comments=self.comments_check.isChecked(),
            gridlines=self.gridlines_check.isChecked(),
            max_width=self.max_width.value(),
            landscape=bool(self.orientation.currentData()),
            tab_color=rgb(self.tab_color) or "000000",
        )

    def _set_look(self, look: Look) -> None:
        self._filling = True
        self.as_table_check.setChecked(look.as_table)
        self.autofilter_check.setChecked(look.autofilter)
        self.title_check.setChecked(look.show_title)
        set_rgb(self.title_color, look.title_color)
        self.title_size.setValue(look.title_size)
        self.header_editor.set_value(look.header)
        self.body_editor.set_value(look.body)
        self.stripes_check.setChecked(look.stripes)
        set_rgb(self.stripe_color, look.stripe_color)
        self.border_style.setCurrentIndex(max(0, self.border_style.findData(look.border_style)))
        set_rgb(self.border_color, look.border_color)
        self.border_scope.setCurrentIndex(max(0, self.border_scope.findData(look.border_scope)))
        self.freeze_check.setChecked(look.freeze_header)
        self.comments_check.setChecked(look.header_comments)
        self.gridlines_check.setChecked(look.gridlines)
        self.max_width.setValue(look.max_width)
        self.orientation.setCurrentIndex(0 if look.landscape else 1)
        set_rgb(self.tab_color, look.tab_color)
        self._filling = False
        self._look_changed()

    def _look_changed(self, *_args: Any) -> None:
        if self._filling:
            return
        # Tabela Excela zawsze ma filtry w nagłówkach — osobny przełącznik dotyczy zwykłego zakresu.
        self.autofilter_check.setEnabled(not self.as_table_check.isChecked())
        self.stripe_color.setEnabled(self.stripes_check.isChecked())
        self.border_color.setEnabled(self.border_style.currentData() != "none")
        self.border_scope.setEnabled(self.border_style.currentData() != "none")
        self.title_color.setEnabled(self.title_check.isChecked())
        self.title_size.setEnabled(self.title_check.isChecked())
        self._update_preview()

    def _apply_preset(self, theme: Theme) -> None:
        self._set_look(look_with_theme(self.look(), theme))

    def _pick_color(self) -> None:
        color = QColorDialog.getColor(QColor(self.custom_color), self, self.t("ui.file.custom"))
        if color.isValid():
            self.custom_color = color.name()
            self._apply_preset(theme_from_color(self.custom_color))

    def _apply_window_style(self) -> None:
        """Wygląd samego okna wtyczki — stały, nie zależy od ustawień pliku."""
        self.setStyleSheet(
            f"#banner {{ background: #{ACCENT}; }}"
            f"#bannerTitle {{ color: #{ACCENT_TEXT}; font-size: 17pt; font-weight: 600; }}"
            f"#bannerTagline {{ color: #{ACCENT_TEXT}; }}"
            "#hint { color: palette(mid); }"
            "#preview { background: white; border: 1px solid #D0D0D0; padding: 10px; }"
            f"#exportButton {{ background: #{ACCENT}; color: #{ACCENT_TEXT};"
            " padding: 6px 18px; font-weight: 600; border-radius: 3px; }"
        )

    def reset_settings(self) -> None:
        """Przywraca domyślny wygląd, arkusze i język pliku (opisy kolumn w projekcie zostają)."""
        self._set_look(Look())
        self._set_sheets(Sheets())
        self.lang_combo.setCurrentIndex(max(0, self.lang_combo.findData(DEFAULT_LANGUAGE)))

    def _switch_language(self, *_args: Any) -> None:
        """Zapisuje język okna i bieżące ustawienia; wtyczka otworzy okno od nowa."""
        self.settings.setValue(UI_LANGUAGE, str(self.ui_lang_combo.currentData()))
        self._save_state()
        self.relaunch = True
        self.reject()

    def sheets(self) -> Sheets:
        return Sheets(
            dictionary=self.dictionary_box.isChecked(),
            dictionary_items=tuple(k for k, c in self.dictionary_checks.items() if c.isChecked()),
            about=self.about_box.isChecked(),
            about_items=tuple(k for k, c in self.about_checks.items() if c.isChecked()),
        )

    def _set_sheets(self, sheets: Sheets) -> None:
        self.dictionary_box.setChecked(sheets.dictionary)
        for key, check in self.dictionary_checks.items():
            check.setChecked(key in sheets.dictionary_items)
        self.about_box.setChecked(sheets.about)
        for key, check in self.about_checks.items():
            check.setChecked(key in sheets.about_items)
        self.author_edit.setEnabled(self.about_checks["author"].isChecked())
        self.note_edit.setEnabled(self.about_checks["note"].isChecked())

    # --------------------------------------------------------------- layer

    def _load_layer(self, layer: Any) -> None:
        if self.layer is not None:
            try:
                self.layer.selectionChanged.disconnect(self._update_scope_labels)
            except (TypeError, RuntimeError):
                pass
        self.layer = layer if isinstance(layer, QgsVectorLayer) else None
        self.table.setRowCount(0)
        self.key_combo.setLayer(self.layer)
        if self.layer is None:
            self._update_scope_labels()
            return
        self.samples = layer_reader.sample_values(self.layer, 200, represented=True)
        fields = [(f.name(), self.layer.attributeDisplayName(i)) for i, f in enumerate(self.layer.fields())]
        self.personal = {f.field: f.reason for f in privacy.detect(fields, self.samples)}
        saved = layer_reader.load_choices(self.layer)
        stored = saved.get("fields", {})
        names = self.layer.fields().names()
        saved_order = [n for n in saved.get("order", []) if n in names]
        order = saved_order + [n for n in names if n not in saved_order]
        self.title_edit.setText(self.layer.name())
        self.key_combo.setField(saved.get("key", ""))
        previous = self.layer.customProperty(PREVIOUS, "")
        self.previous_widget.setFilePath(previous)
        self.changes_check.setChecked(bool(previous and saved.get("key")))

        self._filling = True
        for name in order:
            index = self.layer.fields().indexOf(name)
            info = stored.get(name, {})
            self._add_row(name, index, info)
        self._filling = False
        self.layer.selectionChanged.connect(self._update_scope_labels)
        self._update_scope_labels()
        self._update_preview()

    def _add_row(self, name: str, index: int, info: dict[str, Any]) -> None:
        assert self.layer is not None
        row = self.table.rowCount()
        self.table.insertRow(row)
        field = self.layer.fields().at(index)
        kind = layer_reader.field_kind(field)
        widget = self.layer.editorWidgetSetup(index).type()
        personal = name in self.personal

        item = QTableWidgetItem(name)
        item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        default_export = not personal
        item.setCheckState(Qt.CheckState.Checked if info.get("export", default_export) else Qt.CheckState.Unchecked)
        item.setData(NAME_ROLE, name)
        self.table.setItem(row, COL_FIELD, item)

        kind_text = self.t(f"kind.{kind}")
        if widget in ("ValueMap", "ValueRelation", "RelationReference"):
            kind_text += " · " + widget
        kind_item = QTableWidgetItem(kind_text)
        kind_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        kind_item.setData(KIND_ROLE, kind)
        self.table.setItem(row, COL_KIND, kind_item)

        defaults = {COL_HEADER: self.layer.attributeDisplayName(index), COL_UNIT: "", COL_DESC: field.comment()}
        for col, key in ((COL_HEADER, "header"), (COL_UNIT, "unit"), (COL_DESC, "description")):
            self.table.setItem(row, col, QTableWidgetItem(str(info.get(key) or defaults[col])))

        if personal:
            tip = self.t("ui.privacy.tooltip", reason=self.personal[name])
            for col in (COL_FIELD, COL_KIND):
                self.table.item(row, col).setForeground(PERSONAL_COLOR)
                self.table.item(row, col).setToolTip(tip)
            combo = QComboBox()
            combo.addItem(self.t("ui.privacy.mask"), "mask")
            combo.addItem(self.t("ui.privacy.keep"), "keep")
            combo.setCurrentIndex(max(0, combo.findData(info.get("protect") or "mask")))
            combo.setToolTip(tip)
            combo.currentIndexChanged.connect(self._update_preview)
            self.table.setCellWidget(row, COL_PRIVACY, combo)

    def _row_state(self, row: int) -> tuple[FieldChoice, int]:
        item = self.table.item(row, COL_FIELD)
        combo = self.table.cellWidget(row, COL_PRIVACY)
        choice = FieldChoice(
            name=str(item.data(NAME_ROLE)),
            export=item.checkState() == Qt.CheckState.Checked,
            header=self.table.item(row, COL_HEADER).text().strip(),
            unit=self.table.item(row, COL_UNIT).text().strip(),
            description=self.table.item(row, COL_DESC).text().strip(),
            protect=str(combo.currentData()) if isinstance(combo, QComboBox) else "",
        )
        return choice, self.layer.fields().indexOf(choice.name) if self.layer else -1

    def choices(self) -> list[FieldChoice]:
        return [self._row_state(r)[0] for r in range(self.table.rowCount())]

    def _check_all(self, checked: bool) -> None:
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for row in range(self.table.rowCount()):
            self.table.item(row, COL_FIELD).setCheckState(state)

    def _move(self, step: int) -> None:
        row = self.table.currentRow()
        target = row + step
        if row < 0 or not 0 <= target < self.table.rowCount():
            return
        states = [self._row_state(r) for r in range(self.table.rowCount())]
        states[row], states[target] = states[target], states[row]
        self._filling = True
        self.table.setRowCount(0)
        for choice, index in states:
            self._add_row(choice.name, index, {"export": choice.export, "header": choice.header,
                                               "unit": choice.unit, "description": choice.description,
                                               "protect": choice.protect})
        self._filling = False
        self.table.selectRow(target)
        self._update_preview()

    def _toggle_changes(self, enabled: bool) -> None:
        self.previous_widget.setEnabled(enabled)
        self.key_combo.setEnabled(enabled)

    def scope(self) -> str:
        return next(s for s, b in self.scope_buttons.items() if b.isChecked())

    def _update_scope_labels(self, *_args: Any) -> None:
        total = self.layer.featureCount() if self.layer else 0
        selected = self.layer.selectedFeatureCount() if self.layer else 0
        self.scope_buttons["all"].setText(self.t("ui.scope.all", n=f"{total:,}".replace(",", " ")))
        self.scope_buttons["selected"].setText(self.t("ui.scope.selected", n=selected))
        self.scope_buttons["selected"].setEnabled(selected > 0)
        self.scope_buttons["extent"].setText(self.t("ui.scope.extent"))
        self.scope_buttons["extent"].setEnabled(self.iface is not None and self.layer is not None
                                                and self.layer.isSpatial())
        if not self.scope_buttons[self.scope()].isEnabled():
            self.scope_buttons["all"].setChecked(True)


    # ------------------------------------------------------------- preview

    def _update_preview(self, *_args: Any) -> None:
        if self._filling or not hasattr(self, "changes_check") or not hasattr(self, "preview"):
            return
        look = self.look()
        lang = str(self.lang_combo.currentData() or self.lang)
        chosen = [c for c in self.choices() if c.export][:4]
        columns = [layer_reader.describe(self.layer, self.layer.fields().indexOf(c.name), c)
                   for c in chosen] if self.layer else []
        side = CSS_BORDER.get(look.border_style, "")
        line = f"{side} #{look.border_color}" if side else ""

        def cell_css(style: CellLook, align: str, r: int, c: int, rows: int, cols: int, fill: str) -> str:
            css = [f"font-family:'{style.font}'", f"font-size:{style.size}pt", f"color:#{style.color}",
                   "padding:3px 7px", f"text-align:{align}"]
            if style.bold:
                css.append("font-weight:bold")
            if style.italic:
                css.append("font-style:italic")
            if style.underline:
                css.append("text-decoration:underline")
            if fill:
                css.append(f"background-color:#{fill}")
            if line:
                edges = {"top": True, "bottom": True, "left": True, "right": True}
                if look.border_scope == "horizontal":
                    edges["left"] = edges["right"] = False
                elif look.border_scope == "outline":
                    edges = {"top": r == 0, "bottom": r == rows - 1, "left": c == 0, "right": c == cols - 1}
                css += [f"border-{edge}:{line}" for edge, on in edges.items() if on]
            return "; ".join(css)

        def align_of(style: CellLook, column: Any) -> str:
            if style.align != "auto":
                return style.align
            numeric = column.kind in ("integer", "decimal") and not column.is_identifier
            return "right" if numeric else "left"

        rows = 4
        head = ""
        for i, col in enumerate(columns):
            css = cell_css(look.header, align_of(look.header, col), 0, i, rows, len(columns), look.header.fill)
            head += f"<th style=\"{css}\">{html.escape(col.header or col.name)}</th>"
        body = ""
        for r in range(1, rows):
            fill = look.stripe_color if look.stripes and r % 2 == 0 else look.body.fill
            cells = ""
            for i, (choice, col) in enumerate(zip(chosen, columns)):
                values = self.samples.get(choice.name, [])
                value = values[r - 1] if r - 1 < len(values) else None
                if choice.name in self.personal and choice.protect != "keep":
                    value = privacy.mask(value)
                text = html.escape(xlsx.display(value, col, lang)[:28])
                css = cell_css(look.body, align_of(look.body, col), r, i, rows, len(columns), fill)
                cells += f"<td style=\"{css}\">{text}</td>"
            body += f"<tr>{cells}</tr>"

        counts = {"all": self.layer.featureCount(), "selected": self.layer.selectedFeatureCount()} if self.layer else {}
        count = counts.get(self.scope())
        more = html.escape(self.t("ui.preview.more", count=count_label(lang, count))) if count is not None else ""
        sheets = self.sheets()
        names = [tr(lang, "sheet.data")]
        if sheets.dictionary:
            names.append(tr(lang, "sheet.dictionary"))
        if sheets.about:
            names.append(tr(lang, "sheet.about"))
        if self.changes_check.isChecked():
            names.append(tr(lang, "sheet.changes"))
        tabs = "".join(
            f"<span style='background:#{look.tab_color if i == 0 else 'E7E7E7'}; color:#"
            f"{'FFFFFF' if i == 0 else '404040'}; padding:2px 8px;'>&nbsp;{html.escape(n)}&nbsp;</span> "
            for i, n in enumerate(names)
        )
        title = ""
        if look.show_title:
            text = html.escape(self.title_edit.text() or (self.layer.name() if self.layer else ""))
            title = (f"<div style=\"font-family:'{look.body.font}'; font-size:{min(look.title_size, 20)}pt;"
                     f" font-weight:600; color:#{look.title_color};\">{text}</div>"
                     f"<div style='color:#595959; font-style:italic; margin-bottom:6px;'>"
                     f"{html.escape(count_label(lang, count or 0))}</div>")
        filters = look.as_table or look.autofilter
        flags = " · ".join(t for t, on in ((self.t("ui.look.as_table"), look.as_table),
                                           (self.t("ui.look.autofilter"), filters and not look.as_table),
                                           (self.t("ui.look.freeze"), look.freeze_header)) if on)
        self.preview.setText(
            f"{title}<table cellspacing='0' cellpadding='0' style='border-collapse:collapse;'>"
            f"<tr>{head}</tr>{body}</table>"
            f"<div style='color:#595959; margin-top:6px;'>{more}</div>"
            f"<div style='color:#595959; margin-top:10px;'>{html.escape(flags)}</div>"
            f"<div style='margin-top:10px;'>{tabs}</div>"
        )

    # -------------------------------------------------------------- export

    def _warn(self, text: str) -> None:
        QMessageBox.warning(self, self.t("ui.error.title"), text)

    def export(self) -> None:
        if self.layer is None:
            self._warn(self.t("ui.error.no_layer"))
            return
        if export_task.busy():
            self._warn(self.t("ui.error.busy"))
            return
        choices = self.choices()
        if not any(c.export for c in choices):
            self._warn(self.t("ui.error.no_fields"))
            return
        key: Optional[str] = None
        previous: Optional[str] = None
        if self.changes_check.isChecked():
            key = self.key_combo.currentField() or None
            previous = self.previous_widget.filePath() or None
            if key is None:
                self._warn(self.t("ui.error.no_key"))
                return
            if previous is None or not os.path.isfile(previous):
                self._warn(self.t("ui.error.no_previous"))
                return

        title = self.title_edit.text().strip() or self.layer.name()
        folder = str(self.settings.value(SETTINGS + "last_dir", os.path.expanduser("~")))
        suggested = os.path.join(folder, f"{safe_filename(title)}_{dt.date.today().isoformat()}.xlsx")
        path, _ = QFileDialog.getSaveFileName(self, self.t("ui.save.title"), suggested, self.t("ui.save.filter"))
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        lang = str(self.lang_combo.currentData())
        look, sheets = self.look(), self.sheets()
        self._remember(path)
        canvas = self.iface.mapCanvas() if self.iface is not None else None
        try:
            plan = layer_reader.plan_package(
                self.layer, choices, scope=self.scope(),
                extent=canvas.extent() if canvas is not None else None,
                extent_crs=canvas.mapSettings().destinationCrs() if canvas is not None else None,
                title=title, author=self.author_edit.text().strip(), note=self.note_edit.toPlainText().strip(),
                language=lang, key=key, personal=self.personal,
            )
        except xlsx.PackageError as exc:
            self._warn(str(exc))
            return
        # Odczyt i zapis idą w tle: okno się zamyka, QGIS zostaje do dyspozycji.
        task = export_task.ExportTask(self.t("ui.exporting", path=os.path.basename(path)), plan, path,
                                      look, sheets, previous)
        task.on_finished = self._finished
        self.task, self.layer_gone = task, False
        # Usunięta warstwa (także przy otwarciu innego projektu) nie dostaje podstawy
        # porównania; szukanie po identyfikatorze trafiłoby w nową warstwę o tym samym id.
        self.layer.willBeDeleted.connect(self._layer_deleted)
        self.progress_item = self._progress(task)
        export_task.start(task)
        self.accept()

    def _progress(self, task: export_task.ExportTask) -> Any:
        """Pasek postępu z przyciskiem „Anuluj” w pasku komunikatów QGIS."""
        if self.iface is None:
            return None
        bar = self.iface.messageBar()
        item = bar.createMessage("Data Handover", task.description())
        progress = QProgressBar()
        progress.setRange(0, 100)

        def show(value: float) -> None:
            with contextlib.suppress(RuntimeError):  # użytkownik zamknął komunikat
                progress.setValue(int(value))

        task.progressChanged.connect(show)
        cancel = QPushButton(self.t("ui.cancel"))
        cancel.clicked.connect(task.cancel)
        item.layout().addWidget(progress)
        item.layout().addWidget(cancel)
        bar.pushWidget(item, Qgis.MessageLevel.Info, 0)
        return item

    def _layer_deleted(self) -> None:
        self.layer_gone = True

    def _finished(self, task: export_task.ExportTask) -> None:
        """Koniec zadania (wątek okna): komunikat i podstawa następnego porównania."""
        if self.progress_item is not None:
            with contextlib.suppress(RuntimeError):
                self.iface.messageBar().popWidget(self.progress_item)
        if task.result is not None:
            # Dopiero udany zapis staje się podstawą następnego porównania.
            if not self.layer_gone and self.layer is not None:
                self.layer.setCustomProperty(PREVIOUS, task.result.path)
            self.last_result = task.result
            self._report(task.result)
        elif isinstance(task.error, xlsx.PackageError):
            self._warn(str(task.error))
        elif task.error is not None:
            QgsMessageLog.logMessage(task.trace, "Data Handover", Qgis.MessageLevel.Critical)
            self._warn(self.t("ui.error.write", error=task.error))
        elif self.iface is not None:
            self.iface.messageBar().pushMessage("Data Handover", self.t("ui.canceled"), Qgis.MessageLevel.Info, 5)

    def _save_state(self) -> None:
        """Ustawienia okna: pola warstwy w projekcie, wygląd i arkusze w ustawieniach użytkownika."""
        if self.layer is not None:
            layer_reader.save_choices(self.layer, self.choices(), self.key_combo.currentField() or None)
        for name, value in (("look", json.dumps(to_dict(self.look()))),
                            ("sheets", json.dumps(to_dict(self.sheets()))),
                            ("custom_color", self.custom_color), ("author", self.author_edit.text().strip()),
                            ("file_language", str(self.lang_combo.currentData()))):
            self.settings.setValue(SETTINGS + name, value)

    def _remember(self, path: str) -> None:
        assert self.layer is not None
        self._save_state()
        self.settings.setValue(SETTINGS + "last_dir", os.path.dirname(path))

    def _report(self, result: xlsx.ExportResult) -> None:
        lang = str(self.lang_combo.currentData())
        text = self.t("ui.done", count=count_label(lang, result.rows), path=result.path)
        if self.iface is None:
            return
        if result.notes:
            # Uwagi (obcięte teksty, pominięte pola osobowe, zmiany kolumn) są
            # w „Metryce” tylko wtedy, gdy ją włączono — autor musi je zobaczyć zawsze.
            text += "\n" + "\n".join(f"• {note}" for note in result.notes)
        bar = self.iface.messageBar()
        item = bar.createMessage("Data Handover", text)
        button = QPushButton(self.t("ui.open"))
        button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(result.path)))
        item.layout().addWidget(button)
        level = Qgis.MessageLevel.Warning if result.notes else Qgis.MessageLevel.Success
        bar.pushWidget(item, level, 0 if result.notes else 10)
