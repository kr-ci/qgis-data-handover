"""Algorytm Processing „Paczka dla odbiorcy”: ten sam plik co z okna, bez okna.

Do eksportu wsadowego, modelu graficznego i `qgis_process` (harmonogram).
Kolumny, nagłówki, opisy, kolejność, ochronę danych osobowych i klucz bierze
z ustawień warstwy zapisanych w projekcie przez okno wtyczki; wygląd, arkusze
dodatkowe, autora i język pliku — z ustawień okna w profilu użytkownika.

Moduł nie importuje openpyxl na poziomie modułu (także pośrednio:
`layer_reader` → `core.xlsx`): dostawcę rejestruje `initGui()`, a bez
openpyxl wtyczka musi się wczytać. Importy rdzenia są w metodach.
"""

from __future__ import annotations

import os
from typing import Any, Optional

from qgis.core import (
    Qgis,
    QgsProcessingAlgorithm,
    QgsProcessingContext,
    QgsProcessingException,
    QgsProcessingFeedback,
    QgsProcessingParameterBoolean,
    QgsProcessingParameterEnum,
    QgsProcessingParameterField,
    QgsProcessingParameterFile,
    QgsProcessingParameterFileDestination,
    QgsProcessingParameterVectorLayer,
    QgsProcessingProvider,
    QgsSettings,
)
from qgis.PyQt.QtGui import QIcon

from .core.i18n import LANGUAGE_NAMES, LANGUAGES, supported, tr
from .core.model import THEMES
from .plugin import SETTINGS, saved_json, ui_language

ICON = os.path.join(os.path.dirname(__file__), "icon.svg")


def _t(text_id: str, **kwargs: object) -> str:
    return tr(ui_language(), text_id, **kwargs)


class PackageAlgorithm(QgsProcessingAlgorithm):
    INPUT = "INPUT"
    SELECTED_ONLY = "SELECTED_ONLY"
    LANGUAGE = "LANGUAGE"
    THEME = "THEME"
    PREVIOUS = "PREVIOUS"
    KEY = "KEY"
    OUTPUT = "OUTPUT"

    def __init__(self) -> None:
        super().__init__()
        self.plan: Any = None
        self.look: Any = None
        self.sheets: Any = None
        self.previous: Optional[str] = None

    def name(self) -> str:
        return "package"

    def displayName(self) -> str:  # noqa: N802 - nazwa wymagana przez QGIS
        return _t("alg.name")

    def shortHelpString(self) -> str:  # noqa: N802
        return _t("alg.help")

    def icon(self) -> QIcon:
        return QIcon(ICON)

    def createInstance(self) -> PackageAlgorithm:  # noqa: N802
        return PackageAlgorithm()

    def initAlgorithm(self, config: Any = None) -> None:  # noqa: N802, ARG002
        # Pozycja 0 list „język” i „motyw” = jak w oknie: wartość czytana przy każdym
        # uruchomieniu, bo ustawienia okna mogą się zmienić po starcie QGIS.
        self.addParameter(QgsProcessingParameterVectorLayer(
            self.INPUT, _t("alg.input"), [Qgis.ProcessingSourceType.Vector]))
        self.addParameter(QgsProcessingParameterBoolean(self.SELECTED_ONLY, _t("alg.selected"), False))
        self.addParameter(QgsProcessingParameterEnum(
            self.LANGUAGE, _t("ui.file.language"),
            [_t("alg.as_window")] + [LANGUAGE_NAMES[code] for code in LANGUAGES], defaultValue=0))
        self.addParameter(QgsProcessingParameterEnum(
            self.THEME, _t("ui.file.theme"),
            [_t("alg.as_window")] + [_t(f"ui.theme.{name}") for name in THEMES], defaultValue=0))
        self.addParameter(QgsProcessingParameterFile(
            self.PREVIOUS, _t("ui.changes.file"), extension="xlsx", optional=True))
        self.addParameter(QgsProcessingParameterField(
            self.KEY, _t("ui.changes.key"), parentLayerParameterName=self.INPUT, optional=True))
        self.addParameter(QgsProcessingParameterFileDestination(
            self.OUTPUT, _t("alg.output"), _t("ui.save.filter")))

    def prepareAlgorithm(self, parameters: dict[str, Any], context: QgsProcessingContext,  # noqa: N802
                         feedback: QgsProcessingFeedback) -> bool:  # noqa: ARG002
        """Wątek okna: tu czytamy warstwę i ustawienia. W tle zostaje tylko odczyt kopii i zapis."""
        try:
            import openpyxl  # noqa: F401
        except ImportError as exc:
            raise QgsProcessingException(_t("ui.error.openpyxl")) from exc
        from . import layer_reader
        from .core.model import Look, Sheets, from_dict, look_with_theme
        from .core.xlsx import PackageError

        layer = self.parameterAsVectorLayer(parameters, self.INPUT, context)
        if layer is None:
            raise QgsProcessingException(_t("ui.error.no_layer"))
        settings = QgsSettings()
        lang_index = self.parameterAsEnum(parameters, self.LANGUAGE, context)
        language = (LANGUAGES[lang_index - 1] if lang_index > 0
                    else supported(str(settings.value(SETTINGS + "file_language", "en") or "en")))
        self.look = from_dict(Look, saved_json("look"))
        theme_index = self.parameterAsEnum(parameters, self.THEME, context)
        if theme_index > 0:
            self.look = look_with_theme(self.look, THEMES[list(THEMES)[theme_index - 1]])
        self.sheets = from_dict(Sheets, saved_json("sheets"))

        personal = layer_reader.personal_data(layer)
        choices, saved_key, geo, summary = layer_reader.saved_choices(layer, personal)
        if not any(c.export for c in choices):
            raise QgsProcessingException(_t("alg.error.no_fields"))
        scope = "selected" if self.parameterAsBoolean(parameters, self.SELECTED_ONLY, context) else "all"
        if scope == "selected" and layer.selectedFeatureCount() == 0:
            raise QgsProcessingException(_t("alg.error.no_selection"))

        # Jak w oknie: klucz liczy się tylko razem z poprzednim plikiem.
        key: Optional[str] = None
        self.previous = self.parameterAsFile(parameters, self.PREVIOUS, context) or None
        if self.previous is not None:
            key = self.parameterAsString(parameters, self.KEY, context) or saved_key
            if not key:
                raise QgsProcessingException(_t("ui.error.no_key"))
            if not os.path.isfile(self.previous):
                raise QgsProcessingException(_t("ui.error.no_previous"))
        try:
            self.plan = layer_reader.plan_package(
                layer, choices, scope=scope, author=str(settings.value(SETTINGS + "author", "") or "").strip(),
                language=language, key=key, personal=personal, geo=geo, summary=summary,
                ellipsoid=context.ellipsoid(), transform_context=context.transformContext())
        except PackageError as exc:
            raise QgsProcessingException(str(exc)) from exc
        return True

    def processAlgorithm(self, parameters: dict[str, Any], context: QgsProcessingContext,  # noqa: N802
                         feedback: QgsProcessingFeedback) -> dict[str, Any]:
        """Może działać w tle: bez dotykania warstwy (tylko kopia źródła z planu)."""
        from .core import xlsx

        path = self.parameterAsFileOutput(parameters, self.OUTPUT, context)
        try:
            # Podział paska jak w oknie (export_task.py): 10% odczyt, 90% zapis.
            pkg = self.plan.read(lambda p: feedback.setProgress(p / 10), feedback.isCanceled)
            result = xlsx.export(path, pkg, self.look, self.sheets, previous=self.previous,
                                 progress=lambda p: feedback.setProgress(10 + p * 0.9), canceled=feedback.isCanceled)
        except xlsx.CanceledError as exc:
            raise QgsProcessingException(_t("ui.canceled")) from exc
        except xlsx.PackageError as exc:
            raise QgsProcessingException(str(exc)) from exc
        except OSError as exc:  # brak uprawnień, plik otwarty w Excelu
            raise QgsProcessingException(_t("ui.error.write", error=exc)) from exc
        for note in result.notes:
            feedback.pushWarning(note)
        return {self.OUTPUT: result.path}


class Provider(QgsProcessingProvider):
    def id(self) -> str:
        return "dataHandover"

    def name(self) -> str:
        return "Data Handover"

    def icon(self) -> QIcon:
        return QIcon(ICON)

    def loadAlgorithms(self) -> None:  # noqa: N802
        self.addAlgorithm(PackageAlgorithm())
