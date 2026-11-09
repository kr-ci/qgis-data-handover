"""Odczyt warstwy QGIS do modelu paczki. Działa w QGIS 3.40+ (Qt5) i 4.x (Qt6)."""

from __future__ import annotations

import datetime as dt
import json
import re
from dataclasses import dataclass
from typing import Any, Callable, Optional

from qgis.core import (
    Qgis,
    QgsApplication,
    QgsCoordinateReferenceSystem,
    QgsCoordinateTransform,
    QgsFeatureRequest,
    QgsField,
    QgsProject,
    QgsRectangle,
    QgsVectorLayer,
    QgsVectorLayerFeatureSource,
)
from qgis.PyQt.QtCore import QDate, QDateTime, QMetaType, QTime

from .core import privacy
from .core.i18n import tr
from .core.model import MAX_ROWS, Column, Package, Value
from .core.xlsx import CanceledError, PackageError

# Kody typów są wspólne dla QVariant.Type (Qt5) i QMetaType.Type (Qt6).
_KIND_BY_TYPE = {1: "boolean", 2: "integer", 3: "integer", 4: "integer", 5: "integer", 6: "decimal",
                 38: "decimal", 14: "date", 15: "time", 16: "datetime"}
# Widżety, których wartość pokazujemy tak, jak w formularzu QGIS (etykieta zamiast kodu).
_REPRESENTED_WIDGETS = {"ValueMap", "ValueRelation", "RelationReference", "CheckBox"}
_ID_TOKENS = {"id", "fid", "gid", "ogc", "nr", "numer", "kod", "code", "teryt", "objectid", "idx"}
_VALUE_MAP_NULL = "{2839923C-8B7D-419E-B84B-CA2FE9B80EC7}"
PROPERTY = "data_handover/fields"
SAVED_PROTECTION = "zapisana decyzja o ochronie"  # powód w dymku okna, obok „nazwa: …”, „wartości: …”
SAMPLE = 200


@dataclass
class FieldChoice:
    """Ustawienia pola z okna wtyczki."""

    name: str
    export: bool = True
    header: str = ""
    unit: str = ""
    description: str = ""
    protect: str = ""  # "" | "mask" | "keep" — tylko dla pól z danymi osobowymi


def type_code(field: QgsField) -> int:
    kind = field.type()
    return int(getattr(kind, "value", kind))


def field_kind(field: QgsField) -> str:
    return _KIND_BY_TYPE.get(type_code(field), "text")


def is_identifier(layer: QgsVectorLayer, index: int) -> bool:
    field = layer.fields().at(index)
    if index in layer.primaryKeyAttributes():
        return True
    tokens = [t for t in re.split(r"[\W_]+", field.name().lower()) if t]
    return field_kind(field) in ("integer", "text") and any(t in _ID_TOKENS for t in tokens)


def to_python(value: Any) -> Value:
    """Wartość atrybutu QGIS → typ Pythona. NULL w Qt5 to QVariant, w Qt6 None."""
    if value is None:
        return None
    if hasattr(value, "isNull") and not isinstance(value, (QDate, QDateTime, QTime)) and value.isNull():
        return None
    if isinstance(value, QDateTime):
        return value.toPyDateTime() if value.isValid() else None
    if isinstance(value, QDate):
        return value.toPyDate() if value.isValid() else None
    if isinstance(value, QTime):
        return value.toPyTime() if value.isValid() else None
    if isinstance(value, (bool, int, float, str, dt.date, dt.time)):
        return value
    if isinstance(value, (list, dict)):
        return json.dumps(value, ensure_ascii=False, default=str)
    if isinstance(value, (bytes, bytearray)) or type(value).__name__ == "QByteArray":
        return f"[{len(value)} B]"
    return str(value)


def value_map_codes(config: dict[str, Any]) -> list[tuple[str, str]]:
    """Konfiguracja widżetu ValueMap → lista (kod, etykieta)."""
    raw = config.get("map", [])
    pairs: list[tuple[str, str]] = []
    items = raw if isinstance(raw, list) else [raw]
    for item in items:
        if isinstance(item, dict):
            pairs += [(str(code), str(label)) for label, code in item.items() if str(code) != _VALUE_MAP_NULL]
    return pairs


def load_choices(layer: QgsVectorLayer) -> dict[str, dict[str, Any]]:
    """Zapamiętane w projekcie ustawienia pól tej warstwy."""
    raw = layer.customProperty(PROPERTY, "")
    try:
        data = json.loads(raw) if raw else {}
    except (TypeError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def saved_choices(layer: QgsVectorLayer, personal: dict[str, str]) -> tuple[list[FieldChoice], Optional[str]]:
    """Ustawienia pól i klucz tak, jak okno pokazuje je po otwarciu.

    Pole bez zapisanych ustawień: eksportowane, chyba że wygląda na dane
    osobowe; pole osobowe eksportowane jawnie — domyślnie maskowane.
    """
    saved = load_choices(layer)
    stored = saved.get("fields", {})
    stored = stored if isinstance(stored, dict) else {}
    names = layer.fields().names()
    order = [n for n in dict.fromkeys(saved.get("order", [])) if n in names]
    order += [n for n in names if n not in order]
    choices = []
    for name in order:
        info = stored.get(name)
        info = info if isinstance(info, dict) else {}
        choices.append(FieldChoice(
            name=name,
            export=bool(info.get("export", name not in personal)),
            header=str(info.get("header") or "").strip(),
            unit=str(info.get("unit") or "").strip(),
            description=str(info.get("description") or "").strip(),
            protect=str(info.get("protect") or "mask") if name in personal else "",
        ))
    return choices, str(saved.get("key") or "") or None


def save_choices(layer: QgsVectorLayer, choices: list[FieldChoice], key: Optional[str]) -> None:
    data = {
        "order": [c.name for c in choices],
        "key": key or "",
        "fields": {c.name: {"export": c.export, "header": c.header, "unit": c.unit,
                            "description": c.description, "protect": c.protect} for c in choices},
    }
    layer.setCustomProperty(PROPERTY, json.dumps(data, ensure_ascii=False))


def fields_only(layer: QgsVectorLayer) -> QgsVectorLayer:
    """Prywatna warstwa w pamięci z tymi samymi polami, bez obiektów.

    Formatter etykiet sięga do warstwy tylko po jej pola. W tle dostaje tę
    kopię zamiast oryginału, bo oryginał użytkownik może w trakcie eksportu
    usunąć z projektu — wywołanie na usuniętej warstwie wysypałoby QGIS.
    """
    copy = QgsVectorLayer("None", "fields", "memory")
    provider = copy.dataProvider()
    for field in layer.fields():
        before = provider.fields().count()
        provider.addAttributes([QgsField(field)])
        if provider.fields().count() == before:
            # Typu nieznanego warstwie w pamięci nie wolno pominąć — przesunąłby numery
            # pól; tekst zachowuje numer, a takie pole i tak nie ma etykiet z widżetu.
            provider.addAttributes([QgsField(field.name(), QMetaType.Type.QString)])
    copy.updateFields()
    if copy.fields().names() != layer.fields().names():
        raise RuntimeError(f"Nie udało się skopiować pól warstwy: {layer.fields().names()}")
    return copy


def representer(layer: QgsVectorLayer, index: int,
                target: Optional[QgsVectorLayer] = None) -> Optional[Callable[[Any], Value]]:
    """Funkcja zamieniająca kod na etykietę (jak w formularzu) albo None.

    Pamięć podręczna etykiet powstaje tu, z `layer`; `target` (zob. `fields_only`)
    dostaje formatter przy każdym wywołaniu — domyślnie ta sama warstwa.
    """
    setup = layer.editorWidgetSetup(index)
    if setup.type() not in _REPRESENTED_WIDGETS:
        return None
    config = setup.config()
    formatter = QgsApplication.fieldFormatterRegistry().fieldFormatter(setup.type())
    cache = formatter.createCache(layer, index, config)
    if setup.type() == "RelationReference" and to_python(cache) is None:
        # Bez pamięci podręcznej (np. relacja wskazuje usuniętą warstwę) formatter
        # szukałby relacji w projekcie i czytał warstwę powiązaną — w tle to
        # niebezpieczne. Wynik byłby i tak ten sam: wartość jako tekst.
        return lambda raw: None if to_python(raw) is None else str(to_python(raw))
    target = layer if target is None else target

    def represent(raw: Any) -> Value:
        if to_python(raw) is None:
            return None
        return to_python(formatter.representValue(target, index, config, cache, raw))

    return represent


def sample_values(layer: QgsVectorLayer, limit: int = SAMPLE, represented: bool = False) -> dict[str, list[Value]]:
    """Pierwsze `limit` wartości każdego pola; z `represented` — etykiety zamiast kodów."""
    request = QgsFeatureRequest().setLimit(limit).setFlags(Qgis.FeatureRequestFlag.NoGeometry)
    names = layer.fields().names()
    convert = [representer(layer, i) if represented else None for i in range(len(names))]
    samples: dict[str, list[Value]] = {n: [] for n in names}
    for feature in layer.getFeatures(request):
        for name, value, fn in zip(names, feature.attributes(), convert):
            samples[name].append(fn(value) if fn is not None else to_python(value))
    return samples


def personal_data(layer: QgsVectorLayer, samples: Optional[dict[str, list[Value]]] = None) -> dict[str, str]:
    """{nazwa pola: powód} dla pól z danymi osobowymi — wykrytych albo już chronionych.

    Próbki z etykietami (jak w oknie), żeby okno i algorytm Processing oceniły
    pola tak samo. Pole z zapisaną decyzją o ochronie zostaje osobowe, choćby
    dzisiejsza próbka go nie wykryła (np. filtr zostawił mniej niż trzy wartości):
    inaczej zapisane maskowanie przestałoby działać bez ostrzeżenia.
    """
    names = layer.fields().names()
    fields = [(name, layer.attributeDisplayName(i)) for i, name in enumerate(names)]
    samples = sample_values(layer, represented=True) if samples is None else samples
    found = {f.field: f.reason for f in privacy.detect(fields, samples)}
    stored = load_choices(layer).get("fields", {})
    for name, info in (stored.items() if isinstance(stored, dict) else []):
        if name in names and isinstance(info, dict) and info.get("protect") in ("mask", "keep"):
            found.setdefault(name, SAVED_PROTECTION)
    return found


def describe(layer: QgsVectorLayer, index: int, choice: FieldChoice) -> Column:
    """Kolumna paczki dla pola warstwy (bez decyzji o prywatności)."""
    field = layer.fields().at(index)
    setup = layer.editorWidgetSetup(index)
    represented = setup.type() in _REPRESENTED_WIDGETS
    kind = "text" if represented else field_kind(field)
    precision = field.precision()
    return Column(
        name=choice.name,
        header=choice.header or layer.attributeDisplayName(index),
        kind=kind,
        unit=choice.unit,
        description=choice.description or field.comment(),
        decimals=precision if kind == "decimal" and 0 < precision <= 6 else 2,
        is_identifier=is_identifier(layer, index) and not represented,
        represented=represented,
        codes=value_map_codes(setup.config()) if setup.type() == "ValueMap" else [],
    )


def _request(layer: QgsVectorLayer, scope: str, extent: Optional[QgsRectangle],
             extent_crs: Optional[QgsCoordinateReferenceSystem], indices: list[int]) -> QgsFeatureRequest:
    request = QgsFeatureRequest().setSubsetOfAttributes(indices)
    if scope == "selected":
        request.setFilterFids(layer.selectedFeatureIds())
    if scope == "extent" and extent is not None:
        rect = QgsRectangle(extent)
        if extent_crs is not None and extent_crs.isValid() and extent_crs != layer.crs():
            transform = QgsCoordinateTransform(extent_crs, layer.crs(), QgsProject.instance())
            rect = transform.transformBoundingBox(rect)
        request.setFilterRect(rect)
        request.setFlags(Qgis.FeatureRequestFlag.ExactIntersect)
    else:
        request.setFlags(Qgis.FeatureRequestFlag.NoGeometry)
    return request


@dataclass
class PackagePlan:
    """Wszystko, co trzeba ustalić w wątku okna, zanim odczyt pójdzie w tle.

    Warstwy QGIS nie wolno czytać z innego wątku — `source` to jej kopia do
    odczytu (`QgsVectorLayerFeatureSource`), razem z filtrem i edycjami.
    Pamięć podręczna etykiet (value map, relacje) powstaje tu, w wątku okna;
    w tle wywołujemy już tylko `representValue` na kopii pól (`fields_only`).
    """

    source: QgsVectorLayerFeatureSource
    request: QgsFeatureRequest
    indices: list[int]
    converters: list[Optional[Callable[[Any], Value]]]
    columns: list[Column]
    facts: list[tuple[str, str]]
    notes: list[str]
    title: str
    language: str
    key: Optional[str]
    expected: int  # przewidywana liczba obiektów — do paska postępu

    def read(self, progress: Optional[Callable[[float], None]] = None,
             canceled: Optional[Callable[[], bool]] = None) -> Package:
        """Odczyt obiektów (może działać w tle). `progress` dostaje 0–100."""
        rows: list[list[Value]] = []
        step = max(1, self.expected // 100)
        for feature in self.source.getFeatures(self.request):
            row: list[Value] = []
            for index, convert in zip(self.indices, self.converters):
                raw = feature.attribute(index)
                row.append(convert(raw) if convert is not None else to_python(raw))
            rows.append(row)
            if len(rows) > MAX_ROWS:
                raise PackageError(tr(self.language, "error.too_many_rows", n=len(rows), limit=MAX_ROWS))
            if len(rows) % step == 0:
                if canceled is not None and canceled():
                    raise CanceledError
                if progress is not None:
                    progress(min(100.0, 100.0 * len(rows) / max(1, self.expected)))
        facts = [(k, str(len(rows)) if k == "count" else v) for k, v in self.facts]
        return Package(title=self.title, columns=self.columns, rows=rows, facts=facts,
                       notes=self.notes, language=self.language, key=self.key)


def build_package(layer: QgsVectorLayer, choices: list[FieldChoice], **kwargs: Any) -> Package:
    """Składa paczkę z warstwy w bieżącym wątku (argumenty jak w `plan_package`)."""
    return plan_package(layer, choices, **kwargs).read()


def plan_package(
    layer: QgsVectorLayer,
    choices: list[FieldChoice],
    *,
    scope: str = "all",
    extent: Optional[QgsRectangle] = None,
    extent_crs: Optional[QgsCoordinateReferenceSystem] = None,
    title: str = "",
    author: str = "",
    note: str = "",
    language: str = "pl",
    key: Optional[str] = None,
    personal: Optional[dict[str, str]] = None,
) -> PackagePlan:
    """Plan paczki z warstwy (w wątku okna). `personal` — wynik `personal_data()`, jeśli już policzony."""
    fields = layer.fields()
    personal = personal_data(layer) if personal is None else personal
    columns: list[Column] = []
    indices: list[int] = []
    converters: list[Optional[Callable[[Any], Value]]] = []
    notes: list[str] = []
    stand_in = fields_only(layer)

    for choice in choices:
        index = fields.indexOf(choice.name)
        if index < 0:
            continue
        if choice.name in personal and not choice.export:
            notes.append(tr(language, "note.privacy.skipped", field=layer.attributeDisplayName(index)))
            continue
        if not choice.export:
            continue
        masked = choice.name in personal and choice.protect != "keep"
        if choice.name in personal:
            key_note = "note.privacy.masked" if masked else "note.privacy.kept"
            notes.append(tr(language, key_note, field=choice.header or layer.attributeDisplayName(index)))
        column = describe(layer, index, choice)
        column.masked = masked
        if masked:
            column.kind = "text"
            column.is_identifier = False
        columns.append(column)
        indices.append(index)
        converters.append(representer(layer, index, stand_in))

    if key is not None and key not in [c.name for c in columns]:
        key = None

    request = _request(layer, scope, extent, extent_crs, indices)
    if scope == "selected" and layer.selectedFeatureCount() > MAX_ROWS:
        raise PackageError(tr(language, "error.too_many_rows", n=layer.selectedFeatureCount(), limit=MAX_ROWS))
    if scope == "all" and layer.featureCount() > MAX_ROWS:
        raise PackageError(tr(language, "error.too_many_rows", n=layer.featureCount(), limit=MAX_ROWS))

    # Filtr warstwy wpisany do „Metryki” mógłby zdradzić wartość pola osobowego,
    # którego nie eksportujemy jawnie (np. "email" = 'anna@example.com').
    plain = {c.name for c in columns if not c.masked}
    subset = layer.subsetString()
    if privacy.mentions(subset, [name for name in personal if name not in plain]):
        subset = tr(language, "filter.hidden")

    crs = layer.crs()
    facts = [
        ("title", title or layer.name()),
        ("layer", layer.name()),
        ("note", note),
        ("source", layer.publicSource()),
        ("provider", layer.providerType()),
        ("crs", f"{crs.authid()} — {crs.description()}" if crs.isValid() else ""),
        ("scope", tr(language, f"scope.{scope}")),
        ("filter", subset),
        ("count", ""),  # uzupełnia PackagePlan.read()
        ("total", str(layer.featureCount())),
        ("author", author),
        ("qgis", f"QGIS {Qgis.version()} · Data Handover"),
    ]
    expected = layer.selectedFeatureCount() if scope == "selected" else layer.featureCount()
    return PackagePlan(QgsVectorLayerFeatureSource(layer), request, indices, converters, columns, facts,
                       notes, title or layer.name(), language, key, expected)
