"""Odczyt warstwy QGIS do modelu paczki. Działa w QGIS 3.40+ (Qt5) i 4.x (Qt6)."""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import math
import re
from dataclasses import dataclass
from dataclasses import field as dc_field
from typing import Any, Callable, Optional

from qgis.core import (
    Qgis,
    QgsApplication,
    QgsCoordinateReferenceSystem,
    QgsCoordinateTransform,
    QgsCoordinateTransformContext,
    QgsCsException,
    QgsDistanceArea,
    QgsEllipsoidUtils,
    QgsFeatureRequest,
    QgsField,
    QgsGeometry,
    QgsPointXY,
    QgsProject,
    QgsRectangle,
    QgsUnitTypes,
    QgsVectorLayer,
    QgsVectorLayerFeatureSource,
)
from qgis.PyQt.QtCore import QDate, QDateTime, QMetaType, QTime

from .core import privacy
from .core.i18n import tr
from .core.model import (
    GEO_PREFIX,
    Column,
    GeoColumns,
    Package,
    Summary,
    Value,
    from_dict,
    map_url,
    row_limit,
    to_dict,
)
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
    sum: Optional[bool] = None  # None = reguła domyślna (pole dziesiętne jest sumowane)


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


def summable(column: Column) -> bool:
    """Czy kolumnę wolno sumować: liczba, nie identyfikator, nie maska, nie etykieta, nie kolumna z geometrii."""
    return (column.kind in ("integer", "decimal") and not column.is_identifier and not column.masked
            and not column.represented and not column.name.startswith(GEO_PREFIX))


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


def saved_choices(layer: QgsVectorLayer, personal: dict[str, str]
                  ) -> tuple[list[FieldChoice], Optional[str], GeoColumns, Summary]:
    """Ustawienia pól, klucz, kolumny z geometrii i podsumowanie tak, jak okno pokazuje je po otwarciu.

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
            sum=info["sum"] if isinstance(info.get("sum"), bool) else None,
        ))
    return (choices, str(saved.get("key") or "") or None, from_dict(GeoColumns, saved.get("geometry")),
            from_dict(Summary, saved.get("summary")))


def save_choices(layer: QgsVectorLayer, choices: list[FieldChoice], key: Optional[str],
                 geo: Optional[GeoColumns] = None, summary: Optional[Summary] = None) -> None:
    """Zapis ustawień warstwy w projekcie; bez `geo` i `summary` zapisane wartości zostają."""
    saved = load_choices(layer)
    if geo is None:
        geo = from_dict(GeoColumns, saved.get("geometry"))
    if summary is None:
        summary = from_dict(Summary, saved.get("summary"))
    data = {
        "order": [c.name for c in choices],
        "key": key or "",
        "fields": {c.name: {"export": c.export, "header": c.header, "unit": c.unit,
                            "description": c.description, "protect": c.protect,
                            **({"sum": c.sum} if c.sum is not None else {})} for c in choices},
        "geometry": to_dict(geo),
        "summary": to_dict(summary),
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
             extent_crs: Optional[QgsCoordinateReferenceSystem], indices: list[int],
             geometry: bool = False) -> QgsFeatureRequest:
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
    elif not geometry:
        request.setFlags(Qgis.FeatureRequestFlag.NoGeometry)
    return request


# --------------------------------------------------------------------------
# Kolumny z geometrii (R7)
# --------------------------------------------------------------------------

_GEOMETRY_KINDS = {Qgis.GeometryType.Point: "point", Qgis.GeometryType.Line: "line",
                   Qgis.GeometryType.Polygon: "polygon"}
WGS84 = "EPSG:4326"


def geometry_kind(layer: QgsVectorLayer) -> str:
    """„polygon”, „line”, „point” albo „none” (tabela bez geometrii lub typ nieznany)."""
    return _GEOMETRY_KINDS.get(layer.geometryType(), "none") if layer.isSpatial() else "none"


def _ellipsoid(crs: QgsCoordinateReferenceSystem, project_ellipsoid: str) -> str:
    """Elipsoida projektu, a bez niej elipsoida układu warstwy; „” = żadnej nie da się użyć."""
    for candidate in (project_ellipsoid, crs.ellipsoidAcronym()):
        if candidate and candidate.upper() != "NONE" and QgsDistanceArea().setEllipsoid(candidate):
            return candidate
    return ""


def _ellipsoid_name(acronym: str) -> str:
    for definition in QgsEllipsoidUtils.definitions():
        if definition.acronym == acronym and definition.description:
            return str(definition.description)
    return acronym


@dataclass
class _GeoPlan:
    """Parametry liczenia ustalone w wątku okna; `calculator()` wołamy w wątku, który liczy."""

    crs: QgsCoordinateReferenceSystem
    context: QgsCoordinateTransformContext
    ellipsoid: str  # „” = płaszczyzna układu warstwy
    planar_factor: float  # jednostka układu → metr (tylko płaszczyzna)
    measure: str  # „area”, „length” albo „” (bez pomiaru)
    divisor: float  # m² → ha, m → km
    decimals: int
    xy: bool
    xy_decimals: int
    wgs84: bool
    link: bool
    service: str

    def calculator(self) -> tuple[Callable[[Optional[QgsGeometry]], list[Value]], list[int]]:
        """Funkcja: geometria → wartości kolumn, i licznik obiektów, których nie udało się policzyć.

        Obiekty transformacji powstają w wątku wołającego. Niepusta geometria, dla której
        pomiar, punkt wewnętrzny albo WGS84 zawiodły, dostaje puste komórki i trafia do licznika
        — autor dostaje uwagę, zamiast brać błąd obliczeń za brak geometrii.
        """
        calc = QgsDistanceArea()
        calc.setSourceCrs(self.crs, self.context)
        if self.ellipsoid:
            calc.setEllipsoid(self.ellipsoid)
        to_wgs = QgsCoordinateTransform(self.crs, QgsCoordinateReferenceSystem(WGS84), self.context)
        need_wgs = self.wgs84 or self.link
        width = (1 if self.measure else 0) + 2 * self.xy + 2 * self.wgs84 + self.link
        failed = [0]

        def wgs84(xy: QgsPointXY) -> tuple[Optional[float], Optional[float]]:
            try:
                point = to_wgs.transform(xy)
            except QgsCsException:
                return None, None
            lat, lon = point.y(), point.x()
            if not (math.isfinite(lat) and math.isfinite(lon) and abs(lat) <= 90 and abs(lon) <= 180):
                return None, None
            return round(lat, 6), round(lon, 6)

        def values(geometry: Optional[QgsGeometry]) -> list[Value]:
            if geometry is None or geometry.isNull() or geometry.isEmpty():
                return [None] * width
            out: list[Value] = []
            ok = True
            if self.measure:
                size = calc.measureArea(geometry) if self.measure == "area" else calc.measureLength(geometry)
                factor = self.planar_factor ** (2 if self.measure == "area" else 1) if not self.ellipsoid else 1.0
                ok = math.isfinite(size)
                out.append(round(size * factor / self.divisor, self.decimals) if ok else None)
            if width == len(out):
                failed[0] += not ok
                return out
            point = geometry.pointOnSurface()
            if point.isNull() or point.isEmpty():
                failed[0] += 1
                return out + [None] * (width - len(out))
            xy = point.asPoint()
            if self.xy:
                out += [round(xy.x(), self.xy_decimals), round(xy.y(), self.xy_decimals)]
            if need_wgs:
                lat, lon = wgs84(QgsPointXY(xy))
                ok = ok and lat is not None
                if self.wgs84:
                    out += [lat, lon]
                if self.link:
                    out.append(map_url(self.service, lat, lon) if lat is not None and lon is not None else None)
            failed[0] += not ok
            return out

        return values, failed


def _crs_key(crs: QgsCoordinateReferenceSystem) -> str:
    """Tożsamość układu w nazwie kolumny: kod EPSG itp., dla układu własnego — skrót definicji."""
    authid = crs.authid()
    if authid and not authid.upper().startswith("USER:"):
        return authid
    return "wkt-" + hashlib.sha1(crs.toWkt().encode("utf-8"), usedforsecurity=False).hexdigest()[:10]


def _geo_plan(layer: QgsVectorLayer, geo: GeoColumns, project_ellipsoid: str, language: str, notes: list[str],
              context: QgsCoordinateTransformContext, taken: set[str]) -> tuple[Optional[_GeoPlan], list[Column]]:
    """Kolumny z geometrii i parametry ich liczenia (wątek okna). Uwagi dopisuje do `notes`.

    Nazwa kolumny zawiera jednostkę i metodę (układ, elipsoidę): po zmianie metody porównanie
    wersji pokaże kolumnę usuniętą i dodaną, a nie fałszywą zmianę każdego wiersza. `taken` —
    nazwy już zajęte (pola warstwy); kolidująca nazwa dostaje numer.
    """
    kind = geometry_kind(layer)
    if kind == "none" or not geo.any:
        return None, []
    crs = layer.crs()
    authid = crs.authid() or crs.description() or "?"
    crs_key = _crs_key(crs)
    ellipsoid = _ellipsoid(crs, project_ellipsoid) if crs.isValid() else ""
    units = crs.mapUnits()
    planar_ok = (crs.isValid() and not crs.isGeographic()
                 and units not in (Qgis.DistanceUnit.Degrees, Qgis.DistanceUnit.Unknown))
    to_wgs_ok = crs.isValid() and QgsCoordinateTransform(crs, QgsCoordinateReferenceSystem(WGS84), context).isValid()
    columns: list[Column] = []

    def add(name: str, header: str, kind: str, **extra: Any) -> None:
        unique, n = name, 1
        while unique in taken:
            n += 1
            unique = f"{name}~{n}"
        taken.add(unique)
        columns.append(Column(unique, header, kind, **extra))

    measure = ""
    divisor, decimals = 1.0, 2
    if geo.measure and kind in ("polygon", "line"):
        if not ellipsoid and not planar_ok:
            notes.append(tr(language, "note.geo.measure_skipped", crs=authid))
        else:
            measure = "area" if kind == "polygon" else "length"
            unit = geo.area_unit if measure == "area" else geo.length_unit
            divisor, decimals, shown = {"ha": (10_000.0, 4, "ha"), "m2": (1.0, 2, "m²"),
                                        "m": (1.0, 2, "m"), "km": (1000.0, 3, "km")}[unit]
            method = (tr(language, f"geo.desc.{measure}", ellipsoid=_ellipsoid_name(ellipsoid), crs=authid)
                      if ellipsoid else tr(language, f"geo.desc.{measure}.planar", crs=authid))
            key = f"{crs_key}:{ellipsoid}" if ellipsoid else f"{crs_key}:plane"
            add(f"{GEO_PREFIX}{measure}_{unit}:{key}", tr(language, f"geo.col.{measure}"), "decimal",
                unit=shown, description=method, decimals=decimals)
    xy_decimals = 6 if crs.isGeographic() else 2
    if geo.xy:
        xy_unit = "°" if crs.isGeographic() else (
            QgsUnitTypes.toAbbreviatedString(units) if units != Qgis.DistanceUnit.Unknown else "")
        for axis in ("x", "y"):
            add(f"{GEO_PREFIX}{axis}:{crs_key}", tr(language, f"geo.col.{axis}"), "decimal", unit=xy_unit,
                decimals=xy_decimals, description=tr(language, f"geo.desc.{axis}", crs=authid))
    wgs84 = geo.wgs84 and to_wgs_ok
    link = geo.link and to_wgs_ok
    if (geo.wgs84 or geo.link) and not to_wgs_ok:
        # Nieważna transformacja zwraca punkt bez zmian: lokalne X/Y udawałyby stopnie WGS84.
        notes.append(tr(language, "note.geo.wgs_skipped", crs=authid))
    if wgs84:
        for axis in ("lat", "lon"):
            add(f"{GEO_PREFIX}{axis}", tr(language, f"geo.col.{axis}"), "decimal", unit="°",
                decimals=6, description=tr(language, f"geo.desc.{axis}"))
    if link:
        add(f"{GEO_PREFIX}link", tr(language, "geo.col.link"), "link",
            description=tr(language, "geo.desc.link", service=tr(language, f"geo.service.{geo.map}")))
    if not columns:
        return None, []
    plan = _GeoPlan(crs=QgsCoordinateReferenceSystem(crs), context=QgsCoordinateTransformContext(context),
                    ellipsoid=ellipsoid,
                    planar_factor=QgsUnitTypes.fromUnitToUnitFactor(units, Qgis.DistanceUnit.Meters)
                    if planar_ok else 1.0,
                    measure=measure, divisor=divisor, decimals=decimals, xy=geo.xy, xy_decimals=xy_decimals,
                    wgs84=wgs84, link=link, service=geo.map)
    return plan, columns


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
    geo: Optional[_GeoPlan] = None  # kolumny z geometrii dopisywane na końcu wiersza
    summary: Summary = dc_field(default_factory=Summary)

    def read(self, progress: Optional[Callable[[float], None]] = None,
             canceled: Optional[Callable[[], bool]] = None) -> Package:
        """Odczyt obiektów (może działać w tle). `progress` dostaje 0–100."""
        rows: list[list[Value]] = []
        limit = row_limit(self.columns, self.summary)
        step = max(1, self.expected // 100)
        geo_values, geo_failed = self.geo.calculator() if self.geo is not None else (None, [0])
        for feature in self.source.getFeatures(self.request):
            row: list[Value] = []
            for index, convert in zip(self.indices, self.converters):
                raw = feature.attribute(index)
                row.append(convert(raw) if convert is not None else to_python(raw))
            if geo_values is not None:
                row += geo_values(feature.geometry() if feature.hasGeometry() else None)
            rows.append(row)
            if len(rows) > limit:
                raise PackageError(tr(self.language, "error.too_many_rows", n=len(rows), limit=limit))
            if len(rows) % step == 0:
                if canceled is not None and canceled():
                    raise CanceledError
                if progress is not None:
                    progress(min(100.0, 100.0 * len(rows) / max(1, self.expected)))
        facts = [(k, str(len(rows)) if k == "count" else v) for k, v in self.facts]
        notes = self.notes + ([tr(self.language, "note.geo.failed", n=geo_failed[0])] if geo_failed[0] else [])
        return Package(title=self.title, columns=self.columns, rows=rows, facts=facts,
                       notes=notes, language=self.language, key=self.key, summary=self.summary)


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
    geo: Optional[GeoColumns] = None,
    ellipsoid: str = "",
    transform_context: Optional[QgsCoordinateTransformContext] = None,
    summary: Optional[Summary] = None,
) -> PackagePlan:
    """Plan paczki z warstwy (w wątku okna). `personal` — wynik `personal_data()`, jeśli już policzony.

    `geo` — kolumny z geometrii; `ellipsoid` i `transform_context` — z ustawień projektu albo
    z kontekstu algorytmu Processing (domyślnie kontekst transformacji bieżącego projektu).
    """
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
        column.summed = summable(column) and (choice.sum if choice.sum is not None else column.kind == "decimal")
        columns.append(column)
        indices.append(index)
        converters.append(representer(layer, index, stand_in))

    if key is not None and key not in [c.name for c in columns]:
        key = None

    context = transform_context if transform_context is not None else QgsProject.instance().transformContext()
    geo_plan, geo_columns = _geo_plan(layer, geo or GeoColumns(), ellipsoid, language, notes, context,
                                      set(fields.names()))
    summary = summary if summary is not None else Summary()
    for column in geo_columns:
        column.summed = summary.measure and column.name.startswith((f"{GEO_PREFIX}area_", f"{GEO_PREFIX}length_"))
    columns += geo_columns
    request = _request(layer, scope, extent, extent_crs, indices, geometry=geo_plan is not None)
    limit = row_limit(columns, summary)
    if scope == "selected" and layer.selectedFeatureCount() > limit:
        raise PackageError(tr(language, "error.too_many_rows", n=layer.selectedFeatureCount(), limit=limit))
    if scope == "all" and layer.featureCount() > limit:
        raise PackageError(tr(language, "error.too_many_rows", n=layer.featureCount(), limit=limit))

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
                       notes, title or layer.name(), language, key, expected, geo_plan, summary)
