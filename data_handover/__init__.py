"""Data Handover — wtyczka QGIS: eksport tabeli atrybutów do sformatowanego skoroszytu Excela."""

from typing import Any


def classFactory(iface: Any) -> Any:  # noqa: N802 - nazwa wymagana przez QGIS
    from .plugin import DataHandoverPlugin

    return DataHandoverPlugin(iface)
