"""Data Handover — wtyczka QGIS: eksport tabeli atrybutów do sformatowanego skoroszytu Excela."""

import os
import sys
from typing import Any

# openpyxl dołączona do paczki (ext/) to tylko zapas: katalog trafia na koniec ścieżki importu,
# więc biblioteka z QGIS albo z systemu ma pierwszeństwo.
_EXT = os.path.join(os.path.dirname(__file__), "ext")
if _EXT not in sys.path:
    sys.path.append(_EXT)


def classFactory(iface: Any) -> Any:  # noqa: N802 - nazwa wymagana przez QGIS
    from .plugin import DataHandoverPlugin

    return DataHandoverPlugin(iface)
