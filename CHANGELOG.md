# Changelog

All notable changes to the Data Handover plugin. Versions follow
`version` in [`data_handover/metadata.txt`](data_handover/metadata.txt).

## 1.0.1

- New plugin icon (a table handed over on an open palm), also in the window
  header and title bar.
- The **openpyxl** library (3.1.5, MIT licence) and its dependency
  et_xmlfile (2.0.0, MIT) are bundled in `data_handover/ext`. They are used
  only when QGIS does not provide its own openpyxl, so the plugin works right
  after installation from ZIP, with no extra packages.

## 1.0.0 — first public release

- **Data** sheet formatted to your settings: Excel table or plain range,
  fonts, colours, fills, alignment, wrapping, banded rows, borders, title,
  frozen header, column width, print orientation, tab colour; colour presets
  and a live preview.
- Readable content: headers from aliases, value map and relation labels
  instead of codes, numbers with their unit kept as numbers, identifiers
  (TERYT, parcel numbers) without thousands separators and never summed
  (text identifiers keep their leading zeros), years without a thousands
  separator.
- Optional sheets, off by default: **Dictionary** (column descriptions),
  **About** (source, date, CRS, scope, filter, author, notes), **Summary**
  (counts, sums, breakdown by one column, missing data) and **Changes**
  (added, removed and changed features since the previous export, with
  changed cells highlighted).
- Personal data: fields that look personal (by name or by values: PESEL,
  e-mail, phone, land register number) are left out or masked by default.
- Columns from geometry: area or length on the ellipsoid, X/Y in the layer
  CRS, WGS 84 latitude/longitude, a "Show on map" link to OpenStreetMap or
  Geoportal.
- Totals row below the table (`SUBTOTAL`, follows Excel filters).
- Export runs in the background with progress and a Cancel button.
- Processing algorithm `dataHandover:package` for batch mode, graphical
  models and `qgis_process` (title = layer name, no author's note).
- Window and file contents in 10 languages: English, German, French,
  Spanish, Portuguese, Polish, Russian, Korean, Japanese, Simplified Chinese.
- QGIS 3.40–3.44 (Qt5) and 4.x (Qt6) from one code base.
