# Changelog

All notable changes to the Data Handover plugin. Versions follow
`version` in [`data_handover/metadata.txt`](data_handover/metadata.txt).

## 1.0.2

- Fix: the window failed to open (`TypeError ... ValueRelationCache`) on a
  layer with a Value Relation field. Labels now come from the referenced
  layer as in the attribute form (the layer is found the same way as QGIS
  finds it, and the widget filter applies unless it depends on the form),
  read once before the export; with multiple values allowed they are joined
  with commas. A code missing from the referenced layer, or a missing
  referenced layer, is written as the code itself.
- The file language follows the plugin language: switching the plugin
  language also switches the file language, and **Restore defaults** sets
  it back to the plugin language. Choosing a different file language never
  changes the window language. The Processing algorithm uses the plugin
  language when no file language has been saved yet.
- The window fits small laptop screens and 4K monitors with display
  scaling (checked in all ten languages down to 683 × 384 logical pixels,
  e.g. 1366 × 768 at 200 %). It no longer opens taller or wider than the
  screen; when it is smaller than its content, the tabs scroll instead of
  squeezing the fields, and on a very narrow screen the bottom bar takes two
  rows. The field table scrolls sideways, and the description column keeps a
  usable width.
- Look tab: a divider between the settings and the preview, and **+** / **−**
  buttons to zoom the preview (50–200 %).
- The plugin remembers the window size, the divider position and the preview
  zoom; **Restore defaults** resets them.
- A shorter plugin description and changelog in the Plugin Manager.

## 1.0.1

- New plugin icon (a table handed over on an open palm), also in the window
  header and title bar.
- The **openpyxl** library (3.1.5, MIT licence) and its dependency
  et_xmlfile (2.0.0, MIT) are bundled in `data_handover/ext`. They are used
  only when QGIS does not provide its own openpyxl, so the plugin works right
  after installation from ZIP, with no extra packages.
- Fix (Linux, macOS): exporting over an existing file keeps its permissions.
  Before, a file readable only by its owner became readable by others after
  the next export to the same path.

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
