# Data Handover — QGIS plugin

**English** · [Polski](README.pl.md)

A QGIS plugin that exports the attribute table of a vector layer to a
formatted Excel workbook — styled the way you choose, with readable
headers and labels instead of codes.

## What you set

| Tab | What it holds |
|---|---|
| **Data** | layer, scope (all / selected / in map extent), columns and their order, header, unit and description of each column (saved in the project), the decision about personal data |
| **Look** | Excel table or plain range (with or without filters); separately for headers and content: font, size, colour, bold, italic, underline, fill, alignment, wrapping; banded rows; borders (line, colour, all lines / horizontal only / outline only); title above the data, frozen header, gridlines, description in a header comment, column width, print orientation, tab colour; colour presets and a live preview |
| **Sheets** | title, file language (10 languages, English by default), extra sheets — off by default — with a choice of what goes in |
| **Changes** | comparison with the previous export by a key column |

## What is in the file

| Sheet | What it holds |
|---|---|
| **Data** | the data formatted according to your settings; headers from aliases; labels instead of codes (value map, relations); numbers with a unit that are still numbers; identifiers (e.g. TERYT, parcel numbers) kept as text; ready to print |
| **Dictionary** (optional) | chosen items for each column: source field, type, unit, description, codes and labels |
| **About** (optional) | chosen items: data as of date, layer, source, coordinate reference system, scope, filter, feature count, author, reading guide, release notes |
| **Changes** (optional) | features added, removed and changed since the previous file, with "was / is" values; changed cells highlighted in the data |

Fields that look like personal data (by field name or by values: Polish
PESEL with a valid checksum, e-mail, phone number, land register number)
are unchecked by default; you can mask them or export them on purpose.
Detection is a hint, not a guarantee of anonymisation.

Export runs in the background: QGIS stays usable, and the message bar
shows progress with a **Cancel** button. Cancelling leaves the target
file as it was. With very large lookup layers (Value Relation, Relation
Reference fields) QGIS may pause briefly before the export starts, while
the labels are loaded.

The **Restore defaults** button at the bottom of the window resets the
look, the sheets and the file language (column descriptions saved in the
project stay). The window and the file are available in English, German,
French, Spanish, Portuguese, Polish, Russian, Korean, Japanese and
Simplified Chinese. Pick the window language from the "Plugin language"
list in the footer (English by default); the file language is a separate
choice on the Sheets tab. Column headers and values come from the layer,
so they stay in the language of the data.

## Processing, batch mode and `qgis_process`

The same export is available as a Processing algorithm, **Data Handover →
Hand over data (XLSX)** (`dataHandover:package`): in the Processing
Toolbox, in batch mode, in graphical models and from the command line.
It produces the same file as the window. Columns, headers, units,
descriptions, column order, personal data choices and the key column come
from the layer settings the window saved in the project; the look, extra
sheets, author and file language come from the window's saved settings.
Parameters: layer, selected features only, file language and colour preset
(both "As in the plugin window" by default), previous file and key column
(for the Changes sheet; an empty key means the key saved in the window).

```bash
qgis_process plugins enable data_handover
qgis_process help dataHandover:package      # parameters and their values
qgis_process run dataHandover:package --PROJECT_PATH=project.qgz -- \
    INPUT=parcels OUTPUT=/data/parcels.xlsx
```

`qgis_process` reads the settings of the QGIS user profile, so run it as
the same system user who set up the window. Without that profile the
default look applies; columns still come from the project.

## Requirements and installation

- QGIS 3.40–3.44 or 4.x (tested: 3.40, 3.44.15, 4.2.3).
- The **openpyxl** Python library — included in the QGIS standalone
  installer for Windows; in OSGeo4W add the `python3-openpyxl` package, on
  Linux install `python3-openpyxl`. Without it the plugin shows
  installation instructions.

Install from a file: `scripts/package.sh` builds
`build/data_handover-<version>.zip`; in QGIS: *Plugins → Manage and Install
Plugins → Install from ZIP*. The plugin appears on the toolbar, in the
*Plugins* menu and in the context menu of a vector layer.

## Development

```bash
uvx ruff check .
uv run --no-project --python 3.12 --with mypy --with openpyxl --with types-openpyxl mypy data_handover/core
uv run --no-project --python 3.12 --with pytest --with openpyxl pytest -q
scripts/qgis-test.sh 4.2.3          # test in real QGIS (Docker); CI also runs 3.40 and 3.44.15
scripts/render-xlsx.sh build/przyklad_qgis4.2.3.xlsx   # sheet screenshots (LibreOffice)
```

License: GPL-2.0-or-later ([`LICENSE`](LICENSE)).
