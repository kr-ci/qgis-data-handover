#!/usr/bin/env bash
# Pakuje wtyczkę do ZIP gotowego do instalacji („Zainstaluj z pliku ZIP”) i do plugins.qgis.org.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="$(sed -n 's/^version=//p' "$ROOT/data_handover/metadata.txt")"
mkdir -p "$ROOT/build"
OUT="$ROOT/build/data_handover-$VERSION.zip"
rm -f "$OUT"
cd "$ROOT" && python3 -I -c "
import pathlib, sys, zipfile
out = sys.argv[1]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for p in sorted(pathlib.Path('data_handover').rglob('*')):
        if p.is_file() and '__pycache__' not in p.parts:
            z.write(p)
    z.write('LICENSE', 'data_handover/LICENSE')
print(out)
" "$OUT"
