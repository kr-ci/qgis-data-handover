#!/usr/bin/env bash
# Sonda w prawdziwym QGIS, w obrazie Docker qgis/qgis.
# Użycie: N=200000 scripts/qgis-probe.sh 4.2.3 sondy/<plik>.py [argumenty]
# Wynik to linia „tryb=…”; jej brak albo kod wyjścia ≠ 0 oznacza błąd sondy.
set -euo pipefail
TAG="${1:?podaj wersję obrazu qgis/qgis, np. 4.2.3}"
SCRIPT="${2:?podaj ścieżkę sondy}"
shift 2
ROOT="$(pwd)"
docker run --rm -e "N=${N:-200000}" -v "$ROOT:/src" -w /src "qgis/qgis:$TAG" bash -c "
  set -euo pipefail
  export DEBIAN_FRONTEND=noninteractive
  apt-get update -qq >/dev/null
  apt-get install -y -qq python3-openpyxl >/dev/null
  timeout 1200 xvfb-run -a python3 -u $SCRIPT $*
"
