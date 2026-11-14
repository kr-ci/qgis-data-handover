#!/usr/bin/env bash
# Test integracyjny wtyczki w prawdziwym QGIS, w obrazie Docker qgis/qgis.
# Użycie: scripts/qgis-test.sh 3.44.15   (lub 3.40, 4.2.3)
set -euo pipefail
TAG="${1:?podaj wersję obrazu qgis/qgis, np. 3.44.15}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
docker run --rm -v "$ROOT:/src" -w /src "qgis/qgis:$TAG" bash -c "
  set -e
  status=0
  # Najpierw bez openpyxl: wtyczka musi się wczytać i pokazać instrukcję instalacji.
  xvfb-run -a python3 -u tests/qgis_no_openpyxl.py || status=\$?
  # Paczka bez openpyxl w systemie: eksport musi użyć kopii dołączonej do ZIP (data_handover/ext).
  scripts/package.sh >/dev/null && OPENPYXL=bundled xvfb-run -a python3 -u tests/qgis_zip_install.py || status=\$?
  # Osobne polecenia: przy set -e błąd każdego z nich zatrzymuje skrypt (w „a && b” błąd „a” nie zatrzymuje).
  export DEBIAN_FRONTEND=noninteractive
  apt-get update -qq >/dev/null
  apt-get install -y -qq python3-openpyxl >/dev/null
  xvfb-run -a python3 -u tests/qgis_smoke.py || status=\$?
  # Okno na małym ekranie laptopa (125 %) i na 4K ze skalowaniem 200 %.
  for screen in 1366x768:1.25 3840x2160:2; do
    QT_SCALE_FACTOR=\${screen#*:} xvfb-run -a -s \"-screen 0 \${screen%:*}x24\" python3 -u tests/qgis_window.py || status=\$?
  done
  # Paczka jak do plugins.qgis.org: instalacja z ZIP i uruchomienie.
  scripts/package.sh >/dev/null && OPENPYXL=system xvfb-run -a python3 -u tests/qgis_zip_install.py || status=\$?
  chmod -R a+rwX build 2>/dev/null || true
  exit \$status
"
