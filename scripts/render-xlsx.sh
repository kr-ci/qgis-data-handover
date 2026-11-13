#!/usr/bin/env bash
# Zrzuty arkuszy pliku XLSX (LibreOffice → PDF → PNG) do wizualnej kontroli.
# Użycie: [DPI=110] scripts/render-xlsx.sh build/plik.xlsx   → build/render/<plik>-N.png
# Uwaga: LibreOffice pokazuje układ wydruku; komentarzy i części stylów tabel Excela nie odda.
set -euo pipefail
FILE="${1:?podaj ścieżkę do pliku .xlsx w katalogu build/}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
IMAGE=data-handover-render:cjk  # nowa nazwa = przebudowa obrazu po zmianie czcionek
if ! docker image inspect "$IMAGE" >/dev/null 2>&1; then
  printf 'FROM debian:trixie-slim\nRUN apt-get update -qq && DEBIAN_FRONTEND=noninteractive apt-get install -y -qq --no-install-recommends libreoffice-calc-nogui poppler-utils fonts-dejavu fonts-crosextra-carlito fonts-noto-cjk >/dev/null && rm -rf /var/lib/apt/lists/*\n' \
    | docker build -q -t "$IMAGE" - >/dev/null
fi
NAME="$(basename "$FILE" .xlsx)"
docker run --rm -v "$ROOT/build:/w" "$IMAGE" bash -c "
  mkdir -p /w/render && cd /tmp
  soffice --headless --convert-to pdf --outdir /tmp '/w/${FILE#build/}' >/dev/null 2>&1
  pdftoppm -png -r ${DPI:-70} '/tmp/$NAME.pdf' '/w/render/$NAME'
  chmod -R a+rwX /w/render
"
ls "$ROOT/build/render/$NAME"*
