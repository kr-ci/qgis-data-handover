#!/usr/bin/env bash
# Zrzuty okna, arkuszy i animacja do README → screenshots/ (QGIS 4.2.3 i LibreOffice w Dockerze).
# Użycie: scripts/readme-media.sh
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
rm -rf build/readme && rm -f build/render/parcels-*.png
scripts/qgis-probe.sh 4.2.3 scripts/readme_media.py | grep -q "tryb=readme"
docker run --rm -v "$ROOT/build:/b" qgis/qgis:4.2.3 chmod -R a+rwX /b/readme
DPI=110 scripts/render-xlsx.sh build/readme/parcels.xlsx >/dev/null
uv run --no-project --python 3.12 --with pillow python3 - <<'PY'
from pathlib import Path

from PIL import Image, ImageChops

src, out = Path("build"), Path("screenshots")
out.mkdir(exist_ok=True)


def sheet(page: int) -> Image.Image:
    """Strona wydruku bez marginesów i stopki (stopka to dolne 8% strony)."""
    img = Image.open(src / "render" / f"parcels-{page}.png").convert("RGB")
    body = img.crop((0, 0, img.width, int(img.height * 0.92)))
    box = ImageChops.difference(body, Image.new("RGB", body.size, "white")).getbbox()
    assert box, page
    pad = 16
    return body.crop((max(box[0] - pad, 0), max(box[1] - pad, 0), box[2] + pad, box[3] + pad))


shots = {
    "window-data": Image.open(src / "readme" / "window_data.png").convert("RGB"),
    "window-look": Image.open(src / "readme" / "window_look.png").convert("RGB"),
    "window-sheets": Image.open(src / "readme" / "window_sheets.png").convert("RGB"),
    "sheet-data": sheet(1),
    "sheet-summary": sheet(2),
    "sheet-dictionary": sheet(3),
    "sheet-about": sheet(4),
    "sheet-changes": sheet(5),
}
for name, img in shots.items():
    if name != "window-sheets":
        img.save(out / f"{name}.png", optimize=True)

# Animacja: trzy zakładki okna, potem arkusze pliku po eksporcie.
size = (1000, 800)
frames = []
for name in ("window-data", "window-look", "window-sheets", "sheet-data", "sheet-summary", "sheet-changes"):
    img = shots[name].copy()
    img.thumbnail(size, Image.Resampling.LANCZOS)
    frame = Image.new("RGB", size, "white")
    frame.paste(img, ((size[0] - img.width) // 2, (size[1] - img.height) // 2))
    frames.append(frame.quantize(colors=128, method=Image.Quantize.MEDIANCUT))
frames[0].save(out / "export.gif", save_all=True, append_images=frames[1:], duration=2500, loop=0, optimize=True)
for p in sorted(out.iterdir()):
    print(f"{p}  {p.stat().st_size // 1024} KB")
PY
