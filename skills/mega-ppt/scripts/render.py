# /// script
# requires-python = ">=3.10"
# dependencies = ["pymupdf"]
# ///
"""pptx → PNG per slide, for visual QA. Needs LibreOffice.
Usage: uv run render.py deck.pptx [outdir] [dpi]  → outdir/slide-01.png ...
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf

SOFFICE = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"

pptx = Path(sys.argv[1]).resolve()
out = Path(sys.argv[2] if len(sys.argv) > 2 else pptx.with_suffix("")).resolve()
dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 80
out.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory() as tmp:
    try:
        subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(pptx)],
                       check=True, capture_output=True)
    except (FileNotFoundError, subprocess.CalledProcessError) as e:
        sys.exit(f"LibreOffice 변환 실패 ({e}). 설치: brew install --cask libreoffice")
    doc = pymupdf.open(Path(tmp) / (pptx.stem + ".pdf"))
    for i, page in enumerate(doc, 1):
        page.get_pixmap(dpi=dpi).save(out / f"slide-{i:02d}.png")
print(f"{len(doc)} slides → {out}")
