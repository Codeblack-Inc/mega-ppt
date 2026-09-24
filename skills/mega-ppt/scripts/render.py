# /// script
# requires-python = ">=3.10"
# dependencies = ["pymupdf"]
# ///
"""pptx → PNG per slide, for visual QA. Needs LibreOffice.
Usage: uv run render.py deck.pptx [outdir] [dpi]  → outdir/slide-01.png ...
If the deck font (Pretendard) is not installed, LibreOffice picks a random fallback (serif or
display faces on macOS), so the deck is re-rendered with a fixed system Korean font instead.
"""
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import pymupdf

SOFFICE = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
# LibreOffice on macOS only resolves Apple SD Gothic Neo by its localized family name
FALLBACK = "Apple SD 산돌고딕 Neo" if sys.platform == "darwin" else "Noto Sans CJK KR"

pptx = Path(sys.argv[1]).resolve()
out = Path(sys.argv[2] if len(sys.argv) > 2 else pptx.with_suffix("")).resolve()
dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 80
out.mkdir(parents=True, exist_ok=True)


def to_pdf(src, tmp):
    try:  # private profile: works while LibreOffice is open, and for parallel renders
        subprocess.run([SOFFICE, f"-env:UserInstallation=file://{tmp}/profile", "--headless",
                        "--convert-to", "pdf", "--outdir", tmp, str(src)],
                       check=True, capture_output=True)
    except (FileNotFoundError, subprocess.CalledProcessError) as e:
        sys.exit(f"LibreOffice 변환 실패 ({e}). 설치: brew install --cask libreoffice")
    return pymupdf.open(Path(tmp) / (src.stem + ".pdf"))


def swap_font(src, dst, old, new):
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it)
            if it.filename.endswith(".xml"):
                data = data.replace(f'typeface="{old}"'.encode(), f'typeface="{new}"'.encode())
            zo.writestr(it, data)


with tempfile.TemporaryDirectory() as tmp:
    doc = to_pdf(pptx, tmp)
    used = {f[3] for page in doc for f in page.get_fonts()}
    if not any("Pretendard" in f for f in used):
        print(f"Pretendard 미설치 → '{FALLBACK}'로 렌더링 (설치: brew install --cask font-pretendard)")
        alt = Path(tmp) / "alt" / pptx.name
        alt.parent.mkdir()
        swap_font(pptx, alt, "Pretendard", FALLBACK)
        doc = to_pdf(alt, alt.parent)
    for i, page in enumerate(doc, 1):
        page.get_pixmap(dpi=dpi).save(out / f"slide-{i:02d}.png")
print(f"{len(doc)} slides → {out}")
