# /// script
# dependencies = ["python-pptx>=1.0"]
# ///
"""Smoke test: sample.json + gallery/showcase.json use every panel/layout, build
warning-free, round-trip.
uv run tests/test_build.py"""
import json
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent / "skills/mega-ppt"
sys.path.insert(0, str(ROOT / "scripts"))
from build_deck import FULL, PANELS, WARN, build, fit_size  # noqa: E402
from pptx import Presentation  # noqa: E402

deck = json.loads((ROOT / "examples/sample.json").read_text(encoding="utf-8"))
extra = json.loads((ROOT.parent.parent / "gallery/showcase.json").read_text(encoding="utf-8"))
deck["slides"] += extra["slides"]
layouts = {s.get("layout", "content") for s in deck["slides"]}
assert layouts == set(FULL) | {"content"}, f"examples missing layouts: {set(FULL) - layouts}"


def panels(body):
    for row in body:
        for c in (row["cols"] if isinstance(row, dict) else row):
            yield c["type"]


used = {t for s in deck["slides"] for t in panels(s.get("body", []))}
assert used == set(PANELS), f"examples missing panels: {set(PANELS) - used}"

with tempfile.TemporaryDirectory() as tmp:
    out = Path(tmp) / "t.pptx"
    assert build(deck, out, ROOT / "examples") == len(deck["slides"])
    assert not WARN, WARN
    assert len(Presentation(out).slides) == len(deck["slides"])

# fit: long text shrinks and warns, short text keeps size
WARN.clear()
assert fit_size("가" * 400, 12, 2, 0.5) == 8 and WARN
assert fit_size("짧은 문장", 12, 4, 0.5) == 12
print("ok")
