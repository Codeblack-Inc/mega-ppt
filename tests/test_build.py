# /// script
# dependencies = ["python-pptx>=1.0", "pymupdf"]
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
from build_deck import FULL, PANELS, WARN, _fmt, _nice_min, build, fit_size, share  # noqa: E402
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
# grow: short text in a roomy box grows to max_size; siblings share the smallest fitting size
assert fit_size("짧은 문장", 11, 4, 2, max_size=13) == 13
long_, short = ["가" * 60] * 4, ["가" * 10]
alone = fit_size(long_, 11, 3, 3, bullets=True, max_size=13, quiet=True)
assert share([(long_, 3, 3), (short, 3, 3)], 11, 13, bullets=True) == alone < 13
# chart helpers: axis floor below the data, number formats
assert _nice_min([-5, 3, 8]) < -5 and _nice_min([5, 5, 5]) < 5 and _nice_min([52, 71]) == 40
assert (_fmt(0.1, "0.0%"), _fmt(5, '0"억"'), _fmt(1234.4, "#,##0")) == ("10.0%", "5억", "1,234")
print("ok")
