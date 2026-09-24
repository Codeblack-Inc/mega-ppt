# /// script
# dependencies = ["python-pptx>=1.0"]
# ///
"""Smoke test: every layout in sample.json builds and round-trips. uv run tests/test_build.py"""
import json
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent / "skills/mega-ppt"
sys.path.insert(0, str(ROOT / "scripts"))
from build_deck import LAYOUTS, build  # noqa: E402
from pptx import Presentation  # noqa: E402

deck = json.loads((ROOT / "examples/sample.json").read_text(encoding="utf-8"))
used = {s["layout"] for s in deck["slides"]}
assert used == set(LAYOUTS), f"sample.json missing layouts: {set(LAYOUTS) - used}"

with tempfile.TemporaryDirectory() as tmp:
    out = Path(tmp) / "t.pptx"
    assert build(deck, out) == len(deck["slides"])
    prs = Presentation(out)
    assert len(prs.slides) == len(deck["slides"])
    assert prs.slides[3].notes_slide.notes_text_frame.text  # notes survive
print("ok")
