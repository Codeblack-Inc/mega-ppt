# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx>=1.0", "pymupdf"]
# ///
"""Build the gallery site: sample deck × themes → pptx + slide images + data.json.
Usage: uv run gallery/build.py [outdir=site]   (needs LibreOffice)
"""
import copy
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "skills/mega-ppt"
sys.path.insert(0, str(SKILL / "scripts"))
from build_deck import FULL, WARN, build  # noqa: E402

THEMES = {
    "mega": {"name": "Mega", "theme": {}},
    "blue": {"name": "Blue", "theme": {"accent": "#2152E0", "ink": "#14213D",
                                          "text": "#252C3A", "muted": "#5B6475",
                                          "rule": "#D5DAE3", "soft": "#F3F5F9"}},
    "teal": {"name": "Teal · bar", "theme": {"accent": "#0E8A7E", "ink": "#0F2A2E",
                                            "title_style": "bar"}},
    "vermilion": {"name": "Vermilion", "theme": {"accent": "#E4572E", "ink": "#1E1E24",
                                                 "soft": "#F5F3F0", "rule": "#DDD8D2"}},
}

out = ROOT / (sys.argv[1] if len(sys.argv) > 1 else "site")
shutil.rmtree(out, ignore_errors=True)
(out / "decks").mkdir(parents=True)
deck = json.loads((SKILL / "examples/sample.json").read_text(encoding="utf-8"))
extra = json.loads((Path(__file__).parent / "showcase.json").read_text(encoding="utf-8"))
deck["slides"] = deck["slides"][:-1] + extra["slides"] + deck["slides"][-1:]  # closing last
GROUPS = ["기본 구성", "요약·메시지", "근거·데이터", "전략·구조", "실행·일정", "조직·인력", "이미지·화면",
          "효과·요약"]

for key, t in THEMES.items():
    d = copy.deepcopy(deck)
    d["theme"] = {**d.get("theme", {}), **t["theme"]}
    pptx = out / "decks" / f"mega-ppt-sample-{key}.pptx"
    build(d, pptx, SKILL / "examples")
    if WARN:
        sys.exit("build warnings:\n" + "\n".join(WARN))
    subprocess.run(["uv", "run", "-q", str(SKILL / "scripts/render.py"), str(pptx),
                    str(out / "img" / key), "144"], check=True)


def pretty(obj):
    """indent=2, but scalar-only arrays stay on one line."""
    def join(m):
        flat = re.sub(r"\n\s*", " ", m.group(0)).replace("[ ", "[").replace(" ]", "]")
        return flat if len(flat) < 90 else m.group(0)
    return re.sub(r"\[[^\[\]{}]*\]", join, json.dumps(obj, ensure_ascii=False, indent=2))


def panels(s):
    for row in s.get("body", []):
        for c in (row["cols"] if isinstance(row, dict) else row):
            yield c["type"]


slides = []
for i, s in enumerate(deck["slides"], 1):
    g = s.get("gallery", {})
    spec = {k: v for k, v in s.items() if k != "gallery"}
    slides.append({
        "n": i,
        "name": g.get("name", f"Slide {i}"),
        "desc": g.get("desc", ""),
        "group": g.get("group", "기타"),
        "kind": s.get("layout", "content"),
        "panels": sorted(set(panels(s))) if s.get("layout", "content") not in FULL else [],
        "json": pretty(spec),
    })
slides.sort(key=lambda x: (GROUPS.index(x["group"]) if x["group"] in GROUPS else 99, x["n"]))
data = {"themes": [{"key": k, "name": t["name"], "accent": t["theme"].get("accent", "#B5452D")}
                   for k, t in THEMES.items()], "slides": slides}
(out / "data.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
shutil.copy(Path(__file__).parent / "index.html", out / "index.html")
shutil.copy(Path(__file__).parent / "mega-ppt.svg", out / "mega-ppt.svg")
shutil.copy(Path(__file__).parent / "symbol.svg", out / "symbol.svg")
print(f"gallery → {out} ({len(slides)} slides × {len(THEMES)} themes)")
