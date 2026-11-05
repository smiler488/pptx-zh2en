"""Extract -> translate -> apply on a generated deck, in both directions (no network)."""
import json
import subprocess
import sys
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "translate_pptx_inline.py"


def make_deck(path, title, bullet, cell, note):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title
    slide.placeholders[1].text_frame.text = bullet
    table = slide.shapes.add_table(1, 2, Inches(1), Inches(4), Inches(4), Inches(1)).table
    table.cell(0, 0).text = cell
    table.cell(0, 1).text = "5000"
    slide.notes_slide.notes_text_frame.text = note
    prs.save(path)


def run(direction, src, mapping, tmp):
    deck, js, out = tmp / f"{direction}.pptx", tmp / f"{direction}.json", tmp / f"{direction}_out.pptx"
    make_deck(deck, *src)
    subprocess.run([sys.executable, SCRIPT, "--mode", "extract", "--direction", direction,
                    "--input", deck, "--translations-file", js], check=True, capture_output=True)
    entries = json.loads(js.read_text(encoding="utf-8"))
    for e in entries:
        e["translated"] = mapping[e["original"]]
    js.write_text(json.dumps(entries, ensure_ascii=False), encoding="utf-8")
    res = subprocess.run([sys.executable, SCRIPT, "--mode", "apply", "--direction", direction,
                          "--input", deck, "--translations-file", js, "--output", out],
                         check=True, capture_output=True, text=True)
    prs = Presentation(out)
    texts = [s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame]
    texts += [c.text for s in prs.slides[0].shapes if s.has_table for row in s.table.rows for c in row.cells]
    texts.append(prs.slides[0].notes_slide.notes_text_frame.text)
    return texts, res.stdout


with tempfile.TemporaryDirectory() as d:
    tmp = Path(d)
    zh = ("冠层光合模型", "叶面积指数（LAI）是关键参数", "品种", "这里讲一下背景。")
    en = ("Canopy Photosynthesis Model", "Leaf area index (LAI) is a key parameter", "Cultivar", "Here is the background.")
    texts, log = run("zh2en", zh, dict(zip(zh, en)), tmp)
    assert set(en) <= set(texts), texts
    assert not any("一" <= ch <= "鿿" for t in texts for ch in t), "Chinese left in English deck"
    texts, _ = run("en2zh", en, dict(zip(en, zh)), tmp)
    assert set(zh) <= set(texts), texts
print("pptx-zh2en round trip: OK")
