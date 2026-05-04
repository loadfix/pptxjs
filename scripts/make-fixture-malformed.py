"""Generate a .pptx fixture with one intentionally broken slide.

Regenerates the `basic` fixture into `tests/render-test/malformed/`, then
post-processes the zip to corrupt slide2.xml's XML so it can't be parsed.
Exercises the per-slide try/catch in Presentation.load — the deck should
render N-1 healthy slides plus a red error banner for the broken one.

Also emits a totally-broken zip at `not-a-pptx.bin` for testing the
package-level rejection path (caller supplies a non-pptx buffer).
"""
from pathlib import Path
import runpy
import shutil
import sys
import zipfile

HERE = Path(__file__).resolve().parent
BASIC_OUT = HERE.parent / "tests" / "render-test" / "basic" / "presentation.pptx"
MALFORMED_DIR = HERE.parent / "tests" / "render-test" / "malformed"
OUT = MALFORMED_DIR / "presentation.pptx"
NON_PPTX = MALFORMED_DIR / "not-a-pptx.bin"

# The basic fixture is the source deck we corrupt. Generate it on demand
# only when absent, so repeat runs of this script don't rewrite the basic
# file (python-pptx embeds timestamps, so bytes churn every run).
if not BASIC_OUT.exists():
    runpy.run_path(str(HERE / "make-fixture-basic.py"), run_name="__main__")

MALFORMED_DIR.mkdir(parents=True, exist_ok=True)

# Strategy: the actual parser treats "document with parsererror root" the
# same as "empty XML" — it silently produces an empty slide. That's actually
# healthy behavior, not a test of the try/catch. To force the per-slide
# catch path, we instead drop slide2's file entirely from the archive while
# keeping its entry in presentation.xml.rels. The loader then raises
# `slide part missing` during load, which the per-slide try/catch traps
# into a parseError banner while still rendering the rest of the deck.
TARGET_SLIDE = "ppt/slides/slide2.xml"

TMP = OUT.with_suffix(".tmp.pptx")
with zipfile.ZipFile(BASIC_OUT, "r") as src, zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as dst:
    for item in src.infolist():
        if item.filename == TARGET_SLIDE:
            # Drop this part from the archive entirely.
            continue
        data = src.read(item.filename)
        dst.writestr(item, data)

shutil.move(TMP, OUT)
print(f"wrote {OUT} (slide2.xml removed — its rel dangles)")

# Bogus bytes for the package-level reject path test.
NON_PPTX.write_bytes(b"this is definitely not a zip file, hello world")
print(f"wrote {NON_PPTX}")
