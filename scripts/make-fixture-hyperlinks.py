"""Generate a .pptx fixture exercising hyperlink rendering.

Covers the cases wired up in the P9 renderer:
  1. Slide 1 — text run with an external https:// link.
  2. Slide 1 — text run with a mailto: link.
  3. Slide 2 — shape whose click-action jumps to the next slide (ppaction).
  4. Slide 3 — shape whose click-action jumps directly to slide 1.
  5. Slide 3 — image with an external hyperlink.
  6. Slide 3 — javascript: URL that must be rejected by the scheme whitelist.
"""
from pathlib import Path

from lxml import etree

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


def add_ppaction_next_slide(shape) -> None:
	"""Attach `<a:hlinkClick action="ppaction://hlinkshowjump?jump=nextslide"/>`.

	python-pptx's ActionSetting only exposes HYPERLINK (external URL) and
	NAMED_SLIDE (target_slide=…) cleanly; to force an `hlinkshowjump`
	ppaction URI we drop down to the XML directly and register the
	relationship by hand.
	"""
	part = shape.part
	rId = part.relate_to(
		"ppaction://hlinkshowjump?jump=nextslide",
		RT.HYPERLINK,
		is_external=True,
	)
	cNvPr = shape._element.find(qn("p:nvSpPr") + "/" + qn("p:cNvPr"))
	if cNvPr is None:
		cNvPr = shape._element.find(".//" + qn("p:cNvPr"))
	nsmap = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
			 "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
	hlink = etree.SubElement(cNvPr, qn("a:hlinkClick"))
	hlink.set(qn("r:id"), rId)
	hlink.set("action", "ppaction://hlinkshowjump?jump=nextslide")

OUT = Path(__file__).resolve().parent.parent / "tests" / "render-test" / "hyperlinks" / "presentation.pptx"
IMG = Path("/home/ben/code/python-pptx/tests/test_files/python-powered.png")

prs = Presentation()
BLANK = prs.slide_layouts[5]

# --- Slide 1 — external link + mailto on text runs -------------------------
slide1 = prs.slides.add_slide(BLANK)
slide1.shapes.title.text = "Text hyperlinks"
tb = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(2))
tf = tb.text_frame
para = tf.paragraphs[0]
r = para.add_run()
r.text = "Visit "
r2 = para.add_run()
r2.text = "example.com"
r2.hyperlink.address = "https://example.com"
r2.font.size = Pt(24)
r3 = para.add_run()
r3.text = " or email "
r4 = para.add_run()
r4.text = "hi@example.com"
r4.hyperlink.address = "mailto:hi@example.com"
r4.font.size = Pt(24)

# --- Slide 2 — shape click-action → next slide -----------------------------
slide2 = prs.slides.add_slide(BLANK)
slide2.shapes.title.text = "Shape → next slide"
btn = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3), Inches(3), Inches(4), Inches(1))
btn.text_frame.text = "Go to next slide"
add_ppaction_next_slide(btn)

# --- Slide 3 — shape → specific slide, image → external --------------------
slide3 = prs.slides.add_slide(BLANK)
slide3.shapes.title.text = "Mixed hyperlinks"

back = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(4), Inches(3), Inches(1))
back.text_frame.text = "Back to slide 1"
back.click_action.target_slide = slide1

img = slide3.shapes.add_picture(str(IMG), Inches(5), Inches(2), Inches(3), Inches(1.5))
img.click_action.hyperlink.address = "https://python.org"

# javascript: URL on a text run — must be dropped by the scheme whitelist.
evil = slide3.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(0.6))
evil_run = evil.text_frame.paragraphs[0].add_run()
evil_run.text = "unsafe link"
evil_run.hyperlink.address = "javascript:alert(1)"

OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print(f"wrote {OUT}")
