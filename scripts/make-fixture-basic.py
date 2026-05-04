"""Generate a minimal 2-slide .pptx fixture for development."""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

OUT = Path(__file__).resolve().parent.parent / "tests" / "render-test" / "basic" / "presentation.pptx"

prs = Presentation()

title_slide = prs.slides.add_slide(prs.slide_layouts[0])
title_slide.shapes.title.text = "pptxjs fixture"
title_slide.placeholders[1].text = "A minimal presentation"

content_slide = prs.slides.add_slide(prs.slide_layouts[1])
content_slide.shapes.title.text = "Bullet points"
tf = content_slide.placeholders[1].text_frame
tf.text = "First bullet"
for line in ("Second bullet", "Third bullet"):
    p = tf.add_paragraph()
    p.text = line
    p.font.size = Pt(18)

shapes_slide = prs.slides.add_slide(prs.slide_layouts[5])
shapes_slide.shapes.title.text = "A shape"
tb = shapes_slide.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
run = tb.text_frame.paragraphs[0].add_run()
run.text = "Free-floating textbox"
run.font.size = Pt(24)
run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
run.font.bold = True

image_slide = prs.slides.add_slide(prs.slide_layouts[5])
image_slide.shapes.title.text = "An image"
image_path = Path("/home/ben/code/python-pptx/tests/test_files/python-powered.png")
image_slide.shapes.add_picture(str(image_path), Inches(3), Inches(2), Inches(4), Inches(2))

# Slide with a custom solid background to exercise slide-level <p:bg>.
bg_slide = prs.slides.add_slide(prs.slide_layouts[5])
bg_slide.shapes.title.text = "Tinted background"
bg_slide.background.fill.solid()
bg_slide.background.fill.fore_color.rgb = RGBColor(0xFF, 0xF4, 0xE0)

# Shapes with fills and borders.
fill_slide = prs.slides.add_slide(prs.slide_layouts[5])
fill_slide.shapes.title.text = "Filled shapes"

filled = fill_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(3), Inches(1.5))
filled.fill.solid()
filled.fill.fore_color.rgb = RGBColor(0x4F, 0x81, 0xBD)
filled.line.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
filled.line.width = Emu(38100)  # 3 pt
filled.text_frame.text = "Filled + bordered"
filled.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

bordered = fill_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(2), Inches(3), Inches(1.5))
bordered.fill.background()  # transparent
bordered.line.color.rgb = RGBColor(0xC0, 0x50, 0x4D)
bordered.line.width = Emu(19050)  # 1.5 pt
bordered.text_frame.text = "Border only"

# Slide with a table.
table_slide = prs.slides.add_slide(prs.slide_layouts[5])
table_slide.shapes.title.text = "A table"
rows, cols = 3, 3
left, top, width, height = Inches(1), Inches(2), Inches(8), Inches(3)
tbl = table_slide.shapes.add_table(rows, cols, left, top, width, height).table
headers = ["Name", "Role", "Score"]
for c, h in enumerate(headers):
    tbl.cell(0, c).text = h
data = [("Alice", "Engineer", "92"), ("Bob", "Designer", "88")]
for r, row in enumerate(data, start=1):
    for c, val in enumerate(row):
        tbl.cell(r, c).text = val

# Typography slide — exercises underline, strike, super/sub, letter-spacing,
# and paragraph line-spacing. python-pptx's high-level Font only exposes
# bold/italic/underline/color/size, so we patch the rPr/pPr XML for the rest.
typo_slide = prs.slides.add_slide(prs.slide_layouts[5])
typo_slide.shapes.title.text = "Typography"
typo_tb = typo_slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4))
typo_tf = typo_tb.text_frame
typo_tf.word_wrap = True


def _set_rpr_attr(run, name, value):
    rPr = run._r.get_or_add_rPr()
    rPr.set(name, value)


# Paragraph 1: plain + underlined + strikethrough runs.
p1 = typo_tf.paragraphs[0]
r = p1.add_run(); r.text = "Plain then "
u = p1.add_run(); u.text = "underlined"; u.font.underline = True
p1.add_run().text = " then "
s = p1.add_run(); s.text = "strikethrough"
_set_rpr_attr(s, "strike", "sngStrike")
p1.add_run().text = "."

# Paragraph 2: super/sub and letter-spacing.
p2 = typo_tf.add_paragraph()
p2.add_run().text = "E = mc"
sup = p2.add_run(); sup.text = "2"
_set_rpr_attr(sup, "baseline", "30000")
p2.add_run().text = ", H"
sub = p2.add_run(); sub.text = "2"
_set_rpr_attr(sub, "baseline", "-25000")
p2.add_run().text = "O, "
spc = p2.add_run(); spc.text = "l e t t e r   s p a c e d"
# spc attr is in 1/100 pt — 200 = 2pt.
_set_rpr_attr(spc, "spc", "200")

# Paragraph 3: 1.5x line-spacing on a wrapping block so you can see the gap.
p3 = typo_tf.add_paragraph()
p3.text = (
    "This paragraph uses 1.5x line spacing so you can see the extra "
    "vertical room between wrapped lines. Quick brown foxes and lazy dogs."
)
# Set a:spcPct on pPr — val is per-mille. 150000 = 1.5x.
pPr = p3._pPr if p3._pPr is not None else p3._p.get_or_add_pPr()
lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
etree.SubElement(lnSpc, qn("a:spcPct")).set("val", "150000")
# Re-order: lnSpc must appear before other children per the schema; since this
# is a fresh pPr the order will be correct.

OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print(f"wrote {OUT}")
