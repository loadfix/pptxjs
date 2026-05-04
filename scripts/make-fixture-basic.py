"""Generate a minimal 2-slide .pptx fixture for development."""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
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

# Slide exercising image crop / grayscale / brightness+contrast / alpha.
# python-pptx has no high-level API for <a:srcRect>/<a:lum>/<a:alphaModFix>
# etc., so we patch the <a:blipFill>/<a:blip> XML directly after each
# picture is added.
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

def _a(name):
    return f"{{{A_NS}}}{name}"

def _sub(parent, tag, **attrs):
    el = etree.SubElement(parent, _a(tag))
    for k, v in attrs.items():
        el.set(k, str(v))
    return el

def _blipfill(pic):
    return pic.find(f"{{{P_NS}}}blipFill")

def _blip(pic):
    return _blipfill(pic).find(_a("blip"))

img_adjust_slide = prs.slides.add_slide(prs.slide_layouts[5])
img_adjust_slide.shapes.title.text = "Image adjustments"
src_img = "/home/ben/code/python-pptx/tests/test_files/python-powered.png"

# 1) Cropped: trim 20% left, 10% top, 20% right, 10% bottom.
cropped = img_adjust_slide.shapes.add_picture(src_img, Inches(0.5), Inches(1.8), Inches(2.5), Inches(1.8))
bf = _blipfill(cropped._element)
# Insert <a:srcRect> as first child of <a:blipFill>, before <a:blip>.
src_rect = etree.Element(_a("srcRect"), l="20000", t="10000", r="20000", b="10000")
bf.insert(0, src_rect)

# 2) Grayscale version of the same image.
gray = img_adjust_slide.shapes.add_picture(src_img, Inches(3.3), Inches(1.8), Inches(2.5), Inches(1.8))
_sub(_blip(gray._element), "grayscl")

# 3) High-contrast + brightness boost (approx +20% bright, +30% contrast).
contrasted = img_adjust_slide.shapes.add_picture(src_img, Inches(6.1), Inches(1.8), Inches(2.5), Inches(1.8))
_sub(_blip(contrasted._element), "lum", bright="20000", contrast="30000")

# 4) Alpha-reduced image (50% opacity).
alpha = img_adjust_slide.shapes.add_picture(src_img, Inches(0.5), Inches(4.2), Inches(2.5), Inches(1.8))
_sub(_blip(alpha._element), "alphaModFix", amt="50000")

# 5) biLevel (threshold) — collapses toward black/white.
bilevel = img_adjust_slide.shapes.add_picture(src_img, Inches(3.3), Inches(4.2), Inches(2.5), Inches(1.8))
_sub(_blip(bilevel._element), "biLevel", thresh="50000")

# 6) Duotone — map luminance to two colors (fallback to blend mode in renderer).
duo = img_adjust_slide.shapes.add_picture(src_img, Inches(6.1), Inches(4.2), Inches(2.5), Inches(1.8))
duotone_el = etree.SubElement(_blip(duo._element), _a("duotone"))
_sub(duotone_el, "srgbClr", val="1F3A93")
_sub(duotone_el, "srgbClr", val="FFD9B3")

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

OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print(f"wrote {OUT}")
