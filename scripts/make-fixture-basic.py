"""Generate a minimal 2-slide .pptx fixture for development."""
from pathlib import Path

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

# Preset-geometry showcase — exercises prstGeom rendering.
preset_slide = prs.slides.add_slide(prs.slide_layouts[5])
preset_slide.shapes.title.text = "Preset shapes"
# Lay shapes out on a 3x2 grid of ~2" cells.
presets = [
    (MSO_SHAPE.ROUNDED_RECTANGLE, 0, 0),
    (MSO_SHAPE.DIAMOND, 1, 0),
    (MSO_SHAPE.OVAL, 2, 0),
    (MSO_SHAPE.UP_ARROW, 0, 1),
    (MSO_SHAPE.STAR_5_POINT, 1, 1),
    (MSO_SHAPE.PENTAGON, 2, 1),
]
for shape_enum, col, row in presets:
    left = Inches(0.75 + col * 2.5)
    top = Inches(1.8 + row * 2)
    s = preset_slide.shapes.add_shape(shape_enum, left, top, Inches(2), Inches(1.5))
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor(0x4F, 0x81, 0xBD)
    s.line.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

# Slide exercising rotation + flipH/flipV (Wave 2, P2).
rot_slide = prs.slides.add_slide(prs.slide_layouts[5])
rot_slide.shapes.title.text = "Rotated & flipped"

# Text box rotated 30 degrees.
rtb = rot_slide.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
rtb.rotation = 30
rrun = rtb.text_frame.paragraphs[0].add_run()
rrun.text = "Rotated 30 degrees"
rrun.font.size = Pt(20)

# Right-arrow shape, horizontally flipped.
flipped_arrow = rot_slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5), Inches(2), Inches(2), Inches(1))
flipped_arrow.fill.solid()
flipped_arrow.fill.fore_color.rgb = RGBColor(0x4F, 0x81, 0xBD)
# python-pptx doesn't expose flipH directly; poke the xfrm element.
flipped_arrow.element.spPr.xfrm.set("flipH", "1")

# A 45-degree rotated image.
rot_img = rot_slide.shapes.add_picture(str(image_path), Inches(4), Inches(4), Inches(2), Inches(1))
rot_img.rotation = 45

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
