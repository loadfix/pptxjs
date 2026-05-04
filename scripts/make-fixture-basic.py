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

# Slide exercising <a:effectLst>: outer shadow, glow, blur. python-pptx's
# shadow_format API is too limited to emit <a:effectLst> directly, so we
# splice the XML onto each shape's <p:spPr> with lxml.
effects_slide = prs.slides.add_slide(prs.slide_layouts[5])
effects_slide.shapes.title.text = "Effects"

A_URI = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _append_effect_lst(shape, inner_xml):
    """Append <a:effectLst>inner_xml</a:effectLst> to the shape's <p:spPr>."""
    spPr = shape._element.spPr
    # Remove any existing effectLst so re-runs stay idempotent.
    for existing in spPr.findall(qn("a:effectLst")):
        spPr.remove(existing)
    effect_lst = etree.SubElement(spPr, qn("a:effectLst"))
    fragment = etree.fromstring(f"<root xmlns:a='{A_URI}'>{inner_xml}</root>")
    for child in fragment:
        effect_lst.append(child)


shadowed = effects_slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(2), Inches(2.5), Inches(1.2)
)
shadowed.fill.solid()
shadowed.fill.fore_color.rgb = RGBColor(0x4F, 0x81, 0xBD)
shadowed.text_frame.text = "Outer shadow"
shadowed.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
_append_effect_lst(
    shadowed,
    '<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">'
    '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>'
    "</a:outerShdw>",
)

glowing = effects_slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(4), Inches(2), Inches(2.5), Inches(1.2)
)
glowing.fill.solid()
glowing.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
glowing.line.color.rgb = RGBColor(0x00, 0x99, 0x66)
glowing.text_frame.text = "Glow"
_append_effect_lst(
    glowing,
    '<a:glow rad="63500">'
    '<a:srgbClr val="00CC66"><a:alpha val="60000"/></a:srgbClr>'
    "</a:glow>",
)

blurred = effects_slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(7.25), Inches(2), Inches(2.5), Inches(1.2)
)
blurred.fill.solid()
blurred.fill.fore_color.rgb = RGBColor(0xD9, 0x43, 0x6E)
blurred.text_frame.text = "Blur"
blurred.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
_append_effect_lst(blurred, '<a:blur rad="38100" grow="1"/>')

inner_shadowed = effects_slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(4), Inches(2.5), Inches(1.2)
)
inner_shadowed.fill.solid()
inner_shadowed.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)
inner_shadowed.text_frame.text = "Inner shadow"
_append_effect_lst(
    inner_shadowed,
    '<a:innerShdw blurRad="63500" dist="25400" dir="13500000">'
    '<a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr>'
    "</a:innerShdw>",
)

soft_edge = effects_slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(4), Inches(4), Inches(2.5), Inches(1.2)
)
soft_edge.fill.solid()
soft_edge.fill.fore_color.rgb = RGBColor(0xF7, 0x96, 0x46)
soft_edge.text_frame.text = "Soft edge"
_append_effect_lst(soft_edge, '<a:softEdge rad="50800"/>')

reflected = effects_slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(7.25), Inches(4), Inches(2.5), Inches(1.2)
)
reflected.fill.solid()
reflected.fill.fore_color.rgb = RGBColor(0x4F, 0x81, 0xBD)
reflected.text_frame.text = "Reflection"
reflected.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
_append_effect_lst(
    reflected,
    '<a:reflection blurRad="6350" stA="50000" stPos="0" endA="300" endPos="50000"'
    ' dist="38100" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>',
)

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
