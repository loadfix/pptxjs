"""Generate a minimal 2-slide .pptx fixture for development."""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

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

# Slide with a gradient-filled shape to exercise GradientFill render path.
gradient_slide = prs.slides.add_slide(prs.slide_layouts[5])
gradient_slide.shapes.title.text = "Gradient fill"
grad_shape = gradient_slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(6), Inches(2)
)
# python-pptx doesn't expose gradient fills directly; inject the OOXML.
sp_pr = grad_shape.fill._xPr  # <p:spPr>
# Drop any existing fill (python-pptx sets a default fill).
for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:blipFill", "a:pattFill"):
    for el in sp_pr.findall(qn(tag)):
        sp_pr.remove(el)
grad_xml = (
    "<a:gradFill xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'"
    " flip='none' rotWithShape='1'>"
    "  <a:gsLst>"
    "    <a:gs pos='0'><a:srgbClr val='4F81BD'/></a:gs>"
    "    <a:gs pos='100000'><a:srgbClr val='C0504D'/></a:gs>"
    "  </a:gsLst>"
    "  <a:lin ang='2700000' scaled='1'/>"
    "</a:gradFill>"
)
sp_pr.append(etree.fromstring(grad_xml))
grad_shape.text_frame.text = "Linear gradient"
grad_shape.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

# Slide with a picture-filled shape (BlipFill) to exercise the blip render path.
pic_fill_slide = prs.slides.add_slide(prs.slide_layouts[5])
pic_fill_slide.shapes.title.text = "Picture fill"
pic_shape = pic_fill_slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(5), Inches(3)
)
pic_sp_pr = pic_shape.fill._xPr
for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:blipFill", "a:pattFill"):
    for el in pic_sp_pr.findall(qn(tag)):
        pic_sp_pr.remove(el)
# Add the image as a relationship, then reference it via blipFill.
image_path = Path("/home/ben/code/python-pptx/tests/test_files/python-powered.png")
rId = pic_fill_slide.part.relate_to(
    str(image_path),
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    is_external=False,
) if False else None
# Simpler: add_picture first, grab its rId, then remove the picture shape.
tmp_pic = pic_fill_slide.shapes.add_picture(str(image_path), 0, 0, Inches(1), Inches(1))
blip = tmp_pic._element.find(".//" + qn("a:blip"))
blip_fill_rid = blip.get(qn("r:embed"))
tmp_pic._element.getparent().remove(tmp_pic._element)
blip_xml = (
    "<a:blipFill xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'"
    " xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>"
    f"  <a:blip r:embed='{blip_fill_rid}'/>"
    "  <a:stretch><a:fillRect/></a:stretch>"
    "</a:blipFill>"
)
pic_sp_pr.append(etree.fromstring(blip_xml))
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
# Slide with a chart — exercises the chart graphicFrame fallback path.
# python-pptx doesn't emit a cached preview image, so rendering is expected
# to show the "[Chart]" placeholder; the point of the fixture is to
# confirm the frame is detected rather than silently dropped.
chart_slide = prs.slides.add_slide(prs.slide_layouts[5])
chart_slide.shapes.title.text = "A chart"
chart_data = CategoryChartData()
chart_data.categories = ["A", "B", "C"]
chart_data.add_series("Series 1", (1, 2, 3))
chart_slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(2), Inches(6), Inches(4),
    chart_data,
)

OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print(f"wrote {OUT}")
