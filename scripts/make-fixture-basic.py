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
# Speaker notes on the title slide exercise the notesSlide rel + body
# placeholder parsing in src/notes.ts.
title_slide.notes_slide.notes_text_frame.text = (
    "Speaker notes here.\nSecond line of notes for testing."
)

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

OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)

# python-pptx has no comment API, so patch the saved .pptx zip directly to
# add one comment on slide 1. This exercises commentAuthors.xml parsing
# plus the per-slide comments part + rel wiring in src/comments.ts.
import shutil
import zipfile

AUTHORS_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="0" name="Fixture Author" initials="FA" lastIdx="1" clrIdx="0"/>
</p:cmAuthorLst>
"""

COMMENTS_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2024-01-15T10:00:00Z" idx="1">
    <p:pos x="1828800" y="1143000"/>
    <p:text>Sample review comment for testing.</p:text>
  </p:cm>
</p:cmLst>
"""

# Relationships to add: presentation → commentAuthors; slide1 → comments1.
SLIDE1_RELS_ADD = '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>'
PRES_RELS_ADD = '<Relationship Id="rId98" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>'

CT_OVERRIDES = [
    '<Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>',
    '<Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>',
]

TMP = OUT.with_suffix(".tmp.pptx")
with zipfile.ZipFile(OUT, "r") as src, zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as dst:
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "ppt/slides/_rels/slide1.xml.rels":
            data = data.replace(b"</Relationships>", SLIDE1_RELS_ADD.encode() + b"</Relationships>")
        elif item.filename == "ppt/_rels/presentation.xml.rels":
            data = data.replace(b"</Relationships>", PRES_RELS_ADD.encode() + b"</Relationships>")
        elif item.filename == "[Content_Types].xml":
            data = data.replace(b"</Types>", "".join(CT_OVERRIDES).encode() + b"</Types>")
        dst.writestr(item, data)
    dst.writestr("ppt/commentAuthors.xml", AUTHORS_XML)
    dst.writestr("ppt/comments/comment1.xml", COMMENTS_XML)

shutil.move(TMP, OUT)
print(f"wrote {OUT}")
