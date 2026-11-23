"""Check what the skill examples wrote, after run_examples.py: the files each example promises, with the
content it describes. Reads them with oxoffice, python-docx and python-pptx. A document's paragraphs are read from
`oxoffice docx text --json` (one call per file, cached): on a build with debug assertions, walking `Document.paragraphs`
costs tens of milliseconds per paragraph, minutes for the whole check.

    python tests/check_outputs.py WORK_DIR
"""

import functools
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import docx
import pptx

import oxoffice.docx
import oxoffice.pptx

WORK = Path(sys.argv[1])
d = WORK / "docx"


@functools.cache
def paragraphs(name):
    """The paragraphs of `oxoffice docx text --json`, in document order: the body's and, marked by their `path`,
    the table cells', each with `text`, `style`, `runs` (text and formatting), `path` and `body_index`. The body's
    are those of Document.paragraphs in a document without tracked paragraph-mark deletions or tables inside
    content controls, as every output here."""
    res = subprocess.run(["oxoffice", "docx", "text", "--json", str(d / name)], capture_output=True,
                         text=True, encoding="utf-8")
    assert res.returncode == 0, f"oxoffice docx text --json {name}: {res.stderr.strip()}"
    return json.loads(res.stdout)["paragraphs"]


in_cell = lambda q: any(step["kind"] == "cell" for step in q["path"])
body = lambda name: [q for q in paragraphs(name) if not in_cell(q)]
texts_of = lambda name: [q["text"] for q in body(name)]


def cell_text(name, first, row, col):
    """The text of a cell of the table whose first cell reads `first`."""
    cells = [q for q in paragraphs(name) if in_cell(q)]
    at = lambda q: tuple(s["index"] for s in q["path"] if s["kind"] in ("row", "cell"))
    table = next(q["body_index"] for q in cells if at(q) == (0, 0) and q["text"] == first)
    return "\n".join(q["text"] for q in cells if q["body_index"] == table and at(q) == (row, col))


texts = texts_of("edited.docx")
assert any("Issue C, for approval" in t for t in texts) and any("two points lower" in t for t in texts)
doc = oxoffice.docx.Document.open(d / "commented.docx")
first = next(c for c in doc.comments if c.parent_id is None)
assert first.resolved and first.date
assert any(c.text == "Rename this column?" and c.date for c in doc.comments)
assert first.anchor_text == "about 12 mm"
reviewed = oxoffice.docx.Document.open(d / "reviewed.docx")
root = next(c for c in reviewed.comments if c.id == first.id)
assert (root.anchor_text, root.resolved, len(reviewed.comments)) == ("bearing positions", True, len(doc.comments))
assert [c.parent_id for c in reviewed.comments if c.text == "The 2019 survey marks."] == [root.id]
assert not any(s.startswith("the bearing at pier 2") for s in texts_of("reviewed.docx"))
assert oxoffice.docx.Document.open(d / "moved-cli.docx").comments[0].anchor_text == "bearing positions"
method = [s for s in texts_of("scoped.docx") if "Drainage performance was checked" in s]
assert method[0].startswith("Drainage performance was checked by pouring water at the high points of each span")
assert all("high points of the deck" in t for t in method[1:]) and len(method) > 1
actions = [cell_text("scoped.docx", "Ref", r, 5) for r in (1, 5)]
assert actions == ["Monitor", "No action"]
issue = texts_of("issue-c.docx")
assert any("Issue C" in s for s in issue) and not any("Issue B" in s for s in issue)
report = texts_of("report.docx")
for name in ("checked.docx", "accepted.docx"):
    lines = texts_of(name)
    assert any("two points lower" in s for s in lines) and not any("three points lower" in s for s in lines)
assert len(oxoffice.docx.Document.open(d / "accepted.docx").revisions) == 0
added = texts_of("added.docx")
k = next(k for k, s in enumerate(added) if s.startswith("None of these defects"))
assert added[k + 1] == "The next principal inspection is due in 2031." and len(added) == len(report) + 1
moved = texts_of("restructured.docx")
lead = report.index("The main findings are:")
assert moved[moved.index("The main findings are:") - 1] == report[lead + 1]   # moved one item down
assert len(moved) == len(report) - 1
def toc_links(name):   # a rebuilt TOC links each entry to a _Toc bookmark and gets its page from a PAGEREF
    with zipfile.ZipFile(d / name) as z:
        xml = z.read("word/document.xml").decode()
    return xml.count('w:anchor="_Toc'), xml.count("PAGEREF _Toc"), xml.count('w:name="_Toc')


assert toc_links("report.docx") == (0, 0, 0)   # the input's TOC is plain cached text
for name in ("toc.docx", "fields.docx"):
    entries = [q["text"] for q in body(name) if re.fullmatch(r"TOC\d", q["style"] or "")]
    assert len(entries) > 3 and all(s.split("\t")[-1].isdigit() for s in entries)
    assert toc_links(name) == (len(entries), len(entries), len(entries)), (name, toc_links(name), len(entries))
page = lambda name: next(s.split("\t")[-1] for s in texts_of(name) if s.startswith("Access arrangements\t"))
assert int(page("fields.docx")) == int(page("toc.docx")) + 1   # update_layout_backed_fields followed the page break
i = next(k for k, s in enumerate(texts_of("styled.docx")) if s.startswith("The overall condition index"))
styled = oxoffice.docx.Document.open(d / "styled.docx").paragraphs[i]
assert styled.text.startswith("The overall condition index")   # the JSON's index is the handle's
assert (styled.style, styled.paragraph_format.space_after, styled.paragraph_format.keep_with_next) == ("Caption", oxoffice.docx.Pt(12), True)
assert [(c.anchor_text, c.text) for c in oxoffice.docx.Document.open(d / "commented-cli.docx").comments] == [("about 12 mm", "Which reference?")]
assert [(c.anchor_text, c.text) for c in oxoffice.docx.Document.open(d / "commented-cell.docx").comments] == [("spalled concrete", "Which face?")]
assert all(c.date for name in ("commented-cli.docx", "commented-cell.docx") for c in oxoffice.docx.Document.open(d / name).comments)
record = json.loads((d / "report.json").read_text())
assert len(record["paragraphs"]) > 100 and any(s["kind"] == "footer" for s in record["stories"]), record["stories"][:1]
assert sorted(f.name for f in (d / "checked-pages").iterdir()) == [f"checked_page{k}.png" for k in (1, 2, 3)]
assert all(f.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n" for f in (d / "checked-pages").iterdir())
assert json.loads((d / "layout.json").read_text())["page_count"] > 1
assert any(line.startswith("footer\t") for line in (d / "all-stories.txt").read_text().splitlines())
assert (d / "redline-tracked.pdf").read_bytes()[:5] == b"%PDF-"
from_python = oxoffice.docx.Document.open(d / "redline-py.docx")
assert [(r.kind, r.author, r.story.kind) for r in from_python.revisions] == [
    ("deletion", "Reviewer", "body"), ("insertion", "Reviewer", "body")]
red = oxoffice.docx.Document.open(d / "redline.docx")
assert [(r.kind, r.story.kind) for r in red.revisions] == [("deletion", "body"), ("insertion", "body")]
assert any(r["text"] == "safety of path users" and r["formatting"]["bold"] for q in body("bold.docx") for r in q["runs"])
assert not any(s.startswith("Access equipment, traffic management") for s in moved)
assert any(q["text"] == "Replace the handrail caps" for q in paragraphs("table.docx") if in_cell(q))
with zipfile.ZipFile(d / "from-template.docx") as z:
    assert b"wordprocessingml.document.main+xml" in z.read("[Content_Types].xml")
assert texts_of("note.docx") == ["Site visit note", "Visited on 27 September 2026."]
check = docx.Document(d / "checklist.docx")
assert check.core_properties.title == "Inspection checklist" and check.core_properties.author == "Claude"
assert [(p.style.name, p.text) for p in check.paragraphs] == [
    ("Title", "Inspection checklist"), ("Note box", "Bring the 2019 survey marks and a tape."),
    ("List Paragraph", "Check the bearings."), ("List Paragraph", "Photograph the deck joints.")]
assert check.styles["Note box"].font.italic and check.styles["Note box"].base_style.name == "Normal"
num = [p._p.pPr.numPr for p in check.paragraphs[2:]]
assert all(n is not None and n.numId.val == num[0].numId.val for n in num)
toc1 = [q["text"] for q in body("with-toc.docx") if q["style"] == "TOC1"]
assert [s.split("\t")[0] for s in toc1] == ["Scope", "Findings", "Recommendations"]
assert all(s.split("\t")[1].isdigit() for s in toc1)
with zipfile.ZipFile(d / "with-toc.docx") as z:
    assert b'TOC \\o "1-2" \\h' in z.read("word/document.xml")
assert body("visit.docx")[0]["style"] == "Heading1" and cell_text("visit.docx", "Item", 1, 0) == "Pier 2"
form = oxoffice.docx.Document.open(d / "filled-form.docx")
assert [(c.tag, c.text) for c in form.content_controls() if c.tag == "status"] == [("status", "Final")]
assert [(f.name, f.value) for f in form.legacy_form_fields()] == [("Inspector", "Inspector A")]
assert (form.document_protection.mode, form.document_protection.enforcement) == ("forms", True)
findings = texts_of("findings.docx")
assert findings[:2] == ["Findings", "Inserted after the heading."]
with zipfile.ZipFile(d / "chart.docx") as z:
    chart = re.findall(r"<c:v>([^<]*)</c:v>", z.read("word/charts/chart1.xml").decode())
assert chart == ["Inspections", "2024", "2025", "2026", "12", "15", "9"]
restarted = oxoffice.docx.Document.open(d / "restarted.docx").paragraphs
assert [q.text for q in restarted] == ["Scope", "Findings", "Annex A", "Annex B"]
assert restarted[0].numbering is None and restarted[2].numbering == restarted[3].numbering is not None
with zipfile.ZipFile(d / "restarted.docx") as z:
    assert b"w:startOverride" in z.read("word/numbering.xml")
for name in ("embedded.docx", "embedded-py.docx"):
    font = next(f for f in oxoffice.docx.Document.open(d / name).fonts if f.name == "Aileron")
    assert [(e.kind, e.license.identity) for e in font.embedded_fonts] == [("regular", "CC0-1.0: Aileron")]
p = WORK / "pptx"
base = len(pptx.Presentation(p / "deck.pptx").slides)
assert [len(pptx.Presentation(p / n).slides) for n in ("merged.pptx", "appended.pptx", "merged-py.pptx")] == [
    base + 2, 2 * base, base + 2]
merged = oxoffice.pptx.Presentation(p / "merged-py.pptx")
assert [s.shapes.title.text for s in list(merged.slides)[base:]] == ["Appendix A", "Appendix B"]
chart = next(sh for sh in pptx.Presentation(p / "chart.pptx").slides[base].shapes if sh.has_chart).chart
assert list(chart.plots[0].categories) == ["Option A", "Option B", "Option C"]
assert list(chart.plots[0].series[0].values) == [236.0, 230.0, 251.0]
assert "\n# " in (p / "deck.md").read_text() and "<section" in (p / "deck.html").read_text()
for name in ("embedded.pptx", "embedded-py.pptx"):
    fonts = oxoffice.pptx.Presentation(p / name).embedded_fonts
    assert [(t.typeface, [f.format for f in t.faces]) for t in fonts] == [("Aileron", ["eot"])]
    assert oxoffice.pptx.Presentation(p / name).embed_true_type_fonts
frames = lambda name: [sh.text_frame for s in pptx.Presentation(p / name).slides for sh in s.shapes if sh.has_text_frame]
assert any("9 April 2025" in f.text for f in frames("dated.pptx")) and not any("2 April 2025" in f.text for f in frames("dated.pptx"))
works = next(f for f in frames("run-edited.pptx") if f.text.startswith("Works:")).paragraphs[0].runs[0]
assert works.text == "Works: EUR 188,000" and works.font.size is not None
with zipfile.ZipFile(p / "c3.pptx") as z:
    thread = b"".join(z.read(n) for n in z.namelist() if n.startswith("ppt/comments/"))
assert b"Source for the costs?" in thread and b"Framework rates, 2025." in thread and b'status="resolved"' in thread
assert (p / "deck.pdf").read_bytes()[:5] == b"%PDF-"
assert (p / "thumb.png").read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"
slides = json.loads((p / "deck.json").read_text())["slides"]
assert len(slides) == 7 and any("Three options" in json.dumps(s) for s in slides)
assert [f.name for f in (p / "slide-5").iterdir()] == ["deck_slide5.png"]
texts = "\n".join(sh.text_frame.text for s in pptx.Presentation(p / "edited.pptx").slides for sh in s.shapes if sh.has_text_frame)
assert "EUR 236,000" in texts and "eight weeks" in texts
shapes = pptx.Presentation(p / "shapes.pptx").slides[3].shapes
badge = next(sh for sh in shapes if sh.has_text_frame and sh.text_frame.text == "Decision")
assert str(badge.fill.fore_color.rgb) == "7B1E3A"
styled = pptx.Presentation(p / "styled.pptx").slides[4].shapes
card, arrow = styled[len(styled) - 2], styled[len(styled) - 1]
assert card.auto_shape_type == pptx.enum.shapes.MSO_SHAPE.RECTANGLE and str(card.fill.fore_color.rgb) == "F2F2F2"
assert card.shadow.inherit is False and b'<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl"' in card._element.xml.encode()
assert arrow.line.dash_style == pptx.enum.dml.MSO_LINE_DASH_STYLE.DASH
assert b'<a:tailEnd type="triangle" w="lg" len="lg"/>' in arrow._element.xml.encode()
assert b'<a:effectRef idx="0">' in arrow._element.xml.encode()
imported, base = pptx.Presentation(p / "imported.pptx").slides, pptx.Presentation(p / "deck.pptx").slides
assert len(imported) == len(base) + 1 and imported[1].shapes.title.text == "Appendix: survey method"
assert imported[1].slide_layout.name == "Title Only"
def texts_of(shapes):
    for sh in shapes:
        yield from texts_of(sh.shapes) if sh.shape_type == 6 else [sh.text_frame.text if sh.has_text_frame else ""]
scoped = [" ".join(texts_of(s.shapes)) for s in pptx.Presentation(p / "scoped.pptx").slides]
assert "12 mm" in scoped[1] and "14 mm" in scoped[2] and "12 mm" not in scoped[2]
assert "Total: EUR 236,000" in scoped[5] and "EUR 230,000" in scoped[6]
re_deck = pptx.Presentation(p / "reordered.pptx")
assert re_deck.slides[2].notes_slide.notes_text_frame.text == "Mention the photo record for each defect."
assert re_deck.slides[-1].shapes.title.text == "Questions"
with zipfile.ZipFile(p / "from-template.pptx") as z:
    assert b"presentationml.presentation.main+xml" in z.read("[Content_Types].xml")
before, grown = pptx.Presentation(p / "deck.pptx").slides[3].shapes, pptx.Presentation(p / "grown.pptx").slides[3].shapes
table, old = next(sh.table for sh in grown if sh.has_table), next(sh.table for sh in before if sh.has_table)
assert len(table.rows) == len(old.rows) + 1 and table.cell(len(old.rows) - 1, 0).text == "Contingency"
assert table.cell(len(table.rows) - 1, 0).text == old.cell(len(old.rows) - 1, 0).text
assert table.rows[len(old.rows) - 1].height in (old.rows[len(old.rows) - 2].height, old.rows[len(old.rows) - 1].height)
group = grown[len(before)]
assert group.shape_type == 6 and [sh.has_text_frame and sh.text_frame.text for sh in group.shapes][1] == "Legend"
links = pptx.Presentation(p / "links.pptx")
option_b = next(sh for sh in links.slides[3].shapes if sh.has_text_frame and sh.text_frame.text.startswith("Option B"))
assert option_b.click_action.target_slide.slide_id == links.slides[4].slide_id
assert links.slides[0].shapes[3].click_action.hyperlink.address == "https://example.com/footbridge"
with zipfile.ZipFile(p / "links.pptx") as z:
    comment = b"".join(z.read(n) for n in z.namelist() if n.startswith("ppt/comments/"))
assert b"Which option did the board pick?" in comment and f'<ac:spMk id="{option_b.shape_id}"/>'.encode() in comment
one_shot = pptx.Presentation(p / "s5.pptx")
assert len(one_shot.slides) == 8 and one_shot.slides[7].notes_slide.notes_text_frame.text == "Questions from the board."
assert one_shot.core_properties.title == "Footbridge options" and one_shot.core_properties.author == "Claude"
footers = [sh.text_frame.text for s in list(one_shot.slides)[1:] for sh in s.placeholders
           if sh.placeholder_format.type == pptx.enum.shapes.PP_PLACEHOLDER.FOOTER]
assert footers == ["Riverton Footbridge"] * 7
assert json.loads((p / "fit.json").read_text())["overflow"] is True
branded = oxoffice.pptx.Presentation(p / "branded.pptx")
assert str(branded.slide_master.theme.colors["accent1"]) == "7B1E3A" and branded.slide_master.theme.fonts.major.latin == "Georgia"
assert branded.slide_master.shapes[len(branded.slide_master.shapes) - 1].alt_text == "Company logo"
assert {s.transition.type for s in branded.slides} == {"fade"} and branded.core_properties.title == "Footbridge options"
assert [s.header_footer.slide_number for s in branded.slides][:2] == [False, True]
steps = oxoffice.pptx.Presentation(p / "steps.pptx").slides[7].shapes
assert [para.auto_number for para in steps[1].text_frame.paragraphs] == ["arabicPeriod"] * 3
assert steps[1].text_frame.paragraphs[0].runs[0].font.size <= oxoffice.pptx.Pt(24) and steps[2].alt_text == "Survey"
assert b'<a:stCxn id="' in steps[4].xml and b"<a:endCxn" in steps[4].xml
raw = oxoffice.pptx.Presentation(p / "raw.pptx").slides[3].shapes
assert any('numCol="2"' in sh.text_frame.element.xml for sh in raw if sh.has_text_frame)
photo = (d / "photo.png").read_bytes()
assert any(pic.bytes == photo for pic in oxoffice.docx.Document.open(d / "picture-replaced.docx").pictures())
assert [pic.bytes == photo for pic in oxoffice.docx.Document.open(d / "picture-added.docx").pictures()] == [True]
replaced = next(sh for sh in pptx.Presentation(p / "picture.pptx").slides[0].shapes if sh.shape_type == 13)
assert replaced.image.blob == (p / "photo.png").read_bytes()
assert json.loads((p / "structure.json").read_text())["slide_details"][3]["shape_details"]
print("every example output as described")
