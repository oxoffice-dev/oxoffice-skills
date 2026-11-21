"""Helpers for the steps oxoffice 0.1.0 does not do in one safe call yet. Each one goes away when oxoffice gains
it (oxoffice specs/2026-11-20-skill-helpers-in-the-library.md). Import it from a script run with the Python that has
oxoffice (`$PY`, this folder on PYTHONPATH), or run it as a command:

  $PY docx_ops.py text IN.docx                                                every paragraph of every story, by part
  $PY docx_ops.py count IN.docx TEXT                                          occurrences of TEXT by story kind
  $PY docx_ops.py replace IN.docx OUT.docx --edit OLD NEW COUNT [--edit ...] [--allow-unreached]
  $PY docx_ops.py comment IN.docx OUT.docx --anchor TEXT --text COMMENT --author NAME [--occurrence N] [--in-tables] [--date ISO]

What writes takes an output path and saves with `Document.save`, which refuses the file the document was opened
from. Exit codes: 0 done, 1 refused (nothing written), 2 usage.
"""
import argparse
import collections
import io
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import oxoffice.docx


class EditError(RuntimeError):
    """An edit could not be done exactly as asked: nothing was written, and `doc` was not changed."""


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
STORY_PART = re.compile(r"word/(document|header|footer|footnotes|endnotes)(\d*)\.xml$")
STORY_KIND = {"document": "body", "header": "header", "footer": "footer", "footnotes": "footnote", "endnotes": "endnote"}


def _now():
    """The current UTC time as an RFC 3339 string, the form oxoffice takes for comment dates."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _package(source):
    """The .docx bytes of a Document, bytes, or a path."""
    if hasattr(source, "to_bytes"):
        return source.to_bytes()
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    return Path(source).read_bytes()


def _part_paragraphs(xml):
    """[(style id, text)] for every w:p of one XML part, in document order. A paragraph's text is its own
    w:t and w:tab ("\\t") (a text box's paragraphs are listed on their own, not inside the paragraph that anchors
    them), content controls, tracked insertions, hyperlinks and fields included, deletions (w:delText) out, and
    mc:Fallback copies skipped, so a Word text box counts once."""
    out = []

    def walk(el, acc):
        if el.tag == MC_FALLBACK:
            return
        if el.tag == W_NS + "p":
            slot = len(out)
            out.append(None)
            mine = []
            style = el.find(f"{W_NS}pPr/{W_NS}pStyle")
            for child in el:
                walk(child, mine)
            out[slot] = (style.get(W_NS + "val") if style is not None else None, "".join(mine))
            return
        if el.tag == W_NS + "t":
            acc.append(el.text or "")
        for child in el:
            if child.tag == W_NS + "tab" and el.tag == W_NS + "r":  # a tab character, not a tab stop of w:pPr
                acc.append("\t")
                continue
            walk(child, acc)

    walk(ET.fromstring(xml), [])
    return out


def story_paragraphs(source):
    """[(part name, story kind, style id, text)] for every paragraph of the body (table cells, text boxes,
    content controls included), the headers and footers (each variant part), the footnotes and endnotes, read
    from the package XML: what a reader of the document sees, in the accepted view of tracked changes (a
    paragraph that a tracked deletion removes is listed with empty text)."""
    out = []
    with zipfile.ZipFile(io.BytesIO(_package(source))) as z:
        names = sorted((n for n in z.namelist() if STORY_PART.match(n)),
                       key=lambda n: (n != "word/document.xml", n))
        for name in names:
            kind = STORY_KIND[STORY_PART.match(name).group(1)]
            out += [(name, kind, style, text) for style, text in _part_paragraphs(z.read(name))]
    return out


def all_text(source):
    """[(story kind, paragraph text)] for every paragraph of every story (see story_paragraphs)."""
    return [(kind, text) for _, kind, _, text in story_paragraphs(source)]


def count(source, text):
    """Occurrences of `text` by story kind, as a reader sees them: {"body": n, "footer": n, ...}. A header or
    footer text counts once per variant part (default, first page, even pages), like replacement does."""
    got = collections.Counter()
    for kind, t in all_text(source):
        if text in t:
            got[kind] += t.count(text)
    return dict(got)


def replace_batch(src, out, edits, allow_unreached=False):
    """Apply [(old, new, expected_count), ...] in order, all or nothing. The edits run once, in memory, and are
    checked there before anything is saved: nothing is written, and EditError names every problem, when a
    replacement count differs from its expected count, or when `old` is still in the document after its
    replacement (a match across the edge of a content control, a tracked insertion or a simple field), unless
    allow_unreached=True. Counts are oxoffice's: headers and footers once per variant part, a Word text box once."""
    probe = oxoffice.docx.Document.open(src)
    wrong = []
    for old, new, expected in edits:
        seen = sum(count(probe, old).values())
        got = probe.try_replace_text(old, new)
        if got != expected:
            wrong.append(f"{old!r}: expected {expected}, found {got}")
        if allow_unreached:
            continue
        if old not in new:
            left = [(name, style, t) for name, _, style, t in story_paragraphs(probe) if old in t]
            if left:
                where = "; ".join(f"{name}{f' [{style}]' if style else ''}: {t[:60]!r}" for name, style, t in left[:5])
                toc = any(style and re.match(r"(?i)(toc|tm|inhalt|verzeichnis)\s*\d", style) for _, style, _ in left)
                wrong.append(f"{old!r}: {sum(t.count(old) for _, _, t in left)} occurrence(s) left after the replacement, "
                             f"out of its reach ({where}{'; ...' if len(left) > 5 else ''})"
                             + ("; table of contents entries come back with rebuild_toc" if toc else "")
                             + "; pass allow_unreached=True once handled")
        elif seen > got:
            wrong.append(f"{old!r}: {seen} occurrence(s) seen, {got} replaced; pass allow_unreached=True once handled")
    if wrong:
        raise EditError("; ".join(wrong))
    probe.save(out)
    return Path(out)


def _starts(text, anchor):
    """Every start offset of `anchor` in `text`, overlapping ones included."""
    out, k = [], text.find(anchor)
    while k != -1:
        out.append(k)
        k = text.find(anchor, k + 1)
    return out


def locate(doc, anchor, occurrence=1):
    """(flow index, run offset, body index) of the n-th occurrence of `anchor` within the text of one paragraph
    that is a direct child of the body. Paragraphs in table cells and in block content controls (a table of
    contents, Google Docs blocks) are skipped: comments cannot be anchored there by body index. The offset
    counts the characters of the paragraph's runs (`"".join(r.text for r in p.runs)`), which leave out the
    text inside a simple field, a smart tag or a custom XML element that `Paragraph.text` shows: an anchor
    that sits in such an element, or crosses its edge, is refused."""
    seen = 0
    for i, p in enumerate(doc.paragraphs):
        shown = _starts(p.text, anchor)
        if not shown:
            continue
        try:
            bi = doc.find_content_index(doc.paragraphs[i])
        except ValueError:  # not a direct child of the body
            continue
        if seen + len(shown) < occurrence:
            seen += len(shown)
            continue
        offsets = _starts("".join(r.text for r in doc.paragraphs[i].runs), anchor)
        if len(offsets) != len(shown):
            raise EditError(f"anchor {anchor!r} occurrence {occurrence}: in its paragraph the text sits inside a simple "
                            "field, a smart tag or a custom XML element (or crosses its edge), whose text is not in "
                            "Paragraph.runs: anchor on text outside that element")
        return i, offsets[occurrence - seen - 1], bi
    raise EditError(f"anchor {anchor!r} occurrence {occurrence} not found in the body's own paragraphs "
                    "(table cells and content-control blocks are not searched)")


def isolate(doc, flow_index, start, end):
    """Split runs so that characters [start, end) of paragraph `flow_index` form whole runs. Offsets count the
    characters of the paragraph's runs joined, as `locate` returns them (not `Paragraph.text`, which also holds
    the text of simple fields, smart tags and custom XML). Returns the (first, last_exclusive) run indices
    covering them. The paragraph must be a direct child of the body: split_run takes its body index, as
    RunPosition and find_content_index do."""
    before = doc.paragraphs[flow_index].text
    body_index = doc.find_content_index(doc.paragraphs[flow_index])

    def split_at(offset):
        pos = 0
        for k, r in enumerate(doc.paragraphs[flow_index].runs):
            n = len(r.text)
            if offset == pos:
                return k
            if pos < offset < pos + n:
                doc.split_run(body_index, k, offset - pos)
                return k + 1
            pos += n
        return len(doc.paragraphs[flow_index].runs)

    split_at(end)
    first = split_at(start)
    last = split_at(end)
    if doc.paragraphs[flow_index].text != before:
        raise EditError("splitting the runs changed the paragraph text")
    return first, last


def _comment(doc, anchor, text, author, initials, occurrence, date):
    flow, start, bi = locate(doc, anchor, occurrence)
    first, last = isolate(doc, flow, start, start + len(anchor))
    rng = oxoffice.docx.RunRange(start=oxoffice.docx.RunPosition(body_index=bi, run_index=first),
                         end=oxoffice.docx.RunPosition(body_index=bi, run_index=last))
    return doc.add_comment(rng, author=author, text=text, initials=initials, date=date)


def comment_landings(data):
    """{comment id: True if its range starts inside a table cell} for every comment range of document.xml."""
    out = {}

    def walk(el, in_cell):
        if el.tag == W_NS + "commentRangeStart":
            out[int(el.get(W_NS + "id"))] = in_cell
        for child in el:
            walk(child, in_cell or el.tag == W_NS + "tc")

    with zipfile.ZipFile(io.BytesIO(data) if isinstance(data, bytes) else data) as z:
        walk(ET.fromstring(z.read("word/document.xml")), False)
    return out


def _cell_occurrence(doc, anchor, occurrence):
    """The zero-based occurrence number that `add_comment_on_text` gives the n-th occurrence of `anchor`
    inside a table cell. oxoffice numbers the anchor's occurrences over the whole main story in document order;
    anchoring every one of them on a copy (the text does not change) shows which ones sit in a cell."""
    probe = oxoffice.docx.Document.from_bytes(doc.to_bytes())
    ids = []
    while True:
        try:
            ids.append(probe.add_comment_on_text(anchor, author="probe", text="probe", occurrence=len(ids)))
        except oxoffice.docx.DocxError:
            break
    landing = comment_landings(probe.to_bytes())
    in_cells = [k for k, cid in enumerate(ids) if landing.get(cid)]
    if len(in_cells) < occurrence:
        raise EditError(f"anchor {anchor!r} occurrence {occurrence} not found in table cells "
                        f"({len(in_cells)} there, {len(ids)} in the main story)")
    return in_cells[occurrence - 1]


def _comment_checked(doc, anchor, text, author, initials, occurrence, date, in_tables, k=None):
    """Anchor the comment on `doc` and read the anchored text back; EditError if oxoffice refuses the range or
    anchors it elsewhere, `doc` being then left half edited: callers pass a copy they drop on error. `k`: the
    occurrence number _cell_occurrence gives, when known."""
    if in_tables and k is None:
        k = _cell_occurrence(doc, anchor, occurrence)
    why = "nothing was written; report it as a new gap"
    try:
        if in_tables:
            cid = doc.add_comment_on_text(anchor, author=author, text=text, occurrence=k, initials=initials, date=date)
        else:
            cid = _comment(doc, anchor, text, author, initials, occurrence, date)
    except oxoffice.docx.DocxError as e:
        raise EditError(f"oxoffice refused the range ({e}): {why}") from e
    data = doc.to_bytes()
    got = next((c.anchor_text or None for c in oxoffice.docx.Document.from_bytes(data).comments if c.id == cid), None)
    if got != anchor:
        raise EditError(f"the comment would be anchored on {got!r} instead of {anchor!r}: {why}")
    if in_tables and not comment_landings(data).get(cid):
        raise EditError(f"the comment would be anchored outside the table cells: {why}")
    return cid


def commented(doc, anchor, text, author, initials=None, occurrence=1, date=None, in_tables=False):
    """(a commented copy of `doc`, the comment id): the comment of comment_on_text (same arguments and checks),
    anchored once, on a copy that is checked and returned; `doc` itself is never changed. comment_on_text
    anchors the comment a second time, on `doc`: in a large document, prefer this one."""
    copy = oxoffice.docx.Document.from_bytes(doc.to_bytes())
    return copy, _comment_checked(copy, anchor, text, author, initials, occurrence, date or _now(), in_tables)


def comment_on_text(doc, anchor, text, author, initials=None, occurrence=1, date=None, in_tables=False):
    """Anchor a comment on exactly `anchor`, dated `date` (RFC 3339; default: now, UTC). `occurrence` counts
    from 1 in the body's own paragraphs, or, with `in_tables=True`, in the paragraphs of table cells (nested
    tables included), in document order. The whole operation is first run on a copy and the anchored text
    read back: if oxoffice refuses the range or anchors it elsewhere, EditError is raised and `doc` is left
    unchanged. Returns the comment id. `commented` does the same work once and returns the copy."""
    date = date or _now()
    k = _cell_occurrence(doc, anchor, occurrence) if in_tables else None
    _comment_checked(oxoffice.docx.Document.from_bytes(doc.to_bytes()), anchor, text, author, initials, occurrence, date,
                     in_tables, k)
    if in_tables:
        return doc.add_comment_on_text(anchor, author=author, text=text, occurrence=k, initials=initials, date=date)
    return _comment(doc, anchor, text, author, initials, occurrence, date)


_BOOKMARK_START = re.compile(r"<w:bookmarkStart\b[^>]*/>")


def strip_generated_bookmarks(src, out, prefix="_Toc"):
    """Copy the .docx `src` (a path, bytes or a Document) to `out` without the body bookmarks whose name starts
    with `prefix` (start and end markers), and return how many it removed. A table of contents rebuilt from
    scratch names its heading bookmarks `_Toc1`, `_Toc2`, ... in document order, so between two versions where a
    heading was added or removed a name moves to another heading, and `compare` refuses the pair
    (`oxoffice limitations docx`, docx-compare-toc-bookmarks): run this on both copies first. The TOC's links then point at no bookmark in
    the redline; rebuild the TOC on the accepted result. Refuses `out == src`."""
    if not isinstance(src, (bytes, bytearray)) and not hasattr(src, "to_bytes") and \
            Path(src).resolve() == Path(out).resolve():
        raise EditError("refusing to write over the input; write to a new path")
    removed = []

    def start(match):
        element = match.group(0)
        name = re.search(r'\bw:name="([^"]*)"', element)
        ident = re.search(r'\bw:id="([^"]*)"', element)
        if name and ident and name.group(1).startswith(prefix):
            removed.append(ident.group(1))
            return ""
        return element

    with zipfile.ZipFile(io.BytesIO(_package(src))) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml = _BOOKMARK_START.sub(start, data.decode("utf-8"))
                ids = set(removed)
                xml = re.sub(r"<w:bookmarkEnd\b[^>]*/>",
                             lambda m: "" if (re.search(r'\bw:id="([^"]*)"', m.group(0)) or [None, None])[1] in ids
                             else m.group(0), xml)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    return len(removed)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("text")
    p.add_argument("src")
    p = sub.add_parser("count")
    p.add_argument("src")
    p.add_argument("text")
    p = sub.add_parser("replace")
    p.add_argument("src")
    p.add_argument("out")
    p.add_argument("--edit", nargs=3, action="append", metavar=("OLD", "NEW", "COUNT"), required=True)
    p.add_argument("--allow-unreached", action="store_true")
    p = sub.add_parser("comment")
    p.add_argument("src")
    p.add_argument("out")
    p.add_argument("--anchor", required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--author", required=True)
    p.add_argument("--initials")
    p.add_argument("--occurrence", type=int, default=1)
    p.add_argument("--in-tables", action="store_true", help="count and anchor in table cells")
    p.add_argument("--date", help="RFC 3339, default now (UTC)")
    a = ap.parse_args()
    try:
        if a.cmd == "text":
            for name, kind, style, t in story_paragraphs(a.src):
                print(f"{kind}\t{name}\t{style or ''}\t{t}")
        elif a.cmd == "count":
            print(json.dumps(count(oxoffice.docx.Document.open(a.src), a.text)))
        elif a.cmd == "replace":
            replace_batch(a.src, a.out, [(o, n, int(c)) for o, n, c in a.edit], a.allow_unreached)
            print(f"wrote {a.out}")
        elif a.cmd == "comment":
            doc = oxoffice.docx.Document.open(a.src)
            cid = _comment_checked(doc, a.anchor, a.text, a.author, a.initials, a.occurrence, a.date or _now(),
                                   a.in_tables)
            doc.save(a.out)
            print(f"comment {cid} written to {a.out}")
    except (EditError, oxoffice.docx.DocxError) as e:
        print(f"refused: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
