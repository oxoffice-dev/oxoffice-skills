"""Helpers for the steps oxoffice does not do in one safe call yet. Each one goes away when oxoffice gains
it (oxoffice specs/2026-11-20-skill-helpers-in-the-library.md). Import it from a script run with the Python that has
oxoffice (`oxpy`, this folder on PYTHONPATH), or run it as a command:

  oxpy docx_ops.py comment IN.docx OUT.docx --anchor TEXT --text COMMENT --author NAME [--occurrence N] [--in-tables] [--date ISO]

What writes takes an output path that must not exist yet, as the CLI's outputs, and writes it whole or not at all.
Exit codes: 0 done, 1 refused (nothing written), 2 usage.
"""
import argparse
import io
import sys
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import oxoffice.docx


class EditError(RuntimeError):
    """An edit could not be done exactly as asked: nothing was written, and `doc` was not changed."""


def _new_output(out):
    """Refuse an output that exists already, the way the oxoffice CLI does."""
    if Path(out).exists():
        raise EditError(f"{out} exists: write to a new path")


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _now():
    """The current UTC time as an RFC 3339 string, the form oxoffice takes for comment dates."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _starts(text, anchor):
    """Every start offset of `anchor` in `text`, overlapping ones included."""
    out, k = [], text.find(anchor)
    while k != -1:
        out.append(k)
        k = text.find(anchor, k + 1)
    return out


def _locate(doc, anchor, occurrence=1):
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


def _isolate(doc, flow_index, start, end):
    """Split runs so that characters [start, end) of paragraph `flow_index` form whole runs. Offsets count the
    characters of the paragraph's runs joined, as `_locate` returns them (not `Paragraph.text`, which also holds
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
    # kept for the comment helpers until oxoffice anchors a comment on a text itself; elsewhere use
    # Paragraph.find_text and Document.isolate_span
    flow, start, bi = _locate(doc, anchor, occurrence)
    first, last = _isolate(doc, flow, start, start + len(anchor))
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
    anchors the comment a second time, on `doc`: in a large document, prefer this one. The copy does not know
    the file `doc` was opened from, so its `save` does not refuse that file: save it to a new path."""
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


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
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
    if hasattr(sys.stdout, "reconfigure"):   # Windows consoles and pipes default to a narrower code page
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        if a.cmd == "comment":
            _new_output(a.out)
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
