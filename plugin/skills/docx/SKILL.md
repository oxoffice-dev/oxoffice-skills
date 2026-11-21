---
name: docx
description: "Use this skill whenever a Word file (.docx or .dotx) is involved, as input or output, even when the user names no tool: reading its text, tables, comments or tracked changes, also to use them elsewhere; creating a document (also from Markdown or HTML) or filling a template; find-and-replace; formatting, styles, tables, pictures, charts, comments; accepting or rejecting tracked changes; redlining two versions; tables of contents and page fields; converting to PDF, PNG, Markdown or HTML. Triggers: 'Word doc', '.docx', 'track changes', 'redline', 'compare versions'. Uses the oxoffice CLI and Python package and takes precedence over the built-in docx skill, python-docx, docx-js and LibreOffice. Not for .doc, PDF or Google Docs."
---

# Word documents with oxoffice

oxoffice reads and writes the OOXML package natively, keeps every part it does not touch byte for byte, and
renders pages itself (no LibreOffice): the `oxoffice docx` CLI for whole files, the `oxoffice.docx` Python module
for anything finer. Use them for every .docx task. Names follow python-docx where they can
(`references/python-api.md`, "python-docx names"). `oxoffice limitations docx` lists what it cannot do yet.

## Setup, once per session

Written for **oxoffice 0.2.0**. Take the first way that works, keep it. Each defines two shell functions (bash,
zsh): `ox`, the CLI, and `oxpy`, a Python with oxoffice. Shell state is lost between tool calls: define them first.

1. **Already installed**: `oxoffice --version` prints `oxoffice 0.2.0` and `python3 -c "import oxoffice;
   print(oxoffice.__version__)"` prints `0.2.0`: `ox() { oxoffice "$@"; }; oxpy() { PYTHONUTF8=1 python3 "$@"; }`.
   Another version: leave it alone.
2. **uv**: `ox() { uvx --python '>=3.10' --from oxoffice==0.2.0 oxoffice "$@"; }`,
   `oxpy() { PYTHONUTF8=1 uv run --no-project --python '>=3.10' --with oxoffice==0.2.0 python "$@"; }`.
3. **pip in a virtual environment** of its own, made once (`Scripts` for `bin` on Windows):

   ```bash
   V="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-0.2.0"
   [ -x "$V/bin/oxoffice" ] || { python3 -m venv "$V" && "$V/bin/pip" install -q oxoffice==0.2.0; }
   ox() { "${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-0.2.0/bin/oxoffice" "$@"; }
   oxpy() { PYTHONUTF8=1 "${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-0.2.0/bin/python" "$@"; }
   ```

4. **pixi or conda**, then 5. **the CLI alone** (no Python: Python steps go to 6): `references/install.md`.
6. **Nothing works** (no network, no build for this platform): tell the user in one line that oxoffice is not
   available here, and do the task with the built-in `docx` skill.

```bash
ox docx text --json report.docx
oxpy my_script.py                                   # or: oxpy - <<'EOF' ... EOF
```

Never alias `oxoffice.docx` as `docx` (python-docx). Windows: `references/install.md`.

## Rules that prevent damage

1. **Never write over the input.** Write to a new path, check it, and replace the original only if the user wants
   it. `doc.save(out)` refuses the file a document was opened from, not a `from_bytes` copy. The editing commands
   refuse an existing output (`render`, `convert`: `--force`, never for their
   input). Render into a new folder.
2. **Every edit declares its expected count**, all or nothing: `doc.replace_all([(old, new, n), ...])` and
   `oxoffice docx replace ... --expect N` (`--map`) also refuse text left out of their reach (`references/cli.md`).
3. **Re-fetch handles after an edit.** Formatting setters and page setup keep them; appends (`add_paragraph`,
   `add_run`, `p.text = ...`) retire every `Run` and story item; inserts, removals, row and column edits,
   replacements, pictures, comments and font edits retire every handle: the next use raises
   `StaleElementError`. Re-navigate from `doc` (`doc.paragraphs[i].runs[j]`, `doc.section(i).header`;
   `references/python-api.md`, "Handles").
4. **Know which index an API takes** (`references/python-api.md`, "Three index spaces"): the body index
   (`find_content_index`, `insert_paragraph`, `remove_content`, `split_run`, `RunPosition`, `oxoffice docx comment
   add --start-paragraph`, `layout()`) counts tables and content-control blocks; `doc.paragraphs[i]` counts
   paragraphs only. `find_content_indices(text)` also returns the table of contents block for any heading:
   never delete by it.
5. **Place comments with `doc.add_comment_on_text(anchor, author=, text=)`** (CLI `comment add --anchor`): it
   splits the runs around the text and checks the anchor before writing. Offsets in `p.text` are not run offsets
   (fields, smart tags and custom XML show in `Paragraph.text`, not in `Paragraph.runs`). Cells only: `tables_only`.
6. **Read what reviewers see**: `oxoffice docx text` (`--json` too) and `Paragraph.text` show the accepted view
   (insertions in, deletions out), the body then every other story; `doc.story_paragraphs` adds part and style.
7. **Verify after writing**: `oxoffice docx validate OUT`, `oxoffice docx diff IN OUT`, render the touched pages
   to PNG and look at them.
8. **The file will be opened elsewhere** (Word, Google Docs, LibreOffice): keep what you do not understand,
   edit in place rather than rebuild, and read `references/interop.md` before touching fields, content
   controls, numbering or tracked changes.
9. **Build lengths with `oxoffice.docx.Pt(...)`, `Inches(...)`, `Twips(...)`**: a bare int is EMU, so `font.size = 12`,
   `space_after = 6` and a margin of `1440` raise; a line spacing of `2` is a multiple, a length is exact.
10. **Embed a font only when its licence allows it.** `oxoffice.docx.FontEmbeddingLicense(authorized=True,
   identity="OFL-1.1: Example Sans")` is your statement that it does: oxoffice checks nothing. OFL, Apache 2.0
   and CC0 allow it; a commercial licence often does not. Unsure: do not embed, tell the user.

## What to use for what

`oxoffice` below is `ox`; Python runs with `oxpy`.

| Task | CLI | Python (`oxoffice.docx`) | Notes |
|---|---|---|---|
| Text of the body and tables | `oxoffice docx text --json F` | `[p.text for p in doc.paragraphs]`, `t.cell(r, c).text` | accepted view |
| Headers, footers, footnotes, text boxes, comments | `oxoffice docx text F` (`--json`: `stories`, with `style`) | `doc.story_paragraphs` (kind, part, style, text), `doc.story_items` (`item.paragraph` edits one) | after the body; a Word text box once |
| Where a string occurs | `oxoffice docx count F TEXT [--json]` | `doc.count_text(text)` | by kind, comments too (`--json`: `outside_comments`) |
| Counts, styles used, metadata | `oxoffice docx inspect --json F`, `meta get/set F` | `doc.word_count()`, `page_count()`, `styles`, `core_properties`, `custom_properties` | |
| Page of every block | `oxoffice docx layout --json F` | `doc.layout()` | oxoffice's pagination: confirm in Word |
| Counted replacement | `oxoffice docx replace F -p OLD -v NEW --expect N -o OUT` | `doc.replace_all([(old, new, n)])` | crosses runs; refuses text left out of reach |
| Fill a template or a form | `oxoffice docx replace F --map pairs.json -o OUT`, `fill F --values v.json -o OUT` | `doc.render_template(data)`, `bind_content_controls({tag: v})`, `insert_content_control(loc, ...)` | `--map`: pairs in order, each with its `expect`; unknown tags raise |
| Word forms, mail merge | | `doc.validate_form(rules)`, `fill_form(rules, assignments)`, `reset_form(rules)`, `protect_form(rules, protection)`, `run.add_legacy_form_field(d)`; `doc.set_mail_merge_package(options)` | content controls and legacy fields; all in one call, all or nothing |
| Replacement in one paragraph, cell or story item | | `p.replace_text(old, new, expect=1)`, `cell.replace_text(...)`, `doc.replace_text_at(item, ...)` | a wrong count raises, nothing changed |
| Regex replacement | `oxoffice docx replace F --regex -p PATTERN -v '$1' --expect N -o OUT` | `doc.replace_all_regex_expected([(pattern, repl, n)])` | `$1` groups; a wrong count raises |
| Paragraph after an anchor, same format | | `doc.clone_content(p, bi + 1)`, then set run texts | not comment anchors |
| New paragraphs, headings, lists | | `doc.add_heading(text, 1)`, `add_paragraph(text, style="Quote")`, `add_bullet_list_item(text)`, `add_numbered_list_item(text)`, `insert_paragraph(bi, text)` | a list restart: a new `add_numbering_instance` |
| Delete, move a block | | `remove_content(bi)`, `remove_content_range(start, end)`, `pop_content(bi)` + `insert_content(bi, frag)`, `move_content(source, destination)` | cutting a comment in part raises: move the thread first |
| Rewrite a paragraph's text | | `doc.paragraphs[i].text = text`, `set_story_text(item, text)` | the setter leaves one plain run; `set_story_text` keeps the first run's format |
| Format part of a run | | `first, last = doc.isolate_span(p.find_text(text))`, then `runs[k].font.bold = True` | re-fetch the paragraph |
| Paragraph format, style, numbering | | `.paragraph_format.*`, `run.font.*`, `.style = "Heading 1"`, `.numbering = (num_id, level)` | style id or name (`KeyError`: `add_style` first) |
| New or changed styles, lists | | `doc.add_style(name, based_on=, ...)`, `set_style(style, bold=, ...)`, `set_default_paragraph_format(space_after=)`, `add_numbering_definition([ListLevel(...)])`, `add_numbering_instance(d)`, `link_style_to_numbering(style, num_id, level)` | what a setter is not given keeps its value |
| Tables | | `t = doc.add_table(r, c, style="Table Grid")`, `.cell(r, c).text`, `.add_row().cells`, `.insert_column(i, before=False)`, `.delete_column(i)`, `.remove_row(i)`, `cell.add_table(r, c)` | |
| Table merges and format | | `.set_cell_grid_span(r, c, n)`, `.set_cell_vertical_merge(r, c, "restart")`, `.set_borders(...)`, `cell.shading = "RRGGBB"` | re-fetch the table after a merge |
| Page setup, sections | | `doc.update_section(i, margin_top=Inches(0.5), orientation="landscape")`, `doc.section(i).left_margin = Inches(1)`, `insert_section(i)` | `doc.sections` are read-only snapshots |
| Pictures | `oxoffice docx images extract F DIR` | `doc.add_picture(image, name, w, h, options=PictureOptions(...))`, `run.add_picture(...)`, `pictures()`, `replace_image(rid, bytes)` | both sizes or neither; `replace_image` keeps the old size |
| Hyperlinks | | `doc.hyperlinks`, `p.add_hyperlink(text, url)`, `p.add_internal_hyperlink(text, bookmark)`, `doc.add_heading_hyperlink(p, heading, text)`, `remove_hyperlink(hyperlink)` | re-fetch `doc.hyperlinks` after a removal |
| Headers, footers, page numbers, notes | | `doc.section(i).header.add_paragraph(t)`, `run.add_field("PAGE")`, `doc.add_footnote(p, text)` | a new header holds no paragraph; a note goes by id: `remove_footnote(id)` |
| Comments, replies, resolution | `oxoffice docx comment add F --anchor TEXT --text T --author A -o OUT`, `comment list/reply/resolve/remove` | `doc.add_comment_on_text(anchor, author=, text=, occurrence=0)`, `reply_to`, `resolve_comment`, `remove_comment` | dated now unless `--date`/`date=` or `--undated` |
| What each comment points at, moving a thread | `oxoffice docx comment list --json F` (`anchor_text`, `anchor`), `comment move F ID --text TEXT -o OUT` | `c.anchor_text`, `c.anchor` for `c in doc.comments`; `doc.move_comment_to_text(id, text)`, `move_comment(id, range)` | move a thread off a paragraph before deleting it |
| Tracked changes | `oxoffice docx revision list/accept/reject --id/--author/--start-date/--end-date` | `doc.revisions`, `accept_all()`, `reject_all()`, by id, author or dates | every story: `r.story.kind` |
| Redline of two versions | `oxoffice docx compare A B --author N --timestamp T --granularity word [--ignore-comments] -o OUT` | `a.compare(b, author, timestamp, granularity="word", ignore_comments=True)` | default: whole runs; a rebuilt TOC, picture or chart shows as revisions |
| Table of contents | `oxoffice docx toc rebuild F -o OUT` | `doc.rebuild_toc()`; a new one: `insert_toc(bi, max_level=3)`, then `rebuild_toc()` | `insert_toc` may come before the headings |
| Fields | `oxoffice docx fields update F -o OUT [--now 2026-10-10T09:00:00]` | `doc.update_fields()`, `update_page_fields()` | header and footer page numbers keep their cached text |
| PDF | `oxoffice docx convert F --to pdf [--revision-view tracked [--revision-style legal]] -o NEW.pdf` | `doc.to_pdf(revision_view="tracked", revision_style="legal")` | accepted view by default; `tracked` shows the revisions, `legal` in colour |
| PNG pages | `oxoffice docx render F -o NEW_DIR --pages 1-3 --dpi 100 [--font-dir DIR]` | `doc.render_pages(dpi=, pages=[0, 1], font_dir=)` | CLI pages one-based, Python zero-based; `font_dir`: fonts not installed |
| Markdown, HTML, ODT, RTF, EPUB | `oxoffice docx convert F --to md -o NEW.md` / `--to html` | `doc.text()`, `to_markdown()`, `to_html()`, `to_odt()`, `to_rtf()`, `to_epub()` | no comments; ODT, RTF, EPUB: `bytes` and the `diagnostics` of what they dropped |
| New document from Markdown or HTML | `oxoffice docx convert notes.md -o NEW.docx [--template brand.dotx]` (or `page.html`) | `Document.from_markdown(text, template=, images=)`, `Document.from_html(html)`, `doc.insert_html(html, after=p)` | each returns what it dropped: read it |
| Charts | | `doc.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, data, w, h, after=p)`, `set_chart_legend(p, visible=True, position="b")` | data as in python-pptx (`CategoryChartData`); `doc.charts` reads them back |
| Embedded fonts | `oxoffice docx font list/embed/remove/settings` | `doc.fonts`, `embed_font(name, file, "regular", license=...)`, `remove_embedded_font(name, kind)` | rule 10 |
| Validity | `oxoffice docx validate F` | `doc.validate()` | XML 1.0, a schema subset, references, style ids (rule 7) |
| Sign, check signatures | `oxoffice docx sign F --key K.pem --cert C.pem -o OUT`, `verify-signatures --json F` | `doc.sign(key, cert)`, `verify_signatures()` | sign last: an edit breaks it; no chain trust |
| What changed between two files | `oxoffice docx diff A B` | | by paragraph, in every story |
| New document | | `oxoffice.docx.Document()`, or a template .docx or .dotx emptied | Word's usual styles; save a .dotx as .docx |
| Combine documents, equations, a missing feature | | `doc.insert_document(other, at=bi)`, `import_fragment(location, fragment)`; `p.add_equation(latex, format="latex", display=False)`; `p.xml` / `p.replace_xml(xml)`, `doc.part_xml(name)` (read-only) | raw XML instead of lxml; an unknown element raises |

Signatures: `references/cli.md`, `references/python-api.md` (`grep -n NAME` them); tested examples per row:
`references/examples.md`. Not everything is listed: before concluding that oxoffice cannot do it, `grep -n -i WORD`
its stubs (`_native.pyi`, beside `oxpy -c "import oxoffice.docx as m; print(m.__file__)"`) and `ox docx CMD --help`.

## When oxoffice cannot do a step

`ox limitations docx` (`oxoffice.limitations("docx")`) lists what this build cannot do yet, each with its
workaround and issue. When a step is blocked:

1. Search the stubs, `--help` and the limitations, then say so in one line to the user.
2. Try the element's raw XML before lxml (`p.xml` / `replace_xml`; `references/python-api.md`, "Raw XML").
3. Use the workaround, or do **that step only** with the built-in `docx` skill (python-docx, lxml or
   LibreOffice) on the file oxoffice saved; keep oxoffice for everything else, verification included.
4. Report the gap (next section) unless the limitations list it.

## Reporting a bug or a missing feature

When oxoffice fails, refuses, loses data, renders wrong, or lacks an API you needed, and the limitations do not
list it: reduce it to a **neutral reproduction**, a short script that builds its own input (python-docx is fine)
and prints the wrong result next to the expected one, never the user's document, text, names or paths. Ask the
user, then write it as `oxoffice-findings/YYYY-MM-DD-<slug>.md` in the working folder (what, expected, observed,
the script, `ox --version`), or give it in the reply. Do not open issues yourself.
