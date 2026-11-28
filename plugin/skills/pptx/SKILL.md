---
name: pptx
description: "Use this skill any time a PowerPoint file (.pptx or .potx) is involved, as input or output, even when the user names no tool: reading slide text, outline, notes or comments, also to use them elsewhere; creating a deck or filling a template; editing text without losing formatting; find-and-replace; shapes, pictures, tables, charts; adding, reordering, hiding, importing or removing slides; merging decks; comments; checking that text fits; rendering to PDF or PNG; exporting to Markdown or HTML; embedding fonts. Triggers: 'PowerPoint', 'deck', 'slides', '.pptx'. Uses the oxoffice CLI and Python package and takes precedence over the built-in pptx skill, python-pptx, pptxgenjs and LibreOffice. Not for .ppt, Keynote or Google Slides."
---

# PowerPoint decks with oxoffice

oxoffice reads and writes the presentation package natively, keeps every part it does not touch byte for byte,
lays text out with its own line breaker and renders slides itself (no LibreOffice): the `oxoffice pptx` CLI for
whole decks, the `oxoffice.pptx` Python module for anything finer (main names as in python-pptx, not all). Use them
for every .pptx task. `oxoffice limitations pptx` lists what it cannot do yet; report any new gap (last section).

## Setup, once per session

Needs **oxoffice** 0.10.0 or later (written and checked with 0.10.0): the one installed, else through a package
manager. Take the first way that works, keep it. Each defines two shell functions (bash, zsh): `ox`, the CLI, and
`oxpy`, a Python with oxoffice. Shell state is lost between tool calls: define them first.

1. **Already installed**: `oxoffice --version` and `python3 -c "import oxoffice; print(oxoffice.__version__)"`
   both print 0.10.0 or later: `ox() { oxoffice "$@"; }; oxpy() { PYTHONUTF8=1 python3 "$@"; }`. Older: leave it.
2. **uv**: `ox() { uvx --python '>=3.10' --from 'oxoffice>=0.10.0' oxoffice "$@"; }`,
   `oxpy() { PYTHONUTF8=1 uv run --no-project --python '>=3.10' --with 'oxoffice>=0.10.0' python "$@"; }`.
3. **pip in a virtual environment** of its own, made once (`Scripts` for `bin` on Windows):

   ```bash
   V="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-venv"
   [ -x "$V/bin/oxoffice" ] || { python3 -m venv "$V" && "$V/bin/pip" install -q 'oxoffice>=0.10.0'; }
   ox() { "${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-venv/bin/oxoffice" "$@"; }
   oxpy() { PYTHONUTF8=1 "${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-venv/bin/python" "$@"; }
   ```

4. **pixi or conda**: `references/install.md`.
5. **None works**: tell the user in one line that oxoffice is missing and how to install it (official ways:
   https://github.com/oxoffice-dev/oxoffice#-install), then do the task with the built-in `pptx` skill.

```bash
ox pptx outline deck.pptx
oxpy my_script.py                                   # or: oxpy - <<'EOF' ... EOF
```

Never alias `oxoffice.pptx` as `pptx` (python-pptx). Windows: `references/install.md`.

## Rules that prevent damage

1. **Never write over the input.** `prs.save(out)` refuses the file the deck was opened from (not a `from_bytes`
   copy); the editing commands, `render`, `convert` and `thumbnail` refuse an existing output (`convert`,
   `render`, `thumbnail` and `footer` accept `--force`, never for their own input). Render into a new folder.
2. **Every text replacement declares its expected count**, all or nothing: `oxoffice pptx replace ... --expect N`
   or `--map pairs.json` (slides and notes), `prs.try_replace_text(old, new, expect=n)`,
   `prs.replace_all([(old, new, n), ...])`.
3. **Re-fetch the handles an edit retires**: a handle follows its element through moves, inserts and edits
   elsewhere; `text_frame.text =` / `shape.text =`, `clear()` or a replacement retire the paragraphs and runs
   of that frame, `element.replace_xml` what lies below that element, adding or removing table rows or columns
   the table's cells, removing a shape or slide its own handle; appends keep earlier handles, but not a `runs`
   or `paragraphs` collection taken before them. A retired handle raises `StaleElementError` naming the call:
   write `prs.slides[i].shapes[j]` again (`references/python-api.md`, "Handles").
4. **Walk groups**: `slide.shapes` lists top-level shapes; a group's children are in `shape.shapes`
   (`shape_type == MSO_SHAPE_TYPE.GROUP`, value 6). `slide.shapes.walk()` gives `(depth, shape)` for all of them.
5. **Check the fit after any text change**: `oxoffice pptx fit F` (exit 1 when a frame overflows, with the font scale
   it needs, or cannot be assessed: `cannot-assess`; table cells, layouts and notes too, the whole deck always) or `prs.text_layout()` (the slides' text frames) reports every
   frame whose text overflows, with oxoffice's own line breaks; then render the slide and look at it.
6. **The deck will be opened elsewhere** (PowerPoint, Google Slides, Keynote, LibreOffice): edit in place,
   keep placeholders and layouts, never rebuild a deck to change it.
7. **Embed a font only when its licence allows embedding.** `oxoffice.pptx.FontEmbeddingLicense(authorized=True,
   identity="OFL-1.1: Example Sans")` is your statement that it does: oxoffice checks nothing, not even `fsType`.
   OFL, Apache 2.0 and CC0 allow it; a commercial licence often does not (`fsType` bit 0x0002: Restricted
   License). When it does not, or you cannot tell, do not embed and tell the user.
8. **Lengths and spacings take `Pt(...)` or `Inches(...)`.** A bare int is EMU: `font.size = 12`,
   `space_before = 6` and `font.spacing = 2` raise `ValueError`; write `Pt(6)`, or a float such as 1.5 for a
   multiple of the line. `prs.save("x.pdf")` raises `ValueError`: a PDF is `prs.to_pdf()`.

## What to use for what

| Task | CLI | Python | Notes |
|---|---|---|---|
| Slide text | `oxoffice pptx text [--notes] F` | `shape.text`, `shape.text_frame.paragraphs[k].runs` | |
| Text with structure and formatting | `oxoffice pptx text --json F` | | paragraphs with `path`, `shape_id`, `level`, runs with their formatting; notes |
| Titles and outline | `oxoffice pptx outline [--json] [--notes] F` | `slide.shapes.title` | |
| Structure, shapes, metadata | `oxoffice pptx shapes F [--slide N] [--json]`, `inspect --json F` | `slide.shapes.walk()` (layouts and masters too) | every shape, groups' members included, with id, name, geometry, text |
| Counted replacement | `oxoffice pptx replace F -p OLD -v NEW --expect N -o OUT`, many pairs: `oxoffice pptx replace F --map pairs.json -o OUT` | `prs.try_replace_text(old, new, expect=n)`, `prs.replace_all([(old, new, n), ...])` | keeps run formatting |
| Replacement in one slide or one frame | | `slide.try_replace_text(old, new, expect=n)`, `shape.text_frame.try_replace_text(old, new, expect=n)` | the slide's notes too unless `notes=False` |
| Edit a run, paragraph, text frame | | `run.text`, `run.font.*`, `paragraph.alignment/level/space_*/line_spacing`, `text_frame.margin_*/word_wrap/auto_size/vertical_anchor` | `shape.text = ...` drops run formatting; a bare int is EMU: write `Pt(12)` |
| Bullets, numbering | | `paragraph.bullet = "\u2022"`, `bullet_color`, `auto_number = "arabicPeriod"`, `auto_number_start` | |
| Move, resize, rotate | | `shape.left/top/width/height/rotation` | an inherited placeholder geometry reads None: `shape.effective_geometry()` |
| Fill, line | | `shape.fill.solid()`, `.fill.fore_color.rgb = RGBColor(...)`, `.fill.transparency = 0.3`, `.line.width`, `.line.color.rgb`; `fill.gradient()`, `fill.picture(file)` | |
| Dashes, arrowheads | | `.line.dash_style = MSO_LINE_DASH_STYLE.DASH`, `.line.tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE`, `.width`, `.length`; `head_end` likewise | enums in `oxoffice.pptx.enum.dml` |
| Shadow | | `shape.shadow.visible = True`, `.color.rgb`, `.blur_radius`, `.distance` | `shadow.inherit = False` removes the theme's shadow |
| Connector without the theme effect | | `connector.theme_effect_index = 0` | `add_connector` takes the theme's effect 1, a shadow in the default theme |
| Change a shape's preset | | `shape.auto_shape_type = MSO_SHAPE.RECTANGLE` | autoshapes only |
| Add shapes | | `shapes.add_textbox`, `add_shape(MSO_SHAPE.X, ...)`, `add_connector`, `add_picture`, `add_table`, `add_group_shape()` | a group grows with each `add_*` on its `shapes`; `add_shape` writes python-pptx's theme style (`shape.remove_style()` drops it), `add_textbox` none |
| Z-order, arrange | | `shapes.move(from_, to)`, `shapes.group([...])`, `shape.ungroup()`, `shapes.align([...], "top")`, `shapes.distribute([...], "horizontal")`, `shape.flip_h` | index 0 is the back; distribute takes three shapes or more |
| Alt text | | `shape.alt_text = "..."`, `shape.decorative = True` | every kind of shape |
| Glued connector | | `connector.begin_connect(shape, site)`, `end_connect(shape, site)` | sites of a rectangle: 0 top, 1 left, 2 bottom, 3 right |
| Picture into a placeholder | | `slide.placeholders[idx].insert_picture(file)` | cropped to fill; `shape.is_placeholder`, `placeholder_format.type` |
| Pictures | | `shape.replace_image(file)`, `shape.image.blob` | keeps position, size and crop |
| Tables | | `shape.table.cell(r, c).text`, `.merge(other)`, `table.rows.add_row(i)`, `rows.remove(row)`, `table.columns.add_column(i)`, `columns.remove(col)`, `table.first_row`, `style_id` | re-fetch the cells and rows after adding or removing a row or column; an empty cell (from `add_row()` too) has no run: set `cell.text` |
| Slides | `oxoffice pptx slide add/duplicate/remove/move/hide/show F ... -o OUT` | `slides.add_slide(layout)`, `slides.add_slide_at(layout, i)`, `slides.duplicate(slide)`, `slides.move(i, j)`, `slides.remove(slide)`, `slide.hidden`, `slide.transition.type = "fade"` | |
| Slide numbers, footer, date | `oxoffice pptx footer F --slide-number --footer T --skip-title -o OUT` | `prs.set_header_footer(footer=T)`, `slide.header_footer` | as PowerPoint's Header and Footer dialog; `add_field("slidenum")` in a text box |
| Theme, master, layouts | | `prs.slide_master.theme.colors["accent1"] = ...`, `prs.apply_theme(path)`, `prs.slide_master.shapes.add_picture(logo, ...)`, `next(lay for lay in prs.slide_layouts if lay.name == n)` | a logo on the master shows on every slide |
| Import slides, merge decks | `oxoffice pptx slide import DEST SOURCE [--slides 1,3-5] [--at N] [--keep-source-formatting] -o OUT` | `copies, diagnostics = prs.slides.import_slides(other, [0, 2], index=1, keep_source_formatting=True)`, `prs.slides.import_slide(other.slides[k], layout=prs.slide_layouts[j], index=None)` | charts, objects, comments and notes travel; without `layout=` or `keep_source_formatting`, a layout of the same name must exist here; a dropped slide jump is in the diagnostics |
| Charts | | `slide.shapes.add_chart(oxoffice.pptx.XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data)`, `shape.has_chart`, `shape.chart.categories` / `.series`, `shape.chart.replace_data(chart_data)`, `chart.has_legend`, `chart.legend_position = "b"`, `chart.legend_overlay` | chart data as in python-pptx (`oxoffice.pptx.CategoryChartData`); `replace_data` keeps the series formatting |
| Speaker notes | `oxoffice pptx text --notes F`, `oxoffice pptx notes set F N --text T -o OUT` | `slide.notes_text` (get and set), `slide.notes_slide.notes_text_frame.text` as in python-pptx | None when the slide has no notes |
| Hyperlinks, slide jumps | | `run.hyperlink.address`; `shape.click_action.hyperlink.address = url`, `shape.click_action.target_slide = prs.slides[k]` | any shape, group members included; set None to clear |
| Metadata | `oxoffice pptx meta get F`, `oxoffice pptx meta set F --title T --author A -o OUT` | `prs.core_properties.title = ...`, `prs.app_properties.company`, `prs.custom_properties.set(name, value)` | dates are `datetime` (a str raises); a new deck is stamped 2000-01-01 |
| Comments | `oxoffice pptx comment list/add/reply/resolve/remove` | `prs.add_comment_author`, `slide.add_comment`, `reply_to_comment`, `resolve_comment`, `remove_comment`, `move_comment` | `shape_id=` anchors on a shape; add the author first |
| Text fit | `oxoffice pptx fit [--json] F` | `prs.text_layout(width_factor=1.0)`, `prs.fit_report()`, `text_frame.fit_text(max_size=18)`, `text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` then `prs.refresh_autofit()` | oxoffice's line breaks; `auto_size` alone writes no font scale and no new height: `refresh_autofit()` before the save |
| PDF | `oxoffice pptx convert F --to pdf -o NEW.pdf [--font-dir DIR]` | `prs.to_pdf(font_dir=)`, `prs.to_notes_pdf()`, `to_pdfa()`, `to_handout_pdf(6)` | `font_dir` for fonts that are not installed; what no font of the deck draws (CJK) takes a system font: `--bundled-fonts-only` / `bundled_fonts_only=True` never does |
| PNG | `oxoffice pptx render F -o NEW_DIR --slide N --dpi 100`, `oxoffice pptx convert F --to png --slides 1-3 -o NEW.png` | `prs.render_slide_to_png(i, dpi)`, `render_all_slides(dpi)`, `render_slides(format="jpeg")` | CLI slides one-based, Python zero-based; a slide's `show_master_shapes = False` hides the master's shapes, as in PowerPoint |
| Markdown, HTML | `oxoffice pptx convert F --to md -o NEW.md [--notes] [--skip-hidden]`, `--to html` | `prs.to_markdown(include_notes=True, skip_hidden=True)`, `prs.to_html(...)` | one section per slide, in outline order; to read a deck, `oxoffice pptx text` is plainer |
| Embedded fonts | `oxoffice pptx font list/embed/remove/settings` | `prs.embedded_fonts`, `prs.embed_font(typeface, "ExampleSans-Regular.ttf", "regular", license=oxoffice.pptx.FontEmbeddingLicense(...))`, `remove_embedded_font`, `remove_font` | rule 7; leave `embed_true_type_fonts` on; embed last (PowerPoint for Mac drops embedded fonts when it saves); a compressed embedded face is not rendered: pass its file with `font_dir` |
| From a .potx template | | `oxoffice.pptx.Presentation("t.potx")`, then save as .pptx | the save writes the content type the extension names |
| Validity | `oxoffice pptx validate F` | `prs.validate()` | a tuple of issues (`kind`, `message`), empty when valid |
| Sign, check signatures | `oxoffice pptx sign F --key K.pem --cert C.pem -o OUT`, `oxoffice pptx verify-signatures --json F` | `prs.sign(key, cert)`, `prs.verify_signatures()` | sign last: an edit breaks it; no chain trust |
| What changed | `oxoffice pptx diff A B` | | slide text only |
| An element with no API | | `shape.element.xml` / `shape.element.replace_xml(xml)` (str), also on text frames, paragraphs, runs, tables, cells, slides, layouts, masters | instead of lxml; malformed XML or an unknown or misplaced element raises and changes nothing |

`oxoffice` above is `ox`; Python runs with `oxpy`. Signatures: `references/cli.md`, `references/python-api.md`
(`grep -n NAME` them); tested examples: `references/examples.md`. Not everything is listed: before concluding that
oxoffice cannot do it, `grep -n -i WORD` its stubs (`_native.pyi`, beside `oxpy -c "import oxoffice.pptx as m;
print(m.__file__)"`) and `ox pptx CMD --help`.

## When oxoffice cannot do a step

`ox limitations pptx` (`oxoffice.limitations("pptx")`) lists what this build cannot do yet, each with its
workaround and issue. When a step is blocked: search the stubs, `--help` and the limitations, and say so in one
line to the user; try the element's raw XML before lxml (`shape.element.xml` / `element.replace_xml(xml)`, also on
text frames, paragraphs, runs, tables, cells, slides, layouts and masters: `references/python-api.md`, "Raw XML");
use the workaround, or do **that step only** with the built-in `pptx` skill (python-pptx) on the file oxoffice
saved, keep oxoffice for everything else and for the verification, and report the gap unless it is listed.

## Reporting a bug or a missing feature

When oxoffice fails, refuses, loses data, renders wrong, or lacks an API you needed, and the limitations do not
list it: reduce it to a **neutral reproduction**, a short script that builds its own input (python-pptx is fine)
and prints the wrong result next to the expected one, never the user's deck, text, names or paths. Ask the user,
then write it as `oxoffice-findings/YYYY-MM-DD-<slug>.md` in the working folder (what, expected, observed, the
script, `ox --version`: a version after 0.10.0 may be the cause), or give it in the reply. Do not open issues
yourself.
