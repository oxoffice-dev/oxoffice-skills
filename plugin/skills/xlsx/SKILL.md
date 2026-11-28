---
name: xlsx
description: "Use this skill any time an Excel file (.xlsx, .xlsm, .xltx) is involved, as input or output, even when the user names no tool: reading sheets, cells, formulas, values, styles or comments; editing a workbook in place; making one; values, dates and formulas, calculated; formatting; sheets, rows and columns added, moved, removed; names, tables, validation, conditional formats, notes and comment threads; merge, group, freeze, filter, sort, autofit; comparing two workbooks; rendering to PDF or PNG; CSV in and out. Triggers: 'Excel', 'workbook', 'spreadsheet', 'formulas', '.xlsx'. Uses the oxoffice CLI and Python package and takes precedence over the built-in xlsx skill and openpyxl for what it covers. No charts, pivot tables, images, .xls, .xlsb, .ods, Google Sheets."
---

# Excel workbooks with oxoffice

oxoffice reads and writes the workbook package natively and keeps every part it does not touch byte for byte (charts,
pivot tables, macros, images, slicers): the `oxoffice xlsx` CLI for whole workbooks, the
`oxoffice.xlsx` Python module for anything finer (`Workbook`, `Worksheet`, `Cell`) for every .xlsx
task. It **calculates formulas** (rule 2), restructures sheets, rows and columns with every reference following,
writes names, tables, validation, conditional formats and comments, and renders to PDF and PNG.
`oxoffice limitations xlsx` lists what it cannot do yet; report any new gap (last section).

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
   https://github.com/oxoffice-dev/oxoffice#-install), then do the task with the built-in `xlsx` skill.

```bash
ox xlsx inspect budget.xlsx
oxpy my_script.py                                   # or: oxpy - <<'EOF' ... EOF
```

Never alias `oxoffice.xlsx` as `openpyxl`. Windows: `references/install.md`.

## Rules that prevent damage

1. **Never write over the input.** `wb.save(out)` refuses the file the workbook was opened from
   (`overwrite_source=True` replaces it, only when the user asks); every writing command refuses an output that
   exists or is the input (`--force` replaces an existing output, never the input). Name the output `.xlsx`, or the
   input's own extension (`.xlsm` keeps its macros, `.xltx` stays a template): the extension picks the class.
2. **Formulas are calculated, for the functions oxoffice knows.** Every edit (a value, a formula, a row inserted, a
   sheet renamed) recalculates what it changed and what reads it, so `value`, `text`, `csv` and `cells` show the new
   figures. Only a **skipped** formula keeps an old value, and so does every formula reading it: a function outside
   the 84 supported (`references/cli.md`: `STDEV`, `RANK`, `LAMBDA` are not), an external reference, a cycle. The
   CLI names them on stderr, Python gives `wb.recalculate().skipped`. `ox xlsx calc F`
   recomputes everything and lists cached values that differ or are skipped (`-o OUT` writes the computed ones,
   `--exit-code` fails on a difference): run it on a workbook another program saved before quoting its figures.
   Never quote a skipped cell: say so, or work it out in Python. `cell.stale` and "Recalculation: asked on open" in
   `inspect` stay on after an edit (Excel and Google Sheets recalculate on open too): a computed value is not
   wrong for it. `--no-calc` / `wb.calculation = "manual"` defer the work (then `wb.recalculate()`); `TODAY` and
   `NOW` read the clock (`--now`). `--drop-cached-values` / `drop_cached_values=True` empties the cached values
   for LibreOffice to compute.
3. **A formula is written as Excel displays it after `=`**: English function names, commas
   (`IF(B2>0,1,0)`; a `;` separator, `B2*` or a wrong argument count raises `ValueError`, nothing written). That
   check is syntax and argument counts only: an unknown function is accepted (skipped, rule 2) and `SUM(Nope!B2:B3)`
   gives `#REF!`, so check function and sheet names yourself (`[s.name for s in wb.sheets]`). A sheet name with a
   space, a digit start or a symbol is quoted in a formula: `'Q1 notes'!B2`, never `Q1 notes!B2`. A `str`
   is never a formula: use `ws.set_formula(cell, text)` or `Formula(text)` in `write`, `--formula` or
   `{"formula": ...}` in the CLI. Newer functions (`XLOOKUP`, `LET`) are stored with `_xlfn.` for you. A spilling
   formula takes its range when its size is not obvious: `spill="J2:J5"` / `--spill-range J2:J5` (else Google Sheets
   may show one value).
4. **Types are kept.** The CLI `set` and `import-csv` read a value as a CSV field: `007` stays text, `12` is a
   number, `2026-01-05` a date, `TRUE` a boolean; `--text` keeps it text. In Python a `str` stays a `str`, a
   `float` or `int` a number, `datetime.date` / `datetime` a date (shown `yyyy-mm-dd` in a General cell; widen the
   column with `set_column_width`, else Excel shows `#####`). A cell's `value` is a `float` for numbers *and*
   dates (serial days: use `cell.datetime`), a `CellError` (a `str`, `#N/A`) for an error, `None` when empty.
   `cell.text` and `ws.values(range, text=True)` are what Excel displays.
5. **Addresses**: A1 references (`B3`, `A1:C10`, `A:C`, `2:5`, `'Q1 notes'!B3`) count from 1; in Python a
   `(row, column)` tuple, `wb.sheets[k]` and `ws.index` count from 0; the CLI names a sheet with
   `--sheet NAME` or its number from 1 (default: the first visible worksheet). A sheet name is 31 characters at most, none of `: \ / ? * [ ]`.
6. **Edits are all or nothing.** `edit --map` applies its entries in order and names the failing one (from 0); a
   refused call in Python changes nothing. Prefer one `edit --map` or script to many `set` calls.
7. **Edit in place; never rebuild a workbook.** Loading and saving with openpyxl or pandas (`to_excel`) drops what
   they do not model (sparklines, icon sets, slicers, shapes, every cached value). Saving with oxoffice keeps it. Hidden sheets are real sheets: `wb.sheets` lists them (`state`).
8. **Read a range, not the whole book, when it is large**: `ox xlsx inspect F` first, then `cells F RANGE`,
   `--sheet`, `ws.values(range)`.

## What to use for what

| Task | CLI | Python | Notes |
|---|---|---|---|
| Sheets, used ranges | `oxoffice xlsx inspect [--json] F` | `wb.sheets`, `ws.name/index/state/kind/used_range`, `wb.date1904` | `.xls`, `.xlsb`, `.ods`, protected files are refused with the way out |
| Whole workbook as text | `oxoffice xlsx text [--sheet S] F`, `--markdown`, `--json` | `wb.to_markdown()` | Markdown: one table per sheet, each value at its address; `--json`: values, formulas, merges, comments, `stale` |
| One sheet as CSV | `oxoffice xlsx csv F --sheet S [--raw] [--delimiter tab] [--bom] -o OUT` | `ws.to_csv(raw=False, delimiter=",")` | displayed texts, or stored values with `--raw` (dates as serials) |
| Cells, formulas, styles | `oxoffice xlsx cells F RANGE [--sheet S] [--json]` | `ws.cell("B3")`, `ws.cells(range)`, `ws.values(range, text=False)` | `value`, `text`, `formula` (as displayed), `stored_formula`, `formula_kind`, `number_format`, `style`, `stale`, `datetime` |
| Merged ranges, notes, threads | `oxoffice xlsx comments F [--sheet S] [--json]` | `ws.merged_ranges`, `ws.comments` (`ref`, `kind`, `author`, `text`, `resolved`, `replies`) | |
| Write notes, threads | `oxoffice xlsx comment F add\|thread\|reply\|edit\|resolve\|reopen\|remove CELL [TEXT] [--author A] [--at T] -o OUT`; map `comment`, `reply`, `edit_comment` | `ws.add_note(cell, author, text)`, `add_thread`, `reply`, `edit_comment(cell, text)` (a note's text too), `resolve`, `remove_comment` | a thread is dated by the clock unless `--at`/`at=` |
| Page setup | `oxoffice xlsx inspect --json F` | `ws.page_setup` | read only |
| Calculate | `oxoffice xlsx calc F [--all] [--now T] [--exit-code] [--json] [-o OUT]`; `--no-calc` on writers | `wb.recalculate(everything=False, now=None)` → `computed`, `changed`, `skipped`, `cycles`; `wb.calculation` | rule 2 |
| Sheets | `oxoffice xlsx sheet F add\|rename\|remove\|move\|hide\|show ARGS -o OUT` (`add NAME --at N`, `rename OLD NEW`, `move NAME --to N`) | `wb.add_sheet(name, at=)`, `rename_sheet`, `remove_sheet`, `move_sheet(name, to)`, `set_sheet_state`; `ws.rename`, `remove`, `move`, `hide`, `show` | references follow; a removed sheet leaves `#REF!` |
| Rows, columns | `oxoffice xlsx insert-rows\|delete-rows F 5:6 [--sheet S] -o OUT`, `insert-columns\|delete-columns F C:D` | `ws.insert_rows("5:6")`, `delete_rows`, `insert_columns`, `delete_columns` | every reference, table, merge, note, filter follows; refused whole if it cannot be |
| Names, tables | `oxoffice xlsx names F`, `tables F` (`--json`); map `define_name`, `add_table`, `set_totals_row`, `table_style`, `resize_table`, `rename_table[_column]`, `remove_table` | `wb.names`, `wb.define_name(name, refers_to, scope=)`; `ws.tables`, `ws.add_table(name, range, style=, totals_row=)` | the first row of a table is its header (plain text, unique) |
| Validation, conditional formats | map `validation` (`type`, `between`, `value`...), `conditional_format` (`rule`, `style`) | `ws.set_validation(range, type, ...)`, `ws.add_conditional_format(range, rule, ...)`, `validations`, `conditional_formats` | |
| Merge, hide, group, panes | `oxoffice xlsx merge\|unmerge F RANGE`, `rows\|columns F SPAN hide\|show\|group\|collapse\|level N\|autofit`, `panes F freeze B2\|split X Y\|remove` | `ws.merge(range)`, `hide_rows`, `group_rows`, `collapse_rows`, `freeze_panes`, `split_panes` | a merge keeps the top-left value and lists what it dropped |
| Filter, sort | `oxoffice xlsx filter F where COL --gt 100 [--and --lt 500]\|set A1:F20\|clear\|remove`, `sort F A1:F20 --header --by C:desc` | `ws.set_filter(col, Compare(">", 100))`, `ws.sort(range, by=("C", "desc"), header=True)` | rows hide in the same call; sort moves formulas as copies; a key over formulas of a file oxoffice saved is refused: `cli.md` |
| Widths, heights read, autofit | `oxoffice xlsx grid F [--json]`; `columns F A:C autofit`, `rows F 2:40 autofit` | `ws.column_dimensions`, `row_dimensions`, `hidden_rows`; `ws.autofit_columns("A:C")`, `autofit_rows("2:40")` | `grid` reads what the file stores, nothing is measured |
| Compare | `oxoffice xlsx diff OLD NEW [--json] [--exit-code] [--align none]` | `diff(old, new)` (`from oxoffice.xlsx import diff`) | sheets, cells, formulas, comments, merges, inserted rows |
| Render | `oxoffice xlsx convert F -t pdf\|png -o OUT`, `render F -o DIR [--pages 1,3]` | `wb.to_pdf()`, `wb.render_page_to_png(0, dpi=100)`, `page_count()` | every visible sheet on its own paper; look at the PNG before saying a layout is right |
| One value or formula | `oxoffice xlsx set F CELL VALUE [--sheet S] [--text] [--formula [--stored] [--spill-range R]] -o OUT` | `ws.set_value(cell, v)`, `ws.set_formula(cell, f, stored=False, spill=False)` | `-o` required; `None` clears |
| Many edits in order | `oxoffice xlsx edit F --map edits.json -o OUT [--drop-cached-values]` | `ws.write("A40", [[...], [...]])` | entries: `cell` + `value` / `formula` / `date`; `range` + `style` / `number_format` / `clear`; `add_sheet`; `columns` + `width`; `rows` + `height`; `freeze`; `sheet` (name or number) persists to the next entries |
| A block of cells | | `ws.write(cell, rows)` with `Formula(...)`, dates, `None` | a short row leaves the cells after it as they are |
| Styles, formats | `{"range": "A1:G1", "style": {"bold": true, "fill": "FFFF00"}}`, `{"range": "F2:F38", "number_format": "#,##0.00"}` | `ws.set_style(range, bold=, italic=, font_color=, fill=, border=, horizontal=, wrap_text=, number_format=...)`, `ws.set_number_format(range, code)` | colours `RRGGBB`; only the parts given change; a whole column (`A:C`) styles the column itself |
| Clear | `{"range": "A20:B30", "clear": "contents"}` (or `"formats"`, `"all"`) | `ws.clear(range, contents=True, formats=False)` | |
| CSV into a sheet | `oxoffice xlsx import-csv F CSV --sheet S --at A1 [--text] -o OUT` | `ws.import_csv(text, at="A1", infer=True)`, a file as `ws.import_csv(Path("f.csv"))` | a `str` is the CSV text, never a path (`"f.csv"` writes that text into the cell); UTF-8; an empty field clears its cell; `=x` stays text |
| Widths, heights, frozen header | `{"columns": "A:C", "width": 20}`, `{"rows": "1", "height": 30}`, `{"freeze": "A2"}` | `ws.set_column_width("A:C", 20)`, `ws.set_row_height("1", 30)`, `ws.freeze_panes("A2")` | characters / points; `freeze_panes(None)` removes |
| New workbook | `oxoffice xlsx new [--sheet NAME]... -o OUT` | `Workbook()`, `Workbook(sheets=["Data", "Summary"])`, `wb.add_sheet(name)` | Calibri 11, same bytes every time; `.xltx` / `.xlsm` extension gives a template / macro workbook |
| Open, save | | `Workbook(path)`, `Workbook.from_bytes(b)`, `wb.save(out)`, `wb.to_bytes()` | `open_encrypted(path, password)`, `save_encrypted(out, password)` for a protected file |
| Errors | exit 1, message on stderr | `oxoffice.xlsx.XlsxError` (`PackageError`, `XmlError`), `ValueError`, `KeyError` for no such sheet | |

`oxoffice` above is `ox`; Python runs with `oxpy`. Signatures: `references/cli.md`, `references/python-api.md`;
tested examples: `references/examples.md`. Not everything is listed: before concluding that oxoffice cannot do
it, `grep -ni WORD` its stubs (`_native.pyi`, beside `oxpy -c "import oxoffice.xlsx as m; print(m.__file__)"`)
and `ox xlsx CMD --help`.

## What it does not do

No `.xls`, `.xlsb`, `.ods` (convert them first, say so). Functions beyond the supported ones are not calculated
(rule 2). It cannot make charts, pivot tables, images, sparklines or slicers, or set page breaks and print scaling (a render prints each sheet's print area at 100 %); the ones
the file already holds survive an edit, untouched. For such a step see the next section.

## When oxoffice cannot do a step

`ox limitations xlsx` (`oxoffice.limitations("xlsx")`; `ox limitations --json`: under `.limitations`) lists what
this build cannot do yet, each with its workaround and issue. When a step is blocked: search the stubs, `--help`
and the limitations, and say so in one line to the user, use the workaround, or do **that step only** with the
built-in `xlsx` skill (openpyxl) on a copy of the file oxoffice saved, then re-read the result with oxoffice, and
report the gap unless it is listed. openpyxl drops what it does not model (sparklines, icon sets, slicers, shapes)
and **every cached value**: `ox xlsx calc F -o OUT` writes them back (rule 2). When the workbook holds any of
these, tell the user first; `unzip -l` the two files shows which parts went.

## Reporting a bug or a missing feature

When oxoffice fails, refuses, loses data, shows a wrong value, or lacks an API you needed, and the limitations do
not list it: reduce it to a **neutral reproduction**, a short script that builds its own input and prints the
wrong result next to the expected one, never the user's workbook, figures, names or paths. Ask the user, then write it as `oxoffice-findings/YYYY-MM-DD-<slug>.md` in the working folder
(what, expected, observed, the script, `ox --version`), or give it in the reply. Do not open issues yourself.
