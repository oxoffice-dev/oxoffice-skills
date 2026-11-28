---
name: xlsx
description: "Use this skill any time an Excel file (.xlsx, .xlsm, .xltx) is involved, as input or output, even when the user names no tool: reading sheets, cells, formulas, values, styles or comments, also to use them elsewhere; making a workbook; setting values, dates and formulas; formatting, number formats, column widths, frozen panes; importing a CSV into a sheet, exporting a sheet to CSV or Markdown. Triggers: 'Excel', 'workbook', 'spreadsheet', 'sheet', 'cells', 'formulas', '.xlsx', 'CSV to Excel'. Uses the oxoffice CLI and Python package and takes precedence over the built-in xlsx skill and openpyxl for what it covers. Does not recalculate, make charts or pivot tables, or open .xls, .xlsb, .ods or Google Sheets."
---

# Excel workbooks with oxoffice

oxoffice reads and writes the workbook package natively and keeps every part it does not touch byte for byte (charts,
pivot tables, macros, images, conditional formats, data validation): the `oxoffice xlsx` CLI for whole workbooks,
the `oxoffice.xlsx` Python module for anything finer (`Workbook`, `Worksheet`, `Cell`). Use them for every .xlsx
task. It has **no formula evaluator**: it stores what you write and Excel computes it on open (rule 2).
`oxoffice limitations xlsx` lists what it cannot do yet; report any new gap (last section).

## Setup, once per session

Needs **oxoffice** 0.9.0 or later (written and checked with 0.9.0): the one installed, else through a package
manager. Take the first way that works, keep it. Each defines two shell functions (bash, zsh): `ox`, the CLI, and
`oxpy`, a Python with oxoffice. Shell state is lost between tool calls: define them first.

1. **Already installed**: `oxoffice --version` and `python3 -c "import oxoffice; print(oxoffice.__version__)"`
   both print 0.9.0 or later: `ox() { oxoffice "$@"; }; oxpy() { PYTHONUTF8=1 python3 "$@"; }`. Older: leave it.
2. **uv**: `ox() { uvx --python '>=3.10' --from 'oxoffice>=0.9.0' oxoffice "$@"; }`,
   `oxpy() { PYTHONUTF8=1 uv run --no-project --python '>=3.10' --with 'oxoffice>=0.9.0' python "$@"; }`.
3. **pip in a virtual environment** of its own, made once (`Scripts` for `bin` on Windows):

   ```bash
   V="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-venv"
   [ -x "$V/bin/oxoffice" ] || { python3 -m venv "$V" && "$V/bin/pip" install -q 'oxoffice>=0.9.0'; }
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
   input's own extension (`.xlsm` keeps its macros, `.xltx` stays a template): the extension picks the class and
   nothing refuses `x.xls`.
2. **Nothing is calculated.** A formula you write has no cached value; after any edit the cached values of the
   formulas that depend on it are out of date (`cell.stale`, `wb.recalculation_needed`, `"stale": true` in
   `cells --json`). Excel and Google Sheets recalculate on open. So: never report a formula's value read after an
   edit as the result, work the figure out in Python and say it is not stored; for LibreOffice, which does not
   calculate on open, save with `--drop-cached-values` / `drop_cached_values=True`; for a program that reads the
   file without calculating (pandas, a script), write the value, not a formula, or say the cells are empty until
   Excel opens it.
3. **A formula is written as Excel displays it after `=`**: English function names, commas
   (`IF(B2>0,1,0)`; a `;` separator, `B2*` or a wrong argument count raises `ValueError`, nothing written). A `str`
   is never a formula: use `ws.set_formula(cell, text)` or `Formula(text)` in `write`, `--formula` or
   `{"formula": ...}` in the CLI. Newer functions (`XLOOKUP`, `LET`) are stored with `_xlfn.` for you. A spilling
   formula takes its range, `spill="J2:J5"` / `--spill-range J2:J5` (without one, Google Sheets shows one value).
4. **Types are kept.** The CLI `set` and `import-csv` read a value as a CSV field: `007` stays text, `12` is a
   number, `2026-01-05` a date, `TRUE` a boolean; `--text` keeps it text. In Python a `str` stays a `str`, a
   `float` or `int` a number, `datetime.date` / `datetime` a date (shown `yyyy-mm-dd` in a General cell; widen the
   column with `set_column_width`, else Excel shows `#####`). A cell's `value` is a `float` for numbers *and*
   dates (serial days: use `cell.datetime`), a `CellError` (a `str`, `#N/A`) for an error, `None` when empty.
   `cell.text` and `ws.values(range, text=True)` are what Excel displays.
5. **Addresses**: A1 references (`B3`, `A1:C10`, `A:C`, `2:5`, `'Q1 notes'!B3`) are the spreadsheet's own, from 1;
   in Python a `(row, column)` tuple, `wb.sheets[k]` and `ws.index` count from 0; the CLI names a sheet with
   `--sheet NAME` or its number from 1 (the first visible worksheet by default, so name the sheet on a multi-sheet
   workbook). A sheet name is 31 characters at most, none of `: \ / ? * [ ]`.
6. **Edits are all or nothing.** `edit --map` applies its entries in order and names the failing one (from 0); a
   refused call in Python changes nothing. Prefer one `edit --map` or one script to many `set` calls.
7. **Edit in place; never rebuild a workbook.** Loading and saving with openpyxl or pandas (`to_excel`) drops what
   they do not model (charts, pivot tables, images, conditional formats). Saving with oxoffice keeps
   it. Hidden sheets are real sheets: `wb.sheets` lists them (`state`), `to_markdown` prints them.
8. **Read a range, not the whole book, when it is large**: `ox xlsx inspect F` first (sheets and used ranges), then
   `cells F RANGE`, `--sheet`, `ws.values(range)`.

## What to use for what

| Task | CLI | Python | Notes |
|---|---|---|---|
| Sheets, used ranges, date system | `oxoffice xlsx inspect [--json] F` | `wb.sheets`, `ws.name/index/state/kind/used_range`, `wb.date1904` | `.xls`, `.xlsb`, `.ods` and password-protected files are refused with the way out |
| Whole workbook as text | `oxoffice xlsx text [--sheet S] F`, `--markdown`, `--json` | `wb.to_markdown()` | Markdown: one table per sheet, each value at its address; `--json`: values, formulas, merges, comments, `stale` |
| One sheet as CSV | `oxoffice xlsx csv F --sheet S [--raw] [--delimiter tab] [--bom] -o OUT` | `ws.to_csv(raw=False, delimiter=",")` | displayed texts, or stored values with `--raw` (dates as serials) |
| Cells, formulas, styles | `oxoffice xlsx cells F RANGE [--sheet S] [--json]` | `ws.cell("B3")`, `ws.cells(range)`, `ws.values(range, text=False)` | `value`, `text`, `formula` (as displayed), `stored_formula`, `formula_kind`, `number_format`, `style`, `stale`, `datetime` |
| Merged ranges, notes, threads | `oxoffice xlsx comments F [--sheet S] [--json]` | `ws.merged_ranges`, `ws.comments` (`ref`, `kind`, `author`, `text`, `resolved`, `replies`) | read only: comments cannot be added |
| Page setup | `oxoffice xlsx inspect --json F` | `ws.page_setup` | read only |
| One value or formula | `oxoffice xlsx set F CELL VALUE [--sheet S] [--text] [--formula [--stored] [--spill-range R]] -o OUT` | `ws.set_value(cell, v)`, `ws.set_formula(cell, f, stored=False, spill=False)` | `set_value(cell, None)` clears; `-o` required |
| Many edits in order | `oxoffice xlsx edit F --map edits.json -o OUT [--drop-cached-values]` | `ws.write("A40", [[...], [...]])` | entries: `cell` + `value` / `formula` / `date`; `range` + `style` / `number_format` / `clear`; `add_sheet`; `columns` + `width`; `rows` + `height`; `freeze`; `sheet` (name or number) persists to the next entries |
| A block of cells | | `ws.write(cell, rows)` with `Formula(...)`, dates, `None` | a short row leaves the cells after it as they are |
| Styles, formats | `{"range": "A1:G1", "style": {"bold": true, "fill": "FFFF00"}}`, `{"range": "F2:F38", "number_format": "#,##0.00"}` | `ws.set_style(range, bold=, italic=, underline=, strike=, font_name=, font_size=, font_color=, fill=, border=, border_color=, horizontal=, vertical=, wrap_text=, number_format=)`, `ws.set_number_format(range, code)` | colours `RRGGBB`; only the parts given change; a whole column (`A:C`) styles the column itself |
| Clear | `{"range": "A20:B30", "clear": "contents"}` (or `"formats"`, `"all"`) | `ws.clear(range, contents=True, formats=False)` | |
| CSV into a sheet | `oxoffice xlsx import-csv F CSV --sheet S --at A1 [--text] -o OUT` | `ws.import_csv(text_or_path, at="A1", infer=True)` | UTF-8; an empty field clears its cell; `=x` stays text |
| Widths, heights, frozen header | `{"columns": "A:C", "width": 20}`, `{"rows": "1", "height": 30}`, `{"freeze": "A2"}` | `ws.set_column_width("A:C", 20)`, `ws.set_row_height("1", 30)`, `ws.freeze_panes("A2")` | characters / points; `freeze_panes(None)` removes |
| New workbook | `oxoffice xlsx new [--sheet NAME]... -o OUT` | `Workbook()`, `Workbook(sheets=["Data", "Summary"])`, `wb.add_sheet(name)` | Calibri 11, same bytes every time; `.xltx` / `.xlsm` extension gives a template / macro workbook |
| Open, save | | `Workbook(path)`, `Workbook.from_bytes(b)`, `wb.save(out, drop_cached_values=False)`, `wb.to_bytes()` | `open_encrypted(path, password)`, `save_encrypted(out, password)` for a password-protected file |
| Errors | exit 1, message on stderr | `oxoffice.xlsx.XlsxError` (`PackageError`, `XmlError`), `ValueError`, `KeyError` for no such sheet | |

`oxoffice` above is `ox`; Python runs with `oxpy`. Signatures: `references/cli.md`, `references/python-api.md`
(`grep -n NAME` them); tested examples: `references/examples.md`. Not everything is listed: before concluding that
oxoffice cannot do it, `grep -n -i WORD` its stubs (`_native.pyi`, beside `oxpy -c "import oxoffice.xlsx as m;
print(m.__file__)"`) and `ox xlsx CMD --help`.

## What it does not do

No evaluator (rule 2), no `.xls`, `.xlsb`, `.ods` (convert them first, say so), no rendering to PDF or PNG. It
cannot make charts, pivot tables, images, tables, conditional formats, data validation or defined names, merge
cells, add comments, insert or delete rows and columns, or move, rename or remove sheets; the ones the file
already holds survive an edit, untouched. For such a step see the next section.

## When oxoffice cannot do a step

`ox limitations xlsx` (`oxoffice.limitations("xlsx")`) lists what this build cannot do yet, each with its
workaround and issue. When a step is blocked: search the stubs, `--help` and the limitations, and say so in one
line to the user, use the workaround, or do **that step only** with the built-in `xlsx` skill (openpyxl) on the
file oxoffice saved, then reopen the result with oxoffice to verify it, and report the gap unless it is listed.
openpyxl drops charts and pivot tables it did not make: do not run it over a workbook that holds them; say what the
step needs instead.

## Reporting a bug or a missing feature

When oxoffice fails, refuses, loses data, shows a wrong value, or lacks an API you needed, and the limitations do
not list it: reduce it to a **neutral reproduction**, a short script that builds its own input (XlsxWriter or
`Workbook()` is fine) and prints the wrong result next to the expected one, never the user's workbook, figures,
names or paths. Ask the user, then write it as `oxoffice-findings/YYYY-MM-DD-<slug>.md` in the working folder
(what, expected, observed, the script, `ox --version`: a version after 0.9.0 may be the cause), or give it in the
reply. Do not open issues yourself.
