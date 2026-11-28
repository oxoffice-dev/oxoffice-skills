# oxoffice xlsx CLI reference

`ox xlsx <command>`, with `ox` from SKILL.md's Setup (here written `oxoffice xlsx`). Exit codes: 0 success; 1 error,
with the message on stderr; 2 bad command line; 101 an internal panic (report it). Sheets are named, or numbered from
1 (`--sheet NAME|N`); without `--sheet` a command takes the first visible worksheet, except `text` and `comments`,
which take every sheet. A cell or range may carry its sheet (`'Q1 notes'!B3`, `Budget!A1:C10`) or take `--sheet`.
Ranges: `B3`, `A1:C10`, `A:C` (whole columns, clipped to the used range), `2:5`.

**Outputs.**

| Commands | An existing output |
|---|---|
| `new`, `set`, `edit`, `import-csv` (`-o` required), `csv` (`-o` optional: stdout without it) | refused, exit 1, "output already exists ... (pass --force to replace it)"; `--force` replaces an existing output, never the input ("output is the input file") |

Every writing command saves the extension it is given as the workbook class (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`)
and does not check it against the input: write `.xlsx` for a `.xlsx` input. `--json` prints what was done
(`operation`, `output`, `sheet`, `cell` or `edits`, `schema`). `--drop-cached-values` goes with `set`, `edit` and
`import-csv` (not `new`): it empties the cached value of every formula so that LibreOffice computes them on open (Excel and Google
Sheets calculate an edited workbook anyway; a function LibreOffice lacks, such as `LAMBDA`, then shows an error).

## Reading

| Command | Output |
|---|---|
| `oxoffice xlsx inspect [--json] F` | the file, its class, date system, whether a recalculation is asked, and each sheet: position, name, kind (`worksheet`, `chartsheet`, `dialogsheet`, `macrosheet`), state (hidden ones listed), used range. `--json` adds each sheet's page setup (margins, paper, orientation, scale, header and footer, breaks, print area and titles) |
| `oxoffice xlsx text [--sheet S]... F` | the cached values (rule 2 of SKILL.md: `inspect F` says whether they may be stale); per sheet `# NAME` then its rows, tab-separated, the texts Excel displays (empty rows left out: no addresses; use `--markdown` or `cells` when addresses matter) |
| `oxoffice xlsx text --markdown [--sheet S]... F` | `## NAME` and one table per sheet, headed by the column letters, each row led by its number: every value keeps its address. `\|` is escaped, line breaks are `<br>`, chart and dialog sheets are headed without a table |
| `oxoffice xlsx text --json [--sheet S]... F` | `{"date1904", "recalculation_needed", "schema", "sheets": [{"sheet" (from 1), "name", "kind", "state", "used_range", "merged", "comments", "cells": [{"ref", "text", "type" (`text`, `number`, `boolean`, `error`), "value" (a number, a date as its serial), "formula" and `"stale"` when it is a formula}]}]}`; only cells that hold something |
| `oxoffice xlsx csv [--sheet S] [--raw] [--delimiter C\|tab] [--bom] [-o OUT] F` | one sheet from A1 to the end of its used range: CRLF line ends, RFC 4180 quoting; the displayed texts, or `--raw` the stored values (shortest number, dates as serials, `TRUE` / `FALSE`, error codes) |
| `oxoffice xlsx cells [--sheet S] [--json] F RANGE` | the cells that hold something, `REF<TAB>text<TAB>=formula`. `--json`: `{"name", "sheet", "range", "merged", "cells": [...]}`, each cell with `ref`, `type`, `value`, `text`, `formula` (as Excel displays it), `stored_formula` (`_xlfn.XLOOKUP(...)`), `formula_kind` (`shared`, `array`, `data_table`, null), `array_range`, `stale`, `number_format` and the resolved `style` (`font`, `fill`, `border`, `alignment`, `protection`, colours as `rgb` with their source) |
| `oxoffice xlsx comments [--sheet S]... [--json] F` | notes and comment threads: `Sheet!REF (note\|thread) Author: text`, replies indented. `--json`: per sheet `comments` with `ref`, `kind`, `author`, `date` (threads only), `text`, `resolved`, `replies` |

A file that is not a workbook is refused with its way out (`.xls`, `.xlsb`, `.ods`, a Word or PowerPoint file, a
password-protected workbook: open that one from Python with `Workbook.open_encrypted`).

## Writing

```bash
oxoffice xlsx new [--sheet NAME]... -o OUT [--force] [--json]      # Sheet1 when no --sheet
oxoffice xlsx set F CELL VALUE [--sheet S] [--text | --formula [--stored] [--spill | --spill-range J2:J5]] -o OUT
oxoffice xlsx import-csv F CSV [--sheet S] [--at A1] [--text] -o OUT
oxoffice xlsx edit F --map edits.json -o OUT [--drop-cached-values] [--json]
```

`new` makes Calibri 11 under the Office 2013-2022 theme, the same bytes every time; a `.xltx`, `.xlsm` or `.xltm`
extension makes a template or a macro-enabled workbook.

`set` writes one cell. `VALUE` is read as a CSV field: a number (no leading zero, at most 15 significant digits),
`TRUE` / `FALSE`, a date `YYYY-MM-DD` with a time or not, an empty text clears the cell, any other text as it is;
`--text` keeps it text whatever it looks like (`007`); `--formula` makes it a formula as Excel displays it (an
`=` is accepted), `--stored` as the file stores it. `--spill` writes a dynamic array formula Excel spills from the
cell (Google Sheets shows its first value alone); `--spill-range J2:J5` names the range from the cell, other cells
empty, which Google Sheets spills too. A value that starts with `-` and is not a number goes last, after `--`
(`set F A1 --formula -o OUT -- -A1*2`). A formula Excel would repair (`=1+`, a `;` separator) is refused with the
reason; the check is syntax and argument counts, not function or sheet names (`SUM(Nope!B2)` is written), and a
sheet name with a space is quoted: `'Q1 notes'!B2`.

`import-csv` writes the rectangle the CSV covers from `--at`, an empty field clearing its cell. UTF-8, a byte order
mark skipped. Fields are typed strictly (number, boolean, date, else text; a leading `=` stays text: a CSV holds no
formula); `--text` keeps every field text.

`edit` applies a JSON array of edits in order, all or nothing; a refusal names the entry, counted from 0, and writes
nothing. Each entry does one thing:

| Entry | Does |
|---|---|
| `{"cell": "B2", "value": 3}` | a number, a string, `true`, `false`, or `null` (clears the content); a string is never a formula |
| `{"cell": "D2", "formula": "B2*C2"}` | a formula; `"stored": true` for the stored form, `"spill": true` or `"spill": "J2:J5"` for a spilling one |
| `{"cell": "E2", "date": "2026-01-05"}` | a date, a date and time (`2026-01-05 14:30:00`) or a time |
| `{"range": "A1:F1", "style": {"bold": true, "fill": "FFFF00"}}` | changes the parts given: `bold`, `italic`, `underline`, `strike`, `font_name`, `font_size`, `font_color`, `fill`, `border` (`thin`, `medium`, `thick`, `double`, `none`), `border_color`, `horizontal`, `vertical`, `wrap_text`, `number_format` |
| `{"range": "C2:C8", "number_format": "0.00"}` | a number format code (`yyyy-mm-dd`, `#,##0.00`, `0%`) |
| `{"range": "A20:B30", "clear": "contents"}` | `contents`, `formats` or `all` |
| `{"add_sheet": "Summary"}` | a sheet after the others; it becomes the sheet of the entries after it |
| `{"columns": "A:C", "width": 20}` | widths in characters (above 0, at most 255) |
| `{"rows": "2:5", "height": 30}` | heights in points (at most 409.5) |
| `{"freeze": "B2"}` or `{"freeze": null}` | freezes the rows above and the columns left of a cell, or removes the panes |

An entry names its sheet with `"sheet"` (a name, or a number from 1) and otherwise takes the sheet of the entry
before it, the first entry the first visible worksheet: put `"sheet"` on the first entry of each sheet. A style of a
range of cells writes at most 1,000,000 of them, of whole rows at most 100,000 rows.

What Excel would repair is refused with a plain message: a value under a merged range other than its first cell, a
cell of an array formula written alone, a table header that would no longer name its column, a cell of a pivot
table, a sheet name Excel rejects.

## Known limits

`oxoffice limitations xlsx` lists them with their workarounds (`--json`: the same as data). Today: a spilling formula
written without its range shows one value in Google Sheets (give the range). Not in 0.9.0, and so not listed: see
SKILL.md, "What it does not do".
