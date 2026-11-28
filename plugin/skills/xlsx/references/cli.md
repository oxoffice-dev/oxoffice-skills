# oxoffice xlsx CLI reference

`ox xlsx <command>`, with `ox` from SKILL.md's Setup (here written `oxoffice xlsx`). Exit codes: 0 success; 1 error,
with the message on stderr; 2 bad command line; 101 an internal panic (report it). Sheets are named, or numbered from
1 (`--sheet NAME|N`); without `--sheet` a command takes the first visible worksheet, except `text` and `comments`,
which take every sheet. A cell or range may carry its sheet (`'Q1 notes'!B3`, `Budget!A1:C10`) or take `--sheet`.
Ranges: `B3`, `A1:C10`, `A:C` (whole columns, clipped to the used range), `2:5`.

**Outputs.**

| Commands | An existing output |
|---|---|
| every writing command (`new`, `set`, `edit`, `import-csv`, `comment`, `sheet`, `insert-rows` and the other row and column commands, `merge`, `unmerge`, `rows`, `columns`, `outline`, `panes`, `filter`, `sort`; `-o` required), `calc -o`, `csv` (`-o` optional: stdout without it), `convert` | refused, exit 1, "output already exists ... (pass --force to replace it)"; `--force` replaces an existing output, never the input ("output is the input file") |

Every writing command saves the extension it is given as the workbook class (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`)
and does not check it against the input: write `.xlsx` for a `.xlsx` input. `--json` prints what was done
(`operation`, `output`, `sheet`, `cell` or `edits`, `schema`). Every writing command calculates the formulas it
changed and those that read them (an edit map once, at its end): `--no-calc` (on `set`, `edit`, `import-csv`) keeps
the old cached values. `--drop-cached-values` goes with every writing command (not `new`): it empties the cached
value of every formula so that LibreOffice computes them on open (a function LibreOffice lacks, such as `LAMBDA`,
then shows an error).

## Reading

| Command | Output |
|---|---|
| `oxoffice xlsx inspect [--json] F` | the file, its class, date system, whether a recalculation is asked, and each sheet: position, name, kind (`worksheet`, `chartsheet`, `dialogsheet`, `macrosheet`), state (hidden ones listed), used range. `--json` adds each sheet's page setup (margins, paper, orientation, scale, header and footer, breaks, print area and titles) |
| `oxoffice xlsx text [--sheet S]... F` | the cached values (rule 2 of SKILL.md: those of a skipped formula may be old; `calc F` lists them); per sheet `# NAME` then its rows, tab-separated, the texts Excel displays (empty rows left out: no addresses; use `--markdown` or `cells` when addresses matter) |
| `oxoffice xlsx text --markdown [--sheet S]... F` | `## NAME` and one table per sheet, headed by the column letters, each row led by its number: every value keeps its address. `\|` is escaped, line breaks are `<br>`, chart and dialog sheets are headed without a table |
| `oxoffice xlsx text --json [--sheet S]... F` | `{"date1904", "recalculation_needed", "schema", "sheets": [{"sheet" (from 1), "name", "kind", "state", "used_range", "merged", "comments", "cells": [{"ref", "text", "type" (`text`, `number`, `boolean`, `error`), "value" (a number, a date as its serial), "formula" and `"stale"` when it is a formula}]}]}`; only cells that hold something |
| `oxoffice xlsx csv [--sheet S] [--raw] [--delimiter C\|tab] [--bom] [-o OUT] F` | one sheet from A1 to the end of its used range: CRLF line ends, RFC 4180 quoting; the displayed texts, or `--raw` the stored values (shortest number, dates as serials, `TRUE` / `FALSE`, error codes) |
| `oxoffice xlsx cells [--sheet S] [--json] F RANGE` | the cells that hold something, `REF<TAB>text<TAB>=formula`. `--json`: `{"name", "sheet", "range", "merged", "cells": [...]}`, each cell with `ref`, `type`, `value`, `text`, `formula` (as Excel displays it), `stored_formula` (`_xlfn.XLOOKUP(...)`), `formula_kind` (`shared`, `array`, `data_table`, null), `array_range`, `stale`, `number_format` and the resolved `style` (`font`, `fill`, `border`, `alignment`, `protection`, colours as `rgb` with their source) |
| `oxoffice xlsx comments [--sheet S]... [--json] F` | notes and comment threads: `Sheet!REF (note\|thread) Author: text`, replies indented. `--json`: per sheet `comments` with `ref`, `kind`, `author`, `date` (threads only), `text`, `resolved`, `replies` |
| `oxoffice xlsx grid [--sheet S]... [--json] F` | per sheet the default column width and row height, each column run's width (characters), each row's height (points) and outline level, and the hidden rows, each hidden `by filter` (what `SUBTOTAL` 1 to 11 leave out) or `by manual`. What the file stores: nothing is measured |
| `oxoffice xlsx names [--json] F`, `tables [--json] F` | defined names with their scope and formula; tables with sheet, range and columns |
| `oxoffice xlsx diff [--sheet S]... [--align auto\|none] [--json] [--exit-code] OLD NEW` | what changed: sheets (added, removed, renamed, moved, state), cells (added, removed, value, formula; `recalculated` when only a producer's cached value differs), comments (added, removed, text, replies, resolved), merged ranges, inserted and deleted rows (aligned, so the cells below are not all changes). Exit 0; with `--exit-code` 1 when they differ; 2 on an error. At most 100,000 changes are listed, the summary counts all |

A file that is not a workbook is refused with its way out (`.xls`, `.xlsb`, `.ods`, a Word or PowerPoint file, a
password-protected workbook: open that one from Python with `Workbook.open_encrypted`).

## Writing

```bash
oxoffice xlsx new [--sheet NAME]... -o OUT [--force] [--json]      # Sheet1 when no --sheet
oxoffice xlsx set F CELL VALUE [--sheet S] [--text | --formula [--stored] [--spill | --spill-range J2:J5]] [--no-calc] -o OUT
oxoffice xlsx import-csv F CSV [--sheet S] [--at A1] [--text] [--no-calc] -o OUT
oxoffice xlsx edit F --map edits.json -o OUT [--no-calc] [--drop-cached-values] [--json]
```

`new` makes Calibri 11 under the Office 2013-2022 theme, the same bytes every time; a `.xltx`, `.xlsm` or `.xltm`
extension makes a template or a macro-enabled workbook.

`set` writes one cell. `VALUE` is read as a CSV field: a number (no leading zero, at most 15 significant digits),
`TRUE` / `FALSE`, a date `YYYY-MM-DD` with a time or not, an empty text clears the cell, any other text as it is;
`--text` keeps it text whatever it looks like (`007`); `--formula` makes it a formula as Excel displays it (an
`=` is accepted), `--stored` as the file stores it. `--spill` writes a dynamic array formula over its result's range when the formula or
what it reads gives the size (`SORT(A2:A10)`, `UNIQUE(A2:A10)`), else from its cell alone (Excel spills it, Google
Sheets shows its first value); `--spill-range J2:J5` names the range from the cell, other cells empty, which
Google Sheets spills too. A value that starts with `-` and is not a number goes last, after `--`
(`set F A1 --formula -o OUT -- -A1*2`). A formula Excel would repair (`=1+`, a `;` separator) is refused with the
reason; the check is syntax and argument counts, not function or sheet names (`SUM(Nope!B2)` is written and reads `#REF!`), and a
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

### More edit entries

The same map takes the entries below. Each acts on the sheet of the entry.

| Entry | Does |
|---|---|
| `{"rename_sheet": {"from": "Old", "to": "New"}}`, `{"remove_sheet": "Old"}`, `{"move_sheet": "Old", "to": 1}`, `{"sheet_state": "Old", "state": "hidden"}`, `{"add_sheet": "S", "at": 2}` | sheets; state `visible`, `hidden` or `veryHidden`; every reference follows |
| `{"insert_rows": "5:6"}`, `{"delete_rows": "5"}`, `{"insert_columns": "C:D"}`, `{"delete_columns": "C"}` | rows and columns |
| `{"merge": "A1:C2", "across": false}`, `{"unmerge": "A1:C2"}` | merges |
| `{"rows": "3:5", "hidden": true}`, `{"rows": "3:8", "group": true}`, `{"columns": "B:D", "collapse": true}`, `"level": 2`, `"reset": true` | hide, group, collapse (`false` ungroups, expands), outline level, back to automatic |
| `{"panes": "split", "x": 120, "y": 40}`, `{"panes": null}` | a split view (points); `freeze` as above |
| `{"define_name": "Rate", "refers_to": "Budget!$B$15", "scope": "Budget"}`, `{"remove_name": "Rate"}`, `{"rename_name": "Rate", "to": "Tax"}` | defined names (a workbook name without `scope`); formulas follow a rename |
| `{"add_table": "Costs", "range": "A3:H11", "style": "TableStyleMedium2", "totals_row": true}`, `{"resize_table": "Costs", "range": "A3:H20"}`, `{"rename_table": "Costs", "to": "Spend"}`, `{"rename_table_column": "Costs", "from": "Qty", "to": "Units"}`, `{"table_style": "Costs", "style": null}`, `{"set_totals_row": "Costs", "on": false}`, `{"remove_table": "Costs"}` | tables: the first row of the range holds the column names (plain text, unique); the totals row goes under the range (`Total`, then a `SUBTOTAL`), so that row must be free; `remove_table` keeps the cells |
| `{"validation": "H4:H11", "type": "list", "value": "Open,Done"}`, `{"validation": "C4:C11", "type": "whole", "between": [0, 500], "error_message": "..."}`, `{"remove_validation": "C4:C11"}` | data validation; `type`: `whole`, `decimal`, `list` (`"a,b,c"` or `"=$D$1:$D$5"`), `date`, `time`, `textLength`, `custom`; `operator`, `value`, `value2`, `allow_blank`, `dropdown`, `error_style`, `input_title`, `input_message`, `error_title`, `error_message` |
| `{"conditional_format": "E4:E11", "rule": "cellIs", "operator": "greaterThan", "value": "10000", "style": {"bold": true, "fill": "FFC7CE"}}`, `{"remove_conditional_format": "E4:E11", "priority": 2}` | `rule`: `cellIs`, `expression` (`formula`), `containsText`, `notContainsText`, `beginsWith`, `endsWith` (`text`), `duplicateValues`, `uniqueValues`, `top10` (`rank`, `percent`, `bottom`), `aboveAverage` (`below`), `colorScale` (`colors`), `dataBar` (`color`); a style has no font name, size or alignment |
| `{"comment": "B2", "text": "Check", "author": "Ann"}` (a note; `"thread": true`, `"at": "2026-11-23 09:10"` for a thread), `{"reply": "B2", "text": "...", "author": "Bob"}`, `{"edit_comment": "B2", "text": "..."}`, `{"resolve": "B2", "resolved": false}`, `{"remove_comment": "B2"}`, `{"remove_reply": "B2", "index": 0}` | notes and threads |
| `{"autofilter": "A1:F20"}`, `{"filter": "C", "gt": 100, "and_lt": 500}`, `{"filter": "C", "values": ["a", "b"]}`, `{"filter": "C", "clear": true}`, `{"filter": null}` | filters (`eq`, `ne`, `gt`, `ge`, `lt`, `le`, `begins`, `ends`, `contains`, `not_contains`, the second with `and_` or `or_`; `blanks_only`, `non_blanks`); `"table"` for a table's |
| `{"sort": "A1:F20", "by": [{"column": "C", "descending": true}, {"header": "Name"}], "header": true}`, `{"sort": true, "table": "Costs", "by": [...], "header": true}` | sort; `"filter": true` for the sheet's filter |

## Calculating

`oxoffice xlsx calc F [--all] [--now T] [--exit-code] [--json] [-o OUT]` recomputes every formula (`--all` also lists
the volatile ones: `TODAY`, `NOW`, `RAND`) and prints how many were checked, differ and were skipped, each skipped one
with its reason and the circular references; with `-o` it writes the computed values. `--now 2026-11-28T10:00:00` sets
the clock. `--json`: `checked`, `computed`, `differences`, `skipped`, `cycles`.

The 84 supported functions: **math** `ABS CEILING FLOOR INT MOD POWER PRODUCT RAND ROUND ROUNDDOWN ROUNDUP SQRT
TRUNC`; **aggregates** `SUM SUMIF SUMIFS SUMPRODUCT SUBTOTAL AVERAGE AVERAGEIF COUNT COUNTA COUNTBLANK COUNTIF
COUNTIFS MAX MIN MEDIAN`; **logic** `IF IFS IFERROR IFNA AND OR NOT SWITCH`; **text** `CONCAT CONCATENATE TEXTJOIN
LEFT RIGHT MID LEN LOWER UPPER TRIM SUBSTITUTE FIND SEARCH TEXT VALUE`; **lookup** `VLOOKUP HLOOKUP XLOOKUP INDEX
MATCH CHOOSE ROW COLUMN OFFSET INDIRECT`; **dates** `DATE DAY MONTH YEAR TODAY NOW EDATE EOMONTH WEEKDAY`;
**information** `ISBLANK ISERROR ISNA ISNUMBER ISTEXT ISLOGICAL NA`; **dynamic arrays** `SORT SORTBY FILTER UNIQUE
SEQUENCE TRANSPOSE LET`. Any other (`STDEV`, `RANK`, `LARGE`, financial and engineering functions, `LAMBDA`,
`XMATCH`, `NETWORKDAYS`, `DATEDIF`, `TEXTSPLIT`, `GROUPBY`) is skipped: it keeps its old value, and so does every
formula that reads it.

## Sheets, rows, columns

```bash
oxoffice xlsx sheet F add NAME --at 2 -o OUT              # at a position from 1; after the others without --at
oxoffice xlsx sheet F rename OLD NEW -o OUT               # also remove NAME, show NAME
oxoffice xlsx sheet F move NAME --to 3 -o OUT
oxoffice xlsx sheet F hide NAME --very -o OUT             # --very: not even listed by Excel's Unhide
oxoffice xlsx insert-rows F 5:6 [--sheet S] -o OUT        # also delete-rows
oxoffice xlsx insert-columns F C:D [--sheet S] -o OUT     # also delete-columns
oxoffice xlsx merge F A1:D1 [--across] [--sheet S] -o OUT     # also unmerge F RANGE
oxoffice xlsx rows F 3:8 group -o OUT                         # also hide show ungroup collapse expand reset level N height H autofit; columns: width W
oxoffice xlsx columns F A:C autofit [--min W] [--max W] [--rows 2:100] -o OUT
oxoffice xlsx outline F --summary-below false --summary-right false -o OUT
oxoffice xlsx panes F freeze B2 -o OUT                        # also split X Y (points), remove
```

Everything that holds a reference follows a sheet, row or column change: formulas (shared and array ones included),
names, tables, merges, validations, conditional formats, links, notes, the filter, panes, breaks, charts (a removed
sheet's references become `#REF!`). A change that cannot be made whole (the last sheet, rows through an array
formula, a table's header cut, content pushed off the sheet) is refused and writes nothing. New rows take the
formats of the row above. A removed sheet leaves the cached values of the formulas that read it as they were:
`calc F -o OUT` writes them. `merge` keeps the top-left value and lists on stderr the cells that lost one. `autofit`
measures the text with the render's fonts (a font it lacks gives an estimate) and prints the widths in characters
of the Mac's 6-pixel digit; `grid` reads the 7-pixel one.

## Filter, sort, notes

```bash
oxoffice xlsx filter F set A1:F20 -o OUT                       # the first row is the header; --table T for a table's
oxoffice xlsx filter F show                                    # the range, each column's criterion, the rows it hides
oxoffice xlsx filter F clear C -o OUT                          # without COL: every column; also reapply, remove
oxoffice xlsx filter F where C --gt 100 --and --lt 500 -o OUT  # --header TEXT instead of C; --values a,b [--blanks]; --equals --not-equals --ge --le --begins --ends --contains --or
oxoffice xlsx sort F A1:F20 --header --by C:desc --by name=Region -o OUT     # or --table T, --filter instead of a range; --case-sensitive
oxoffice xlsx comment F add B2 "Checked" --author Ann -o OUT   # also thread, reply, edit, resolve, reopen, remove; --at "2026-11-23 09:10"
```

A filter hides the rows that do not pass in the same call; editing a cell does not change them until `filter F
reapply`. A criterion or a sort key over formulas is refused ("holds a formula whose value is not current: calculate the
workbook first") when their values are skipped or when the file asks for a recalculation on open, as every file
oxoffice saved does (`calc -o` does not clear that): filter and sort a file from Excel as it is, key on a column of
plain values, or chain the steps in one Python session, where the values are current. A sort moves the cells of a
row with their formulas written as if copied, notes and styles; a range is sorted exactly as given. `comment add` is
a note, `thread` a threaded comment (dated by the clock unless `--at`), `edit` changes a note's text or a thread's
first comment, `remove` a note or a whole thread.

## Rendering

```bash
oxoffice xlsx convert F -t pdf [-o OUT.pdf] [--sheet S]... [--date 2026-11-28] [--font-dir D] [--bundled-fonts-only] [--system-fonts]
oxoffice xlsx convert F -t png [-o OUT.png] [--pages 1,3-5] [--dpi 150] [--transparent]      # also jpeg, tiff
oxoffice xlsx render F -o DIR [--pages 1,3-5] [--format png] [--dpi 150]                      # one file per page: DIR/NAME_page1.png
```

Every visible sheet prints, one after the other, as Excel's own export does: its print area, else its used range, on
its own paper and margins, at 100 % with automatic page breaks. A multi-page image writes `OUT_001.png`,
`OUT_002.png`. `--date` is what `&D` and `&T` of a header or footer print (never the clock). Diagnostics (an
approximated feature, a formula drawn with a stale value) go to stderr. The render is oxoffice's, not Excel's: look
at the PNG.

## Known limits

`oxoffice limitations xlsx` lists them with their workarounds (`oxoffice limitations --json`: the same as data, under
`.limitations`). Today: a spilling formula written without its range shows one value in Google Sheets when its size
is not obvious from the formula (give the range), and the render does not shape right-to-left text of a cell.
