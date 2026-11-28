# oxoffice.xlsx Python API

`from oxoffice.xlsx import Workbook, Formula` with the oxoffice Python (`oxpy`). Exact signatures, with their
docstrings, are in `oxoffice/xlsx/_native.pyi` in the installed package (`oxpy -c "import oxoffice.xlsx, os;
print(os.path.dirname(oxoffice.xlsx.__file__))"` prints its folder). The names are oxoffice's own, not openpyxl's:
openpyxl's `load_workbook`, `append`, `iter_rows`, `merge_cells`, `ws["A1"]`, `cell(row, column)` and
`cell.number_format = ...` (a cell is read only) do not exist here; an unknown name raises a plain `AttributeError`: search the stubs
for the oxoffice way. The module exports `Workbook`, `Worksheet`, `Cell`, `Formula`, `Comment`, `Reply`, `Recalculation`, `Table`, `DefinedName`, `Validation`,
`ConditionalFormat`, the filter criteria `Values`, `Compare`, `Blanks`, `NonBlanks`, `diff` (with `Diff` and its changes), `Style`
(with `Font`, `Fill`, `Border`, `BorderSide`, `Alignment`, `Protection`, `Color`, `GradientStop`), `PageSetup`
(with `PageMargins`, `HeaderFooter`, `PageBreak`) and the errors. Every index is 0-based (`wb.sheets[0]`,
`ws.index`, `(row, column)`); A1 references are the spreadsheet's own, from 1.

## Workbook

| Member | Notes |
|---|---|
| `Workbook()` (one empty sheet, `Sheet1`), `Workbook(path)`, `Workbook(sheets=["Data", "Summary"])`, `Workbook.open(path)`, `Workbook.from_bytes(b)` | a new workbook is Calibri 11 under the Office 2013-2022 theme, the same bytes each time; `ValueError` for a sheet name Excel rejects or no name, `XlsxError` for a repeated name (any case) |
| `Workbook.open_encrypted(path, password)`, `from_encrypted_bytes(b, password)`, `save_encrypted(path, password)`, `to_encrypted_bytes(password)` | password-protected files; the plain `Workbook(path)` raises `PackageError` for one |
| `wb.sheets` (a tuple, tab order), `wb.sheet(name)`, `wb.add_sheet(name)` → the new `Worksheet` | `sheet` takes the exact name, then any case, `KeyError` when none; a name is 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History` |
| `wb.date1904`, `wb.recalculation_needed` | the latter is true when the file asks for a recalculation on open, or an edit changed a workbook that holds a formula: it stays true after oxoffice calculated the edit (Excel and Google Sheets recalculate on open too), `Cell.stale` likewise |
| `wb.calculation`, `wb.recalculate(*, everything=False, now=None)` | `"automatic"` (default: every edit call calculates what it changed and what reads it) or `"manual"` (edits keep the cached values until `recalculate()`; setting `"automatic"` calculates them). `recalculate()` returns a `Recalculation`: `computed`, `changed`, `skipped` (`[("Budget!J4", "STDEV is not computed yet")]`: the formulas that keep their old value, with the ones that read them), `cycles`. `everything=True` recomputes every formula, `now` is what `TODAY` and `NOW` read |
| `wb.rename_sheet(old, new)`, `wb.remove_sheet(name)`, `wb.move_sheet(name, to)`, `wb.set_sheet_state(name, "visible"\|"hidden"\|"veryHidden")`, `wb.add_sheet(name, at=)` | every reference follows (formulas, names, conditional formats, validations, charts, links); removing leaves `#REF!`; `XlsxError` for the last sheet or what cannot be made whole |
| `wb.names`, `wb.define_name(name, refers_to, *, scope=None, hidden=False, comment=None)`, `wb.remove_name(name, *, scope=None)`, `wb.rename_name(old, new, *, scope=None)` | `DefinedName`: `name`, `scope`, `refers_to`, `hidden`, `comment`; `refers_to` is a formula as displayed (`"Budget!$B$15"`, `"0.2"`); `ValueError` for a name Excel refuses (`A1`, `R1C1`, spaces) |
| `wb.to_pdf(*, sheets=None, fonts=None, font_dir=None, bundled_fonts_only=False, system_fonts=False, date=None)` → bytes, `wb.page_count()`, `wb.render_page_to_png(page, dpi=150.0)` → bytes or `None` past the last page, `wb.render_pages([0, 1], dpi)`, `wb.render_all_pages(dpi)`, `to_pdf_with_diagnostics` | pages count from 0; every visible sheet, or `sheets=["Budget", 2]` (name or number from 1); `date` is what `&D` and `&T` print; `XlsxError` or `ValueError` for no sheet to print or an unknown one |
| `wb.to_markdown()` | `## NAME` and a table per sheet, hidden ones included, each value at its address |
| `wb.save(path, *, overwrite_source=False, drop_cached_values=False)`, `wb.to_bytes(*, drop_cached_values=False)` | staged and renamed: a failed save changes nothing; `PackageError` for the file the workbook was opened from unless `overwrite_source=True`; the extension (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) selects the class and any other name still writes a workbook; `drop_cached_values` empties every formula's cached value so LibreOffice computes them |

A file that is not a workbook (`.xls`, `.xlsb`, `.ods`, Word, PowerPoint) raises `PackageError` with the way out.

## Worksheet

| Member | Notes |
|---|---|
| `ws.name`, `ws.index`, `ws.state` (`visible`, `hidden`, `veryHidden`), `ws.kind` (`worksheet`, `chartsheet`, `dialogsheet`, `macrosheet`) | a chart, dialog or macro sheet has no used range, merges or comments, and a read of its cells raises `XlsxError` |
| `ws.used_range` | `"A1:H12"`, `None` when empty; read once |
| `ws.cell("B3")` or `ws.cell((2, 1))` → `Cell` | a snapshot: an edit does not change it, read it again; a cell the sheet does not write reads as empty; `ValueError` for a reference that names no cell, `IndexError` outside the sheet |
| `ws.values(range, text=False)` → list of rows | one native call, the fast way to read a block: floats (dates as serial days), `str`, `bool`, `CellError`, `None` for empty; `text=True` the displayed texts, `""` for empty |
| `ws.cells(range)` → `[Cell]` | the cells of the range that hold a value or a formula, row by row |
| `ws.merged_ranges`, `ws.comments`, `ws.page_setup`, `ws.tables`, `ws.validations`, `ws.conditional_formats`, `ws.hidden_rows`, `ws.column_dimensions`, `ws.row_dimensions` | read only; `ws.column_dimension("B")` and `row_dimension("3")` read one (`width`, `height`, `hidden`, `outline_level`, `hidden_by`: `filter` or `manual`); `comments` are `Comment`: `ref`, `kind` (`note`, `thread`), `author`, `date` (threads), `text`, `resolved`, `replies` (`Reply`: `author`, `date`, `text`) |
| `ws.to_csv(*, raw=False, delimiter=",", bom=False)` → str | from A1 to the end of the used range, CRLF line ends; displayed texts, or `raw` stored values |
| `ws.set_value(cell, value)` | number, `str` (never a formula; `""` clears), `bool`, `CellError`, `None` (clears), `datetime.datetime` / `date` / `time` / `timedelta` (a serial; a General cell takes `yyyy-mm-dd`, `yyyy-mm-dd hh:mm:ss`, `hh:mm:ss` or `[h]:mm:ss`), or a `Formula`; keeps the cell's style. `XlsxError` for a cell of an array formula written alone, a value under a merged range but its first cell, a table header that would stop naming its column, a pivot cell; `ValueError` for a value a cell cannot hold; `TypeError` for another type. A refusal changes nothing |
| `ws.set_formula(cell, formula, *, stored=False, spill=False)` | the text as Excel displays it after `=` (an `=` is accepted); `stored=True` takes `_xlfn.XLOOKUP(...)`; `spill=True` or `spill="J2:J5"` writes a spilling formula (give the range when the formula does not show the size: Google Sheets takes it from the file); `ValueError` for a formula Excel would repair (syntax and argument counts are checked, function and sheet names are not; quote a sheet name with a space, `'Q1 notes'!B2`). The formula is calculated at once when its functions are supported (`wb.recalculate().skipped` says otherwise: the cell then reads `None` or its old value) |
| `Formula(text, stored=False, spill=False)` | a formula as a value, for `write` and `set_value` |
| `ws.write(cell, rows)` | rows of values from `cell` down and right, each as `set_value` writes it; a shorter row leaves the cells after it as they are |
| `ws.import_csv(source, *, at="A1", infer=True)` | `source` is the CSV text, its UTF-8 bytes or the path of its file (a `str` that is not a path is read as text: pass a `Path` for a file); the rectangle is written, an empty field clearing its cell; `infer` types a field as number (no leading zero, 15 digits at most), `TRUE` / `FALSE`, date, else `str`; `infer=False` keeps every field a `str` |
| `ws.clear(range, *, contents=True, formats=False)` | `"A1:C10"`, `"B3"`, `"A:C"`, `"2:5"`; a whole array formula is removed, part of one raises `XlsxError` |
| `ws.set_style(range, *, bold, italic, underline, strike, font_name, font_size, font_color, fill, border, border_color, horizontal, vertical, wrap_text, number_format)` | only the arguments given change; colours `"RRGGBB"`; `fill` is solid, `border` the style of the four sides (`"thin"`, `"medium"`, `"thick"`, `"double"`, `"none"`); `horizontal` `"left"`, `"center"`, `"right"`...; a whole column or row range styles the column or row itself, so a cell typed later takes it |
| `ws.set_number_format(range, code)` | `"0.00"`, `"yyyy-mm-dd"`, `"#,##0.00"`, `"0%"` |
| `ws.set_column_width(columns, width)`, `ws.set_row_height(rows, height)` | `"B"` or `"A:C"` in characters (above 0, at most 255); `"3"` or `"2:5"` in points (at most 409.5); `XlsxError` for a sheet that holds no cells |
| `ws.rename(new)`, `ws.remove()`, `ws.move(to)`, `ws.hide(very=False)`, `ws.show()` | the sheet itself, as the `wb` methods |
| `ws.insert_rows("5:6")`, `delete_rows("5")`, `insert_columns("C:D")`, `delete_columns("C")` | every reference, table, merge, validation, conditional format, note and filter follows; the new rows take the formats of the row above; `XlsxError`, nothing changed, for rows through an array formula, a table header, content pushed off the sheet |
| `ws.merge(range, *, across=False)` → the cells that lost a value, `ws.unmerge(range)` | the top-left cell keeps its content |
| `ws.hide_rows("3:5")`, `show_rows`, `hide_columns`, `show_columns`, `group_rows`, `ungroup_rows`, `group_columns`, `collapse_rows`, `expand_rows`, `set_outline_level("3:8", 2)`, `set_outline_summary(below=True, right=True)` | rows and columns, as `"3:5"` / `"B:D"` |
| `ws.reset_column_width(cols)`, `ws.reset_row_height(rows)`, `ws.autofit_columns("A:C", *, min=None, max=None, rows=None)` → `{letter: width}`, `ws.autofit_rows("2:40")` → `{row: points}`, `ws.default_column_width`, `ws.default_row_height`, `ws.sheet_format` | autofit measures the text with the render's fonts (a font that is not measured gives a `UserWarning`: an estimate); widths in characters of the Mac's 6-pixel digit, `column_dimension` reads the 7-pixel one |
| `ws.split_panes(x, y)`, `ws.remove_panes()`, `ws.panes` | a split in points; `freeze_panes` below |
| `ws.set_autofilter(range)`, `ws.autofilter(table=None)`, `ws.set_filter(column, criterion, *, table=None)` → `FilterOutcome` (`shown`, `hidden`), `ws.clear_filter(column=None)`, `ws.reapply_filter()`, `ws.remove_autofilter()` | `column` is `"C"` or `("header", "Amount")`; `criterion`: `Values(["a", "b"], blanks=False)`, `Compare(">", 100)` (`.and_("<", 500)`, `.or_(...)`; operators `=`, `!=`, `<`, `<=`, `>`, `>=`, `begins`, `ends`, `contains`, `not contains`), `Blanks()`, `NonBlanks()`; the rows that do not pass hide in the call; a column whose formula value is skipped raises `XlsxError` |
| `ws.sort(range=None, *, by, header=False, case_sensitive=False, table=None, filter=False)` → `SortOutcome` (`moved`, `rows`) | `by`: `"C"`, `("C", "desc")`, `("header", "Amount", "desc")` or a list of them, the main key first; `header` is required for a table and a filter; cells move with their formulas written as if copied, notes and styles; numbers, then text, booleans, errors, blanks last; a merge, array formula or pivot table in the block raises `XlsxError` |
| `ws.add_table(name, range, *, style=None, totals_row=False, first_column=False, last_column=False, row_stripes=True, column_stripes=False)`, `ws.resize_table(name, range)`, `rename_table(old, new)`, `rename_table_column(table, old, new)`, `set_table_style(name, style)`, `set_totals_row(name, on=True)`, `remove_table(name)` | `Table`: `name`, `sheet`, `range`, `header_row`, `totals_row`, `style`, `columns` (`name`, `formula`, `total`). The first row is the header (plain text, not empty, unique); `style` `TableStyleMedium2` (default), `TableStyleLight1`-`21`, `Medium1`-`28`, `Dark1`-`11`; a totals row is `Total` then a `SUBTOTAL`, in the row under the range, which must be free; `resize_table` refuses a table with a totals row; `remove_table` keeps the cells |
| `ws.set_validation(range, type, *, operator=None, value=None, value2=None, between=None, allow_blank=True, dropdown=True, error_style="stop", input_title=, input_message=, error_title=, error_message=)`, `ws.remove_validation(range)` | `type`: `any`, `whole`, `decimal`, `list` (`"a,b,c"` or `"=$D$1:$D$5"`), `date`, `time`, `textLength`, `custom`; `between=(0, 500)` or `operator` (`greaterThan`...) with `value`; a rule over exactly the cells of an existing one replaces it, one that only shares some raises `XlsxError` |
| `ws.add_conditional_format(range, rule, *, operator, value, value2, formula, text, rank=10, percent, bottom, below, equal, colors, color, bold, italic, underline, strike, font_color, fill, border, border_color, number_format)`, `ws.remove_conditional_format(range, priority=None)` | `rule`: `cellIs`, `expression` (`formula` relative to the range's top-left cell), `containsText`, `notContainsText`, `beginsWith`, `endsWith` (`text`), `duplicateValues`, `uniqueValues`, `top10`, `aboveAverage`, `colorScale` (`colors`, two or three `RRGGBB`), `dataBar` (`color`); the style parts as `set_style`, without font name, size or alignment |
| `ws.add_note(cell, author, text)`, `ws.add_thread(cell, author, text, at=None)`, `ws.reply(cell, author, text, at=None)`, `ws.edit_comment(cell, text)`, `ws.resolve(cell, resolved=True)`, `ws.remove_comment(cell)`, `ws.remove_reply(cell, index)` | `edit_comment` changes a note's text, or a thread's first comment (the replies stay); `add_note` on a note replaces its text, on a thread is refused; `at` (a `datetime`) dates a thread or reply, else the clock is read: pass it to write the same bytes twice |
| `ws.freeze_panes(cell)` | `"B2"` keeps row 1 and column A in view; `None` removes them (the argument is required); `"A1"` raises `ValueError` |

## Comparing

`from oxoffice.xlsx import diff`; `diff(old, new, *, sheets=None, align="auto")`, each a `Workbook` or a path, returns a
`Diff`: `sheets`, `cells`, `comments`, `merges`, `rows` (tuples of changes: `CellChange` has `sheet`, `ref`, `kind`
(`added`, `removed`, `value`, `formula`), `old`, `new`, `recalculated`, `old_formula`, `new_formula`), `summary`
(counts), `to_json()`; a `Diff` is true when anything changed. Cells match by address, formulas by what they mean
there, values as cached; `align="none"` compares each cell with the cell at its address instead of aligning inserted
rows.

## Cell

All read only. `cell.ref` (`"B3"`), `cell.row`, `cell.column` (0-based), `cell.value` (`float`, `str`, `bool`,
`CellError`, `None`; for a formula what its producer cached), `cell.text` (as Excel displays it, en-US),
`cell.formula` (as displayed, shared members moved to their own cell), `cell.stored_formula`, `cell.formula_kind`
(`shared`, `array`, `data_table`, `None`), `cell.array_range`, `cell.stale` (the cached value may be out of date),
`cell.number_format`, `cell.datetime` (the number as a date and time when its format shows one, else `None`),
`cell.duration` (a `timedelta` when its format shows one), `cell.style`: `style.font` (`name`, `size`, `bold`,
`italic`, `underline`, `strike`, `color`), `style.fill` (`kind`, `pattern`, `foreground`, `background`),
`style.border` (`left`, `right`, `top`, `bottom`, each with `style` and `color`), `style.alignment` (`horizontal`,
`vertical`, `wrap_text`, `indent`...), `style.protection`, `style.number_format`, `style.style_name`; a `Color` has
`rgb` (`"RRGGBB"` or `None`), `source` (`rgb`, `theme`, `indexed`, `auto`), `index`, `tint`.

## Errors

`oxoffice.xlsx.XlsxError` is the base; `PackageError` (not a workbook this package opens, or a save that cannot be
made), `XmlError` (also a `ValueError`: a value holding a character XML 1.0 cannot carry). `CellError` is a `str`
holding an error code (`#N/A`, `#DIV/0!`): `cell.value == "#N/A"` works, `isinstance(cell.value, CellError)` tells
an error from a text. Beside them: `ValueError` (a reference, a width, a formula or a value that is wrong),
`KeyError` (no such sheet), `TypeError`, `IndexError`.
