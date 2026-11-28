# oxoffice.xlsx Python API

`from oxoffice.xlsx import Workbook, Formula` with the oxoffice Python (`oxpy`). Exact signatures, with their
docstrings, are in `oxoffice/xlsx/_native.pyi` in the installed package (`oxpy -c "import oxoffice.xlsx, os;
print(os.path.dirname(oxoffice.xlsx.__file__))"` prints its folder). The names are oxoffice's own, not openpyxl's:
openpyxl's `load_workbook`, `append`, `iter_rows`, `merge_cells`, `ws["A1"]`, `cell(row, column)` and
`cell.number_format = ...` (a cell is read only) do not exist here; an unknown name raises a plain `AttributeError`: search the stubs
for the oxoffice way. The module exports `Workbook`, `Worksheet`, `Cell`, `Formula`, `Comment`, `Reply`, `Style`
(with `Font`, `Fill`, `Border`, `BorderSide`, `Alignment`, `Protection`, `Color`, `GradientStop`), `PageSetup`
(with `PageMargins`, `HeaderFooter`, `PageBreak`) and the errors. Every index is 0-based (`wb.sheets[0]`,
`ws.index`, `(row, column)`); A1 references are the spreadsheet's own, from 1.

## Workbook

| Member | Notes |
|---|---|
| `Workbook()` (one empty sheet, `Sheet1`), `Workbook(path)`, `Workbook(sheets=["Data", "Summary"])`, `Workbook.open(path)`, `Workbook.from_bytes(b)` | a new workbook is Calibri 11 under the Office 2013-2022 theme, the same bytes each time; `ValueError` for a sheet name Excel rejects or no name, `XlsxError` for a repeated name (any case) |
| `Workbook.open_encrypted(path, password)`, `from_encrypted_bytes(b, password)`, `save_encrypted(path, password)`, `to_encrypted_bytes(password)` | password-protected files; the plain `Workbook(path)` raises `PackageError` for one |
| `wb.sheets` (a tuple, tab order), `wb.sheet(name)`, `wb.add_sheet(name)` → the new `Worksheet` | `sheet` takes the exact name, then any case, `KeyError` when none; a name is 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History` |
| `wb.date1904`, `wb.recalculation_needed` | the latter is true when the file asks for a recalculation on open, or an edit changed a workbook that holds a formula: the cached values may be out of date |
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
| `ws.merged_ranges`, `ws.comments`, `ws.page_setup` | read only; `comments` are `Comment`: `ref`, `kind` (`note`, `thread`), `author`, `date` (threads), `text`, `resolved`, `replies` (`Reply`: `author`, `date`, `text`) |
| `ws.to_csv(*, raw=False, delimiter=",", bom=False)` → str | from A1 to the end of the used range, CRLF line ends; displayed texts, or `raw` stored values |
| `ws.set_value(cell, value)` | number, `str` (never a formula; `""` clears), `bool`, `CellError`, `None` (clears), `datetime.datetime` / `date` / `time` / `timedelta` (a serial; a General cell takes `yyyy-mm-dd`, `yyyy-mm-dd hh:mm:ss`, `hh:mm:ss` or `[h]:mm:ss`), or a `Formula`; keeps the cell's style. `XlsxError` for a cell of an array formula written alone, a value under a merged range but its first cell, a table header that would stop naming its column, a pivot cell; `ValueError` for a value a cell cannot hold; `TypeError` for another type. A refusal changes nothing |
| `ws.set_formula(cell, formula, *, stored=False, spill=False)` | the text as Excel displays it after `=` (an `=` is accepted); `stored=True` takes `_xlfn.XLOOKUP(...)`; `spill=True` or `spill="J2:J5"` writes a spilling formula (give the range: Google Sheets takes it from the file); `ValueError` for a formula Excel would repair (syntax and argument counts are checked, function and sheet names are not; quote a sheet name with a space, `'Q1 notes'!B2`). No cached value: the cell reads `None` and `stale` until Excel calculates |
| `Formula(text, stored=False, spill=False)` | a formula as a value, for `write` and `set_value` |
| `ws.write(cell, rows)` | rows of values from `cell` down and right, each as `set_value` writes it; a shorter row leaves the cells after it as they are |
| `ws.import_csv(source, *, at="A1", infer=True)` | `source` is the CSV text, its UTF-8 bytes or the path of its file (a `str` that is not a path is read as text: pass a `Path` for a file); the rectangle is written, an empty field clearing its cell; `infer` types a field as number (no leading zero, 15 digits at most), `TRUE` / `FALSE`, date, else `str`; `infer=False` keeps every field a `str` |
| `ws.clear(range, *, contents=True, formats=False)` | `"A1:C10"`, `"B3"`, `"A:C"`, `"2:5"`; a whole array formula is removed, part of one raises `XlsxError` |
| `ws.set_style(range, *, bold, italic, underline, strike, font_name, font_size, font_color, fill, border, border_color, horizontal, vertical, wrap_text, number_format)` | only the arguments given change; colours `"RRGGBB"`; `fill` is solid, `border` the style of the four sides (`"thin"`, `"medium"`, `"thick"`, `"double"`, `"none"`); `horizontal` `"left"`, `"center"`, `"right"`...; a whole column or row range styles the column or row itself, so a cell typed later takes it |
| `ws.set_number_format(range, code)` | `"0.00"`, `"yyyy-mm-dd"`, `"#,##0.00"`, `"0%"` |
| `ws.set_column_width(columns, width)`, `ws.set_row_height(rows, height)` | `"B"` or `"A:C"` in characters (above 0, at most 255); `"3"` or `"2:5"` in points (at most 409.5); `XlsxError` for a sheet that holds no cells |
| `ws.freeze_panes(cell)` | `"B2"` keeps row 1 and column A in view; `None` removes them (the argument is required); `"A1"` raises `ValueError` |

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
