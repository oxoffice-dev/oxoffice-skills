# xlsx examples

Every block runs on each push to oxoffice's `main`, against the wheel and the CLI just built. Conventions, as set
up in SKILL.md: `ox` runs the CLI and Python blocks run with `oxpy`. `budget.xlsx` is the input (a workbook with
formulas, dates, a text code with a leading zero, two notes and a sheet whose name holds a space); outputs go to new
files. Blocks run in order in one folder.

## Read the workbook

```bash
ox xlsx inspect budget.xlsx
ox xlsx text --markdown --sheet Budget budget.xlsx                # every value at its address
ox xlsx cells budget.xlsx E4:E5 --json > cells.json               # value, formula, number format, style, stale
ox xlsx cells budget.xlsx "'Q1 notes'!B2"                         # a sheet name with a space, quoted
ox xlsx comments budget.xlsx
ox xlsx csv budget.xlsx --sheet Budget -o budget.csv              # the displayed texts
ox xlsx text --json budget.xlsx > budget.json
```

```python
import datetime, json
from oxoffice.xlsx import Workbook
wb = Workbook("budget.xlsx")
print([(s.name, s.used_range) for s in wb.sheets])
ws = wb.sheet("Budget")
rows = ws.values("A4:H11")                       # one native call; numbers are floats, dates serial days
assert rows[0][:3] == ["Survey of the river bank", "Field", 2.0] and rows[0][6] == "007"
total = ws.cell("E12")
print(total.formula, total.value, total.text)    # SUM(E4:E11) 61647.0 61,647.00
assert total.formula == "SUM(E4:E11)" and total.value == 61647.0
due = ws.cell("F4")
assert due.value == 46059.0 and due.datetime == datetime.datetime(2026, 2, 6) and due.text == "2026-02-06"
assert [(c.ref, c.author) for c in ws.comments] == [("D6", "Ada"), ("E12", "Grace")]
assert ws.merged_ranges == ["A1:H1"]
cells = json.load(open("cells.json", encoding="utf-8"))["cells"]
assert [c["formula"] for c in cells] == ["C4*D4", "C5*D5"]
assert open("budget.csv", encoding="utf-8").read().splitlines()[2].startswith("Item,Team,Qty")
```

## Change cells from the CLI, one at a time or many in one call

```bash
ox xlsx set budget.xlsx C4 3 --sheet Budget -o qty.xlsx          # the number 3
ox xlsx set budget.xlsx G4 008 --text --sheet Budget -o code.xlsx   # 008 stays text
cat > edits.json <<'JSON'
[
  {"sheet": "Budget", "cell": "C4", "value": 3},
  {"cell": "H4", "value": "Invoiced"},
  {"cell": "I3", "value": "Late?"},
  {"cell": "I4", "formula": "IF(F4<DATE(2026,3,1),\"yes\",\"no\")"},
  {"cell": "F12", "date": "2026-07-01"},
  {"range": "A3:I3", "style": {"bold": true, "fill": "FFFF00", "border": "thin"}},
  {"range": "D4:E12", "number_format": "#,##0.00"},
  {"range": "H5:H6", "clear": "contents"},
  {"columns": "I", "width": 10},
  {"freeze": "A4"},
  {"add_sheet": "Summary"},
  {"cell": "A1", "value": "Total"},
  {"cell": "B1", "formula": "SUM(Budget!E4:E11)"}
]
JSON
ox xlsx edit budget.xlsx --map edits.json -o edited.xlsx
```

```python
from oxoffice.xlsx import Workbook
import json
qty = Workbook("qty.xlsx").sheet("Budget")
assert qty.cell("C4").value == 3.0 and qty.cell("G4").value == "007"
code = Workbook("code.xlsx").sheet("Budget").cell("G4")
assert code.value == "008" and isinstance(code.value, str)
wb = Workbook("edited.xlsx")
ws = wb.sheet("Budget")
assert ws.cell("C4").value == 3.0 and ws.cell("H4").value == "Invoiced" and ws.cell("H5").value is None
assert ws.cell("I4").formula == 'IF(F4<DATE(2026,3,1),"yes","no")'
assert ws.cell("I4").value is None and ws.cell("I4").stale          # nothing is calculated
assert ws.cell("E4").value == 8400.0 and ws.cell("E4").stale        # 3 x 4,200 is 12,600 once Excel calculates
assert wb.recalculation_needed
assert ws.cell("A3").style.font.bold and ws.cell("A3").style.fill.foreground.rgb == "FFFF00"
assert wb.sheet("Summary").cell("B1").formula == "SUM(Budget!E4:E11)"
assert [s.name for s in wb.sheets] == ["Budget", "Q1 notes", "Summary"]
print("E4 as the sheet holds it:", ws.cell("E4").text, "- to report, work it out:", 3 * 4200)
```

## Write from Python: values, dates, formulas, blocks, styles

```python
import datetime
from oxoffice.xlsx import Formula, Workbook

wb = Workbook("budget.xlsx")
ws = wb.sheet("Budget")
ws.set_value("C4", 3)                              # a number; "3" would be text
ws.set_value("G5", "013")                          # a str stays text
ws.set_value("F5", datetime.date(2026, 2, 27))     # shown yyyy-mm-dd in a General cell
ws.set_value("H4", None)                           # clears the content
ws.set_formula("I4", "IF(F4<DATE(2026,3,1),\"yes\",\"no\")")
ws.set_formula("J4", "XLOOKUP(A4,A4:A11,E4:E11)")  # stored with _xlfn., as Excel stores it
ws.write("A14", [["Contingency 10%", None, None, None, Formula("E12*0.1")],
                 ["Grand total", None, None, None, Formula("E12+E14")]])
ws.set_style("A14:E15", bold=True, border="thin", fill="E2EFDA")
ws.set_style("A3:H3", horizontal="center", wrap_text=True, font_color="1F3864")
ws.set_number_format("E14:E15", "#,##0.00")
ws.set_column_width("A:B", 28)
ws.set_row_height("3", 30)
ws.clear("A15:E15", formats=True)                  # contents and formats
wb.save("written.xlsx")                            # drop_cached_values=True for LibreOffice

check = Workbook("written.xlsx").sheet("Budget")
assert check.cell("G5").value == "013" and check.cell("F5").datetime == datetime.datetime(2026, 2, 27)
assert check.cell("E14").formula == "E12*0.1" and check.cell("E14").style.font.bold
assert check.cell("J4").stored_formula == "_xlfn.XLOOKUP(A4,A4:A11,E4:E11)"
assert check.cell("E15").formula is None and not check.cell("E15").style.font.bold
try:
    wb.save("budget.xlsx")                         # the file it was opened from
except Exception as error:
    print(type(error).__name__, error)
else:
    raise SystemExit("the input was overwritten")
try:
    ws.set_formula("K4", "IF(B2>0;1;0)")           # a ";" separator: Excel would repair the file
except ValueError as error:
    print("refused:", error)
```

## A CSV into a sheet, a sheet out as CSV

```bash
printf 'item,qty,due\nnorth,12,2026-02-01\nsouth,007,2026-02-08\n' > extra.csv
ox xlsx import-csv budget.xlsx extra.csv --sheet "Q1 notes" --at A6 -o merged.xlsx
ox xlsx csv merged.xlsx --sheet "Q1 notes" --raw -o notes-raw.csv        # stored values, dates as serials
ox xlsx csv merged.xlsx --sheet Budget --delimiter tab -o budget.tsv
```

```python
from oxoffice.xlsx import Workbook
notes = Workbook("merged.xlsx").sheet("Q1 notes")
assert notes.values("A6:C8") == [["item", "qty", "due"], ["north", 12.0, 46054.0], ["south", "007", 46061.0]]
assert notes.cell("C7").text == "2026-02-01"      # the CSV date is a date, shown yyyy-mm-dd
text_only = Workbook("budget.xlsx")
text_only.sheet("Q1 notes").import_csv("a,b\n1,2\n", at="D1", infer=False)   # every field a str
assert text_only.sheet("Q1 notes").values("D1:E2") == [["a", "b"], ["1", "2"]]
assert "south,007" in open("notes-raw.csv", encoding="utf-8").read()
```

## A new workbook

```bash
ox xlsx new --sheet Data --sheet Summary -o book.xlsx
cat > grid.json <<'JSON'
[
  {"sheet": "Data", "cell": "A1", "value": "Region"}, {"cell": "B1", "value": "Sales"},
  {"cell": "A2", "value": "North"}, {"cell": "B2", "value": 1200.5},
  {"cell": "A3", "value": "South"}, {"cell": "B3", "value": 980},
  {"range": "A1:B1", "style": {"bold": true}}, {"range": "B2:B3", "number_format": "#,##0.00"},
  {"columns": "A:B", "width": 16}, {"rows": "1", "height": 24}, {"freeze": "A2"},
  {"sheet": "Summary", "cell": "A1", "value": "Total"}, {"cell": "B1", "formula": "SUM(Data!B2:B3)"}
]
JSON
ox xlsx edit book.xlsx --map grid.json -o formatted.xlsx
```

```python
from oxoffice.xlsx import Workbook
wb = Workbook(sheets=["Data", "Notes"])            # or Workbook() for one empty "Sheet1"
data = wb.sheet("Data")
data.write("A1", [["Region", "Sales"], ["North", 1200.5], ["South", 980]])
data.set_style("A1:B1", bold=True)
data.set_column_width("A:B", 16)
data.freeze_panes("A2")                            # the header row stays in view
extra = wb.add_sheet("Summary")
extra.set_formula("B1", "SUM(Data!B2:B3)")
wb.save("built.xlsx")
formatted = Workbook("formatted.xlsx")
assert [s.name for s in formatted.sheets] == ["Data", "Summary"]
assert formatted.sheet("Data").values("A1:B3")[1] == ["North", 1200.5]
assert formatted.sheet("Summary").cell("B1").formula == "SUM(Data!B2:B3)"
assert Workbook("built.xlsx").sheet("Data").cell("A1").style.font.bold
```

## Spilling formulas, and a file for LibreOffice

```bash
ox xlsx set built.xlsx D1 "SORT(A2:A3)" --formula --spill-range D1:D2 --sheet Data -o spilled.xlsx
ox xlsx set built.xlsx C1 "B2*2" --formula --sheet Data -o libre.xlsx --drop-cached-values
```

```python
from oxoffice.xlsx import Workbook
spilled = Workbook("spilled.xlsx").sheet("Data").cell("D1")
assert spilled.formula == "SORT(A2:A3)" and spilled.formula_kind == "array" and spilled.array_range == "D1:D2"
wb = Workbook("libre.xlsx")
assert wb.recalculation_needed and wb.sheet("Data").cell("C1").value is None
assert len(wb.to_bytes(drop_cached_values=True)) > 0
```
