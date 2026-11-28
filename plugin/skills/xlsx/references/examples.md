# xlsx examples

Every block runs on each push to oxoffice's `main`, against the wheel and the CLI just built. Conventions, as set
up in SKILL.md: `ox` runs the CLI and Python blocks run with `oxpy`. `budget.xlsx` is the input (a workbook as Excel saves it, with
fresh cached values for its formulas, dates, a text code with a leading zero, two notes and a sheet whose name holds
a space); outputs go to new files. Blocks run in order in one folder. The sections: read; change cells; write from
Python; CSV; a new workbook; spilling formulas; values for another reader; calculated values; sheets, rows and
columns; names, tables, validation, conditional formats; notes and threads; merge, group, filter, sort, autofit;
compare two workbooks; render.

## Read the workbook

```bash
ox xlsx inspect budget.xlsx | tee inspect-budget.txt          # no "Recalculation" line: the cached values are fresh
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
assert not wb.recalculation_needed and "Recalculation" not in open("inspect-budget.txt", encoding="utf-8").read()
ws = wb.sheet("Budget")
rows = ws.values("A4:H11")                       # one native call; numbers are floats, dates serial days
assert rows[0][:3] == ["Survey of the river bank", "Field", 2.0] and rows[0][6] == "007"
total = ws.cell("E12")
print(total.formula, total.value, total.text)    # SUM(E4:E11) 61647.0 61,647.00
assert total.formula == "SUM(E4:E11)" and total.value == 61647.0 and not total.stale   # cached, and fresh
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
  {"cell": "B1", "formula": "SUM(Budget!E4:E11)"},
  {"cell": "B2", "formula": "'Q1 notes'!B2*2"}
]
JSON
ox xlsx edit budget.xlsx --map edits.json -o edited.xlsx
ox xlsx inspect edited.xlsx > inspect-edited.txt                  # "Recalculation: asked on open": Excel recalculates too
ox xlsx text edited.xlsx --sheet Budget                           # the figures are calculated: Qty 3 beside Total 12,600.00
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
assert ws.cell("I4").value == "yes" and ws.cell("E4").value == 12600.0   # calculated by the edit: 3 x 4,200
assert ws.cell("E4").stale and wb.recalculation_needed              # budget.xlsx said False: Excel recalculates on open too
assert ws.cell("A3").style.font.bold and ws.cell("A3").style.fill.foreground.rgb == "FFFF00"
assert wb.sheet("Summary").cell("B1").formula == "SUM(Budget!E4:E11)"
assert wb.sheet("Summary").cell("B2").formula == "'Q1 notes'!B2*2"      # a sheet name with a space is quoted
assert "Recalculation: asked on open" in open("inspect-edited.txt", encoding="utf-8").read()
assert [s.name for s in wb.sheets] == ["Budget", "Q1 notes", "Summary"]
print("E4 as the sheet holds it:", ws.cell("E4").text, "- the same as", 3 * 4200)
```

## Write from Python: values, dates, formulas, blocks, styles

```python
import datetime
from oxoffice.xlsx import Formula, Workbook

wb = Workbook("budget.xlsx")
ws = wb.sheet("Budget")
assert not wb.recalculation_needed and not ws.cell("E4").stale   # fresh, as Excel saved it
ws.set_value("C4", 3)                              # a number; "3" would be text
assert wb.recalculation_needed and ws.cell("E4").value == 12600.0   # the edit calculated the Total (and asks Excel to as well)
ws.set_value("G5", "013")                          # a str stays text
ws.set_value("F5", datetime.date(2026, 2, 27))     # shown yyyy-mm-dd in a General cell
ws.set_value("H4", None)                           # clears the content
ws.set_formula("I4", "IF(F4<DATE(2026,3,1),\"yes\",\"no\")")
ws.set_formula("J4", "XLOOKUP(A4,A4:A11,E4:E11)")  # stored with _xlfn., as Excel stores it
ws.set_formula("J5", "'Q1 notes'!B2+E5")           # a sheet name with a space is quoted; names are not checked
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
assert check.cell("J5").formula == "'Q1 notes'!B2+E5"
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
from pathlib import Path
from_file = Workbook("budget.xlsx")
from_file.sheet("Q1 notes").import_csv(Path("extra.csv"), at="D6")     # a Path is a file
assert from_file.sheet("Q1 notes").values("D6:E7") == [["item", "qty"], ["north", 12.0]]
not_a_file = Workbook("budget.xlsx").sheet("Q1 notes")
not_a_file.import_csv("extra.csv", at="D1")                                # a str is the CSV text itself
assert not_a_file.cell("D1").value == "extra.csv" and not_a_file.cell("E1").value is None
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

## Spilling formulas

```bash
ox xlsx set built.xlsx D1 "SORT(A2:A3)" --formula --spill-range D1:D2 --sheet Data -o spilled.xlsx
```

```python
from oxoffice.xlsx import Workbook
spilled = Workbook("spilled.xlsx").sheet("Data").cell("D1")
assert spilled.formula == "SORT(A2:A3)" and spilled.formula_kind == "array" and spilled.array_range == "D1:D2"
```

## Values for another reader (LibreOffice, pandas)

```bash
ox xlsx set budget.xlsx C4 3 --sheet Budget -o kept.xlsx                         # the Totals are calculated and stored
ox xlsx set budget.xlsx C4 3 --sheet Budget -o dropped.xlsx --drop-cached-values    # every cached value is emptied
ox xlsx inspect dropped.xlsx
```

```python
from oxoffice.xlsx import Workbook
kept, dropped = Workbook("kept.xlsx").sheet("Budget"), Workbook("dropped.xlsx").sheet("Budget")
assert kept.cell("E4").value == 12600.0 and kept.cell("E12").value == 65847.0   # pandas reads these as they are
assert dropped.cell("E4").value is None and dropped.cell("E12").value is None   # LibreOffice computes them on open
assert dropped.cell("E4").formula == "C4*D4" and dropped.cell("C4").value == 3.0
wb = Workbook("budget.xlsx")
assert Workbook.from_bytes(wb.to_bytes()).sheet("Budget").cell("E4").value == 8400.0
assert Workbook.from_bytes(wb.to_bytes(drop_cached_values=True)).sheet("Budget").cell("E4").value is None
```

## Calculated values, and what is skipped

```bash
ox xlsx calc qty.xlsx                                            # "11 formulas checked, 0 differ, 0 skipped"
ox xlsx set budget.xlsx J4 "STDEV(C4:C11)" --formula --sheet Budget -o skipped.xlsx 2> skipped.txt   # not a supported function
ox xlsx calc skipped.xlsx --json > calc.json                     # skipped: Budget!J4, "STDEV is not computed yet"
ox xlsx calc edited.xlsx --now 2026-03-01T09:00:00 -o computed.xlsx   # writes the computed values (a file from another program: do this first)
```

```python
import json
from oxoffice.xlsx import Workbook
assert "STDEV" in open("skipped.txt", encoding="utf-8").read()                 # the CLI says so on stderr
assert [(c["sheet"], c["cell"]) for c in json.load(open("calc.json", encoding="utf-8"))["skipped"]] == [("Budget", "J4")]
wb = Workbook("budget.xlsx")
ws = wb.sheet("Budget")
ws.set_formula("J4", "STDEV(C4:C11)")               # outside the supported functions: skipped, no value
ws.set_formula("J5", "J4*2")                        # reads a skipped formula: skipped too
ws.set_formula("J6", "SUMIF(B4:B11,\"Field\",E4:E11)+SUM(C4:C5)")   # supported: calculated at once
assert ws.cell("J6").value == 8400.0 + 3726.0 + 14.0
result = wb.recalculate()                           # computed, changed, skipped, cycles
assert [ref for ref, reason in result.skipped] == ["Budget!J4", "Budget!J5"] and not result.cycles
assert ws.cell("J4").value is None                  # never quote a skipped cell
wb.calculation = "manual"                           # edits keep their cached values until recalculate()
ws.set_value("C5", 1)
assert ws.cell("E5").value == 3726.0                # the old value
wb.recalculate()
assert ws.cell("E5").value == 310.5
wb.calculation = "automatic"
assert Workbook("computed.xlsx").sheet("Budget").cell("E4").value == 12600.0
```

## Sheets, rows and columns

```bash
ox xlsx sheet budget.xlsx rename "Q1 notes" Notes -o renamed.xlsx           # every formula, name and chart that reads it follows
ox xlsx sheet renamed.xlsx add Summary --at 1 -o added.xlsx
ox xlsx sheet added.xlsx hide Summary -o hidden.xlsx                        # also move Notes --to 3, show, remove
ox xlsx insert-rows budget.xlsx 6:7 --sheet Budget -o inserted.xlsx         # also delete-rows, insert-columns, delete-columns
ox xlsx diff budget.xlsx inserted.xlsx > diff-rows.txt                      # "6:7 rows_inserted": the rows below are not all changes
ox xlsx sheet edited.xlsx remove "Q1 notes" -o gone.xlsx                    # what read it shows #REF!
ox xlsx calc gone.xlsx -o gone-calc.xlsx                                    # write the values the removal changed
```

```python
from oxoffice.xlsx import Workbook
renamed = Workbook("renamed.xlsx")
assert [s.name for s in renamed.sheets] == ["Budget", "Notes"]
assert renamed.sheet("Notes").cell("B2").formula == "VLOOKUP(B1,Budget!A4:E11,5,FALSE)"
hidden = Workbook("hidden.xlsx")
assert [(s.name, s.state) for s in hidden.sheets] == [("Summary", "hidden"), ("Budget", "visible"), ("Notes", "visible")]
inserted = Workbook("inserted.xlsx")
assert inserted.sheet("Budget").cell("E14").formula == "SUM(E4:E13)"            # the total follows
assert inserted.sheet("Q1 notes").cell("B2").formula == "VLOOKUP(B1,Budget!A4:E13,5,FALSE)"   # so does another sheet
assert "rows_inserted" in open("diff-rows.txt", encoding="utf-8").read()
assert Workbook("gone-calc.xlsx").sheet("Summary").cell("B2").formula == "#REF!*2"
assert Workbook("gone-calc.xlsx").sheet("Summary").cell("B2").value == "#REF!"
wb = Workbook("budget.xlsx")
wb.add_sheet("Summary", at=0)                       # at a position; rename_sheet, remove_sheet, move_sheet(name, to)
wb.sheet("Q1 notes").hide()                         # show(), or hide(very=True)
ws = wb.sheet("Budget")
ws.delete_columns("B")                              # the Team column: Qty is now B, Total D
assert ws.cell("D12").formula == "SUM(D4:D11)" and ws.cell("D4").formula == "B4*C4" and ws.cell("D4").value == 8400.0
ws.insert_columns("B")
assert ws.cell("E12").formula == "SUM(E4:E11)"
ws.insert_rows("5:6")                               # two rows before row 5, with the formats of row 4
assert ws.cell("E14").formula == "SUM(E4:E13)" and ws.cell("A7").value == "Soil samples"
wb.save("restructured.xlsx")
```

## Names, tables, validation, conditional formats

```bash
cat > objects.json <<'JSON'
[
  {"sheet": "Budget", "define_name": "TaxRate", "refers_to": "0.2"},
  {"define_name": "Lines", "refers_to": "Budget!$A$4:$H$11"},
  {"add_table": "Costs", "range": "A3:H11", "style": "TableStyleMedium2"},
  {"validation": "H4:H11", "type": "list", "value": "Done,Open,Planned"},
  {"validation": "C4:C11", "type": "whole", "between": [0, 500], "error_message": "A whole number up to 500"},
  {"conditional_format": "E4:E11", "rule": "cellIs", "operator": "greaterThan", "value": "10000", "style": {"bold": true, "fill": "FFC7CE"}},
  {"conditional_format": "D4:D11", "rule": "dataBar", "color": "638EC6"}
]
JSON
ox xlsx edit budget.xlsx --map objects.json -o objects.xlsx
ox xlsx names objects.xlsx
ox xlsx tables objects.xlsx --json > tables.json
```

```python
import json
from oxoffice.xlsx import Workbook
wb = Workbook("objects.xlsx")
assert sorted((n.name, n.refers_to) for n in wb.names) == [("Lines", "Budget!$A$4:$H$11"), ("TaxRate", "0.2")]
ws = wb.sheet("Budget")
assert [(v.range, v.type) for v in ws.validations] == [("H4:H11", "list"), ("C4:C11", "whole")]
assert [(c.range, c.rule) for c in ws.conditional_formats] == [("E4:E11", "cellIs"), ("D4:D11", "dataBar")]
assert json.load(open("tables.json", encoding="utf-8"))["tables"][0]["name"] == "Costs"

sales = Workbook(sheets=["Sales"])
data = sales.sheet("Sales")
data.write("A1", [["Region", "Qty", "Amount"], ["North", 12, 1200.5], ["South", 7, 980], ["East", 30, 4410.25]])
data.add_table("Orders", "A1:C4", style="TableStyleMedium9", totals_row=True)   # the first row names the columns
assert data.cell("A5").value == "Total" and data.cell("C5").formula == "SUBTOTAL(109,Orders[Amount])"
assert data.cell("C5").value == 6590.75              # the totals row is calculated
data.set_formula("E1", "SUM(Orders[Amount])")        # structured references read the table
assert data.cell("E1").value == 6590.75
data.set_table_style("Orders", "TableStyleLight9")
data.set_validation("B2:B4", "whole", between=(0, 500), error_message="Up to 500")
data.set_validation("A2:A4", "list", value="North,South,East,West")
data.add_conditional_format("C2:C4", "cellIs", operator="lessThan", value=1000, fill="FFC7CE", bold=True)
data.add_conditional_format("B2:B4", "colorScale", colors=["F8696B", "63BE7B"])
sales.define_name("Amounts", "Sales!$C$2:$C$4")
sales.save("sales.xlsx")
check = Workbook("sales.xlsx").sheet("Sales")
assert [t.name for t in check.tables] == ["Orders"] and check.tables[0].style == "TableStyleLight9"
assert len(check.validations) == 2 and len(check.conditional_formats) == 2
```

## Notes and threads

```bash
ox xlsx comment budget.xlsx edit D6 "Quote valid until April." --sheet Budget -o note-edited.xlsx   # a note's text
ox xlsx comment budget.xlsx add B4 "Check the unit" --author Ann --sheet Budget -o note-added.xlsx
ox xlsx comment budget.xlsx thread E4 "Is 3 right?" --author Ann --at "2026-11-23 09:10" --sheet Budget -o thread.xlsx
ox xlsx comment thread.xlsx reply E4 "Yes, three." --author Bob --at "2026-11-23 10:00" --sheet Budget -o reply.xlsx
ox xlsx comment reply.xlsx resolve E4 --sheet Budget -o resolved.xlsx                      # also reopen, remove
ox xlsx comments resolved.xlsx --sheet Budget
```

```python
from oxoffice.xlsx import Workbook
import datetime
by_ref = lambda path: {c.ref: c for c in Workbook(path).sheet("Budget").comments}
assert by_ref("note-edited.xlsx")["D6"].text == "Quote valid until April." and by_ref("note-edited.xlsx")["D6"].author == "Ada"
assert by_ref("note-added.xlsx")["B4"].kind == "note"
done = by_ref("resolved.xlsx")["E4"]
assert done.kind == "thread" and done.resolved and [r.text for r in done.replies] == ["Yes, three."]
wb = Workbook("budget.xlsx")
ws = wb.sheet("Budget")
ws.edit_comment("E12", "Checked against the signed budget.")      # the note's text; its author stays
ws.add_thread("F4", "Ann", "Is the date firm?", at=datetime.datetime(2026, 11, 23, 9, 10))
ws.reply("F4", "Bob", "Until March.", at=datetime.datetime(2026, 11, 23, 10, 0))
wb.save("comments.xlsx")
written = {c.ref: c for c in Workbook("comments.xlsx").sheet("Budget").comments}
assert written["E12"].text == "Checked against the signed budget." and written["F4"].replies[0].author == "Bob"
```

## Merge, group, freeze, filter, sort, autofit, widths

```bash
ox xlsx merge budget.xlsx A13:D13 --sheet Budget -o merged-cells.xlsx
ox xlsx rows budget.xlsx 4:5 group --sheet Budget -o grouped.xlsx                  # also hide, show, collapse, level N, height H
ox xlsx panes budget.xlsx freeze B4 --sheet Budget -o frozen.xlsx                  # also split X Y, remove
ox xlsx columns budget.xlsx A:B autofit --sheet Budget -o fitted.xlsx              # the width the cells need
ox xlsx filter budget.xlsx set A3:H11 --sheet Budget -o filter-set.xlsx             # the first row is the header; no row is hidden yet
ox xlsx filter filter-set.xlsx where C --gt 4 --and --lt 50 --sheet Budget -o filtered.xlsx   # Qty over 4 and under 50: the other rows hide
ox xlsx sort budget.xlsx A3:H11 --header --by E:desc --sheet Budget -o sorted.xlsx      # a key of formulas needs values that are current: see cli.md
ox xlsx grid fitted.xlsx --sheet Budget --json > grid-fitted.json                         # widths in characters, heights in points, hidden rows
```

```python
import json
from oxoffice.xlsx import Compare, Workbook
grid = json.load(open("grid-fitted.json", encoding="utf-8"))
assert grid["sheets"][0]["name"] == "Budget"
sorted_sheet = Workbook("sorted.xlsx").sheet("Budget")
assert sorted_sheet.values("A4:A5") == [["Pier design"], ["Deck design"]]            # 18,500 then 14,250
assert sorted_sheet.cell("E4").formula == "C4*D4"                                    # formulas move as copies
assert Workbook("grouped.xlsx").sheet("Budget").row_dimension("4:5").outline_level == 1
assert Workbook("frozen.xlsx").sheet("Budget").panes.columns == 1

wb = Workbook("budget.xlsx")
ws = wb.sheet("Budget")
ws.merge("A13:D13")                                 # the top-left cell keeps its content
ws.group_rows("4:5"); ws.freeze_panes("B4")
ws.set_autofilter("A3:H11")                         # the first row is the header
outcome = ws.set_filter("E", Compare(">", 5000).and_("<", 20000))   # hides the rows that do not pass, now
assert outcome.hidden == 4 and len(ws.hidden_rows) == 4          # a SUBTOTAL leaves them out
ws.clear_filter()
ws.sort("A3:H11", by=["B", ("E", "desc")], header=True)   # by Team, then Total, largest first
assert ws.values("A4:B5") == [["Permit fees", "Admin"], ["Crane hire", "Build"]]
widths = ws.autofit_columns("A:B", min=8, max=40)   # {"A": ..., "B": ...} in characters
assert set(widths) == {"A", "B"} and ws.column_dimension("A").custom_width
heights = ws.autofit_rows("4:11")
ws.set_column_width("A:B", 30)                      # or set one by hand
wb.save("layout.xlsx")
```

## Compare two workbooks

```bash
ox xlsx diff budget.xlsx qty.xlsx                    # sheets, cells, formulas, comments, merges, inserted rows
ox xlsx diff budget.xlsx qty.xlsx --json > diff.json
ox xlsx diff budget.xlsx budget.xlsx --exit-code     # exit 0: nothing changed (1 when they differ)
```

```python
from oxoffice.xlsx import diff
changes = diff("budget.xlsx", "qty.xlsx")
assert {(c.ref, c.kind) for c in changes.cells if c.kind == "value"} >= {("C4", "value"), ("E4", "value")}
assert changes and changes.summary.cells > 0 and not diff("budget.xlsx", "budget.xlsx")
assert "C4" in changes.to_json()
```

## Render to PDF and PNG

```bash
ox xlsx convert budget.xlsx -t pdf -o budget.pdf 2> render-notes.txt             # every visible sheet, on its own paper
ox xlsx convert budget.xlsx -t png --pages 1 --dpi 100 -o page.png
ox xlsx render budget.xlsx --pages 1 -o pages --dpi 80                           # pages/budget_page1.png
```

```python
from oxoffice.xlsx import Workbook
wb = Workbook("budget.xlsx")
pdf = wb.to_pdf()                                    # bytes; sheets=["Budget"] for one sheet
assert pdf.startswith(b"%PDF") and wb.page_count() >= 1
png = wb.render_page_to_png(0, dpi=100)              # page 0, None past the last page
assert png[:4] == b"\x89PNG" and wb.render_page_to_png(99) is None
open("budget-page.png", "wb").write(png)             # look at it before saying the layout is right
assert open("budget.pdf", "rb").read(4) == b"%PDF" and open("page.png", "rb").read(4) == b"\x89PNG"
```
