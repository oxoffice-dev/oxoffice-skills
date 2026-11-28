"""Generate budget.xlsx: a synthetic project budget, written from scratch with XlsxWriter (invented figures), the way
a workbook arrives from Excel or a script: formulas with their cached values, dates, a text code with a leading zero,
notes, a merged title, widths and a frozen header row.

Usage: make_fixture_xlsx.py [OUT]      (default budget.xlsx; needs XlsxWriter; the output is deterministic)

XlsxWriter asks Excel to recalculate on open (`fullCalcOnLoad`); Excel does not once it has saved the workbook, so the
flag is stripped: the workbook starts with fresh cached values, `recalculation_needed` false and no formula `stale`,
as a file saved by Excel. Only an edit then makes them stale, which is what the examples check.

Sheets:
- Budget: a merged title (row 1), a header (row 3), eight lines (rows 4 to 11) with Item, Team, Qty, Unit price,
  Total (a formula), Due (a date), Code (text, "007") and a Status; a total row (row 12) with SUM formulas; two notes
- Q1 notes: a sheet whose name holds a space, with a lookup formula over the Budget sheet
"""
import datetime
import sys
import zipfile

import xlsxwriter

OUT = sys.argv[1] if len(sys.argv) > 1 else "budget.xlsx"
LINES = [
    ("Survey of the river bank", "Field", 2, 4200.00, datetime.date(2026, 2, 6), "007", "Done"),
    ("Soil samples", "Field", 12, 310.50, datetime.date(2026, 2, 20), "012", "Done"),
    ("Pier design", "Design", 1, 18500.00, datetime.date(2026, 3, 13), "021", "Open"),
    ("Deck design", "Design", 1, 14250.00, datetime.date(2026, 3, 27), "022", "Open"),
    ("Permit fees", "Admin", 3, 975.00, datetime.date(2026, 4, 10), "030", "Open"),
    ("Steel order", "Build", 40, 82.40, datetime.date(2026, 5, 8), "041", "Planned"),
    ("Crane hire", "Build", 5, 1650.00, datetime.date(2026, 5, 22), "042", "Planned"),
    ("Site cleanup", "Build", 1, 2300.00, datetime.date(2026, 6, 19), "043", "Planned"),
]

wb = xlsxwriter.Workbook(OUT, {"default_date_format": "yyyy-mm-dd"})
wb.set_properties({"title": "Footbridge budget", "created": datetime.datetime(2026, 1, 5)})
title = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
head = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
money = wb.add_format({"num_format": "#,##0.00"})
date = wb.add_format({"num_format": "yyyy-mm-dd"})
text = wb.add_format({"num_format": "@"})
bold_money = wb.add_format({"bold": True, "num_format": "#,##0.00"})

ws = wb.add_worksheet("Budget")
ws.merge_range("A1:H1", "Footbridge budget, 2026", title)
for col, name in enumerate(["Item", "Team", "Qty", "Unit price", "Total", "Due", "Code", "Status"]):
    ws.write(2, col, name, head)
for k, (item, team, qty, price, due, code, status) in enumerate(LINES):
    r = 3 + k
    ws.write(r, 0, item)
    ws.write(r, 1, team)
    ws.write_number(r, 2, qty)
    ws.write_number(r, 3, price, money)
    ws.write_formula(r, 4, f"C{r + 1}*D{r + 1}", money, qty * price)
    ws.write_datetime(r, 5, datetime.datetime.combine(due, datetime.time()), date)
    ws.write_string(r, 6, code, text)
    ws.write(r, 7, status)
qty_sum = sum(x[2] for x in LINES)
total = sum(x[2] * x[3] for x in LINES)
ws.write(11, 0, "Total", wb.add_format({"bold": True}))
ws.write_formula(11, 2, "SUM(C4:C11)", None, qty_sum)
ws.write_formula(11, 4, "SUM(E4:E11)", bold_money, total)
ws.write_comment("D6", "Quote from the designer, valid until March.", {"author": "Ada"})
ws.write_comment("E12", "Check against the signed budget.", {"author": "Grace"})
ws.set_column("A:A", 26)
ws.set_column("B:D", 11)
ws.set_column("E:F", 14)
ws.freeze_panes(3, 0)

notes = wb.add_worksheet("Q1 notes")
notes.write("A1", "Item to look up")
notes.write("B1", "Steel order")
notes.write("A2", "Its total")
notes.write_formula("B2", "VLOOKUP(B1,Budget!A4:E11,5,FALSE)", money, 40 * 82.40)
notes.write("A4", "Written by the field team on 2026-01-05.")
notes.set_column("A:B", 24)
wb.close()

with zipfile.ZipFile(OUT) as z:
    parts = [(info, z.read(info.filename)) for info in z.infolist()]
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in parts:
        if info.filename == "xl/workbook.xml":
            assert b' fullCalcOnLoad="1"' in data
            data = data.replace(b' fullCalcOnLoad="1"', b"")
        z.writestr(info, data, zipfile.ZIP_DEFLATED)
print(OUT)
