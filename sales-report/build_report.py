"""
Build the marketplace sales report workbook from an EasyEcom-style combined export.

    python3 build_report.py <export.csv> [output.xlsx]

Produces an Excel workbook whose every figure is driven by live formulas off the
'Raw Data' sheet, so pasting a fresh export refreshes the whole report.
"""
import csv, os, sys
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as CL
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import ColorScaleRule

HERE  = os.path.dirname(os.path.abspath(__file__))
CSVF  = sys.argv[1] if len(sys.argv) > 1 else None
OUT   = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "Marketplace_Sales_Report.xlsx")
SKUF  = os.path.join(HERE, "skus.txt")

WORK_LAST = 30001          # Work sheet formulas cover rows 2..30001
NDAYS     = 15
PLATFORMS = ["Amazon", "Flipkart", "Myntra"]

skus = [l.strip() for l in open(SKUF) if l.strip()]
NS   = len(skus)

# ---------- styling ----------
NAVY   = "1F3864"; BLUE = "2E5C8A"; LIGHT = "DCE6F1"; GREY = "F2F2F2"
ACCENT = "FFF2CC"; GREEN = "E2EFDA"
F_TITLE = Font(bold=True, size=14, color="FFFFFF")
F_HEAD  = Font(bold=True, size=10, color="FFFFFF")
F_SUB   = Font(bold=True, size=10, color=NAVY)
F_B     = Font(bold=True, size=10)
FILL_T  = PatternFill("solid", fgColor=NAVY)
FILL_H  = PatternFill("solid", fgColor=BLUE)
FILL_L  = PatternFill("solid", fgColor=LIGHT)
FILL_A  = PatternFill("solid", fgColor=ACCENT)
FILL_G  = PatternFill("solid", fgColor=GREEN)
thin    = Side(style="thin", color="BFBFBF")
BORDER  = Border(left=thin, right=thin, top=thin, bottom=thin)
CTR     = Alignment(horizontal="center", vertical="center")
DATEFMT = "dd-mmm-yy"

def title_row(ws, text, width):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    c = ws.cell(row=1, column=1, value=text); c.font = F_TITLE; c.fill = FILL_T
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24

wb = Workbook()

# =====================================================================
# 1. README
# =====================================================================
ws = wb.active; ws.title = "README"
title_row(ws, "MARKETPLACE SALES REPORT  —  how to use", 2)
lines = [
 ("", ""),
 ("STEP 1  —  Paste the report", ""),
 ("", "Open the 'Raw Data' sheet. Select cell A2. Paste the full export (data rows only, WITHOUT the header row)."),
 ("", "Header row 1 is already in place and must not be changed - every formula depends on those column positions."),
 ("", "Before pasting fresh data clear the old rows: select row 2 down to the last filled row and press Delete."),
 ("", "The export may be from any period - the workbook re-reads it automatically. Nothing else needs editing."),
 ("", ""),
 ("STEP 2  —  Everything recalculates", ""),
 ("", "All sheets update on their own. Press F9 if your Excel is set to manual calculation."),
 ("", ""),
 ("THE 15-DAY WINDOW", ""),
 ("", "The workbook finds the LATEST order date in the pasted data, drops it (it is always a partial day),"),
 ("", "and reports the 15 full days before it."),
 ("", "Example: an export running 1 Sep - 18 Sep reports 3 Sep - 17 Sep."),
 ("", "Change the length, the number of trailing days dropped, or force an exact end date on the 'Settings' sheet."),
 ("", ""),
 ("WHAT COUNTS AS A SALE", ""),
 ("", "Every order line in the export counts, whatever its status - Shipped, Printed, Assigned, Upcoming AND Cancelled."),
 ("", "No status filter is applied anywhere, exactly as requested."),
 ("", "Units per line = the larger of 'Item Quantity' and 'Suborder Quantity' (guards against rows where one field is 0)."),
 ("", "Sales are booked on 'Order Date'."),
 ("", ""),
 ("PLATFORMS", ""),
 ("", "Read from the 'MP Name' column:  contains 'Amazon' -> Amazon,  'Flipkart' -> Flipkart,  'Myntra' -> Myntra."),
 ("", "Anything else is tagged 'Other' and is excluded from the platform sheets (see the Check columns on 'Daily Orders')."),
 ("", ""),
 ("THE SHEETS", ""),
 ("Settings", "Date window controls, and a live summary of the current window."),
 ("Daily Orders", "Day-by-day order volume per platform: units, order lines, and daily averages."),
 ("Combined Daily", "One block per date; under each date the three marketplaces, then a day TOTAL row. SKUs across the top."),
 ("Amazon / Flipkart / Myntra", "Date-wise x SKU-wise units for that platform, plus 15-day Total and Average per Day."),
 ("Best Sellers - ...", "All 182 SKUs ranked high to low for each platform, plus an All-Platforms ranking."),
 ("SKU Master", "The 182 SKUs driving every sheet. Add or remove SKUs here - see the note on that sheet."),
 ("Raw Data", "The paste area."),
 ("Work", "Hidden helper sheet that normalises date, platform, SKU and quantity. Do not edit."),
 ("", ""),
 ("NOTES", ""),
 ("", "Every SKU is listed on every sheet. A SKU with no sales in the window shows 0, not a blank."),
 ("", "Figures are UNITS (quantity), not order counts. 'Daily Orders' also shows order-line counts side by side."),
 ("", "'Work' covers 30,000 data rows (~55 days of this export). To paste more, copy the last Work row further down."),
]
r = 2
for a, b in lines:
    if a and not b:
        c = ws.cell(row=r, column=1, value=a); c.font = Font(bold=True, size=11, color=NAVY); c.fill = FILL_L
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    else:
        if a: ws.cell(row=r, column=1, value=a).font = F_B
        if b: ws.cell(row=r, column=2, value=b)
    r += 1
ws.column_dimensions["A"].width = 28; ws.column_dimensions["B"].width = 118
ws.sheet_view.showGridLines = False

# =====================================================================
# 2. Settings
# =====================================================================
st = wb.create_sheet("Settings")
title_row(st, "SETTINGS  —  reporting window", 4)
rows = [
 (3,  "DETECTED FROM RAW DATA", None, None, None),
 (4,  "Earliest order date in Raw Data", '=IF(COUNT(WDATE)=0,"",MIN(WDATE))', DATEFMT, None),
 (5,  "Latest order date in Raw Data",   '=IF(COUNT(WDATE)=0,"",MAX(WDATE))', DATEFMT, None),
 (6,  "Order lines detected",            '=COUNT(WDATE)', "#,##0", None),
 (8,  "YOU CAN CHANGE THESE", None, None, None),
 (9,  "Number of days to report",              NDAYS, "0",  "15 = a fortnight plus a day. Change to 7, 30, etc."),
 (10, "Trailing days to drop",                 1,     "0",  "The export's last day is partial, so 1 drops it."),
 (11, "Manual END date (blank = automatic)",   None,  DATEFMT, "Fill this only to pin the window to an exact end date."),
 (13, "WINDOW IN USE", None, None, None),
 (14, "END date",   '=IF($B$11<>"",$B$11,IF(N($B$5)=0,"",$B$5-$B$10))', DATEFMT, None),
 (15, "START date", '=IF(N($B$14)=0,"",$B$14-$B$9+1)',                  DATEFMT, None),
 (17, "WINDOW TOTALS", None, None, None),
 (18, "Total units - Amazon",   "=Amazon!Q%d"   % (5 + NS), "#,##0", None),
 (19, "Total units - Flipkart", "=Flipkart!Q%d" % (5 + NS), "#,##0", None),
 (20, "Total units - Myntra",   "=Myntra!Q%d"   % (5 + NS), "#,##0", None),
 (21, "Total units - ALL",      "=SUM(B18:B20)", "#,##0", None),
 (22, "Average units per day - ALL", '=IF($B$9=0,0,B21/$B$9)', "#,##0.0", None),
 (23, "SKUs tracked",           "=COUNTA(SKULIST)", "#,##0", None),
 (24, "SKUs with sales - ALL",  '=SUMPRODUCT(--((Amazon!Q5:Q%d+Flipkart!Q5:Q%d+Myntra!Q5:Q%d)>0))' % (4+NS,4+NS,4+NS), "#,##0", None),
]
for rr, label, val, fmt, note in rows:
    if val is None and fmt is None and note is None:
        c = st.cell(row=rr, column=1, value=label); c.font = F_HEAD; c.fill = FILL_H
        st.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=4)
        continue
    st.cell(row=rr, column=1, value=label).font = F_B
    c = st.cell(row=rr, column=2, value=val)
    if fmt: c.number_format = fmt
    c.border = BORDER; c.alignment = CTR
    if rr in (9, 10, 11): c.fill = FILL_A
    if note: st.cell(row=rr, column=3, value=note).font = Font(italic=True, size=9, color="808080")
st.column_dimensions["A"].width = 36; st.column_dimensions["B"].width = 16
st.column_dimensions["C"].width = 60; st.column_dimensions["D"].width = 4
st.sheet_view.showGridLines = False

# =====================================================================
# 3. SKU Master
# =====================================================================
sm = wb.create_sheet("SKU Master")
title_row(sm, "SKU MASTER  —  the list every sheet is built from", 4)
sm.cell(row=2, column=1, value="Adding a SKU: insert a row inside A3:A%d (never at the very end), then copy the row above it down on "
                               "Amazon / Flipkart / Myntra / Combined Daily / the Best Sellers sheets." % (2 + NS)
        ).font = Font(italic=True, size=9, color="808080")
sm.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
for i, h in enumerate(["SKU", "Product", "Colour", "Size"], start=1):
    c = sm.cell(row=3, column=i, value=h); c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER; c.alignment = CTR
SIZEMAP = {"ss": "S", "sm": "M", "sl": "L", "sxl": "XL", "sxxl": "XXL", "sxxxl": "3XL", "s4xl": "4XL"}
PRODMAP = {"jellybra": "JellyLift Bra", "stayputstrapless": "StayPut Strapless",
           "instatuck": "InstaTuck", "widestr": "Wide Strap Bra"}
for i, s in enumerate(skus):
    p = s.split("-")
    sm.cell(row=4 + i, column=1, value=s).border = BORDER
    sm.cell(row=4 + i, column=2, value=PRODMAP.get(p[0], p[0])).border = BORDER
    c = sm.cell(row=4 + i, column=3, value=p[2][1:] if len(p) > 2 else ""); c.border = BORDER
    c = sm.cell(row=4 + i, column=4, value=SIZEMAP.get(p[3], p[3]) if len(p) > 3 else ""); c.border = BORDER; c.alignment = CTR
sm.column_dimensions["A"].width = 34; sm.column_dimensions["B"].width = 20
sm.column_dimensions["C"].width = 18; sm.column_dimensions["D"].width = 8
sm.freeze_panes = "A4"; sm.auto_filter.ref = "A3:D%d" % (3 + NS)
sm.sheet_view.showGridLines = False

# =====================================================================
# 4. Platform sheets  (Amazon / Flipkart / Myntra)
# =====================================================================
FIRST, LAST = 5, 4 + NS        # SKU rows
TOTR = LAST + 1                # TOTAL row
DCOL0 = 2                      # first date column = B
QCOL, RCOL, SCOL = 17, 18, 20  # Total, Avg/Day, hidden score

def platform_sheet(name):
    p = wb.create_sheet(name)
    title_row(p, "%s  —  DATE-WISE / SKU-WISE SALES  (units)" % name.upper(), RCOL)
    p.cell(row=2, column=1, value="Platform:").font = F_B
    c = p.cell(row=2, column=2, value=name); c.font = F_SUB; c.fill = FILL_A; c.border = BORDER; c.alignment = CTR
    p.cell(row=2, column=4, value="Counts every order status, cancelled included. Sales booked on Order Date."
           ).font = Font(italic=True, size=9, color="808080")
    p.cell(row=3, column=1, value="Window:").font = F_B
    c = p.cell(row=3, column=2, value="=DSTART"); c.number_format = DATEFMT; c.font = F_B; c.alignment = CTR
    p.cell(row=3, column=3, value="to").alignment = CTR
    c = p.cell(row=3, column=4, value="=DEND"); c.number_format = DATEFMT; c.font = F_B; c.alignment = CTR
    p.cell(row=3, column=5, value="Days:").font = F_B
    c = p.cell(row=3, column=6, value="=Settings!$B$9"); c.font = F_B; c.alignment = CTR

    c = p.cell(row=4, column=1, value="SKU"); c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER; c.alignment = CTR
    for d in range(NDAYS):
        col = DCOL0 + d
        c = p.cell(row=4, column=col, value='=IF(N(DSTART)=0,"",DSTART)' if d == 0 else '=IF(N(%s4)=0,"",%s4+1)' % (CL(col-1), CL(col-1)))
        c.number_format = "dd-mmm"; c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", textRotation=90)
        p.column_dimensions[CL(col)].width = 7
    for col, lbl in ((QCOL, "Total (window)"), (RCOL, "Avg / Day")):
        c = p.cell(row=4, column=col, value=lbl); c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    p.row_dimensions[4].height = 62

    for i, s in enumerate(skus):
        r = FIRST + i
        c = p.cell(row=r, column=1, value="='SKU Master'!$A$%d" % (4 + i)); c.border = BORDER; c.font = Font(size=9)
        for d in range(NDAYS):
            col = CL(DCOL0 + d)
            c = p.cell(row=r, column=DCOL0 + d,
                       value="=SUMIFS(WQTY,WSKU,$A%d,WPLAT,$B$2,WDATE,%s$4)" % (r, col))
            c.border = BORDER; c.number_format = "0;;\"0\""; c.font = Font(size=9); c.alignment = CTR
        c = p.cell(row=r, column=QCOL, value="=SUM(%s%d:%s%d)" % (CL(DCOL0), r, CL(DCOL0 + NDAYS - 1), r))
        c.border = BORDER; c.font = F_B; c.fill = FILL_L; c.number_format = "#,##0"; c.alignment = CTR
        c = p.cell(row=r, column=RCOL, value='=IF($F$3=0,0,%s%d/$F$3)' % (CL(QCOL), r))
        c.border = BORDER; c.number_format = "#,##0.00"; c.fill = FILL_L; c.alignment = CTR
        # hidden tie-broken score for the Best Sellers ranking
        p.cell(row=r, column=SCOL, value="=%s%d+(%d-ROW())/1000000" % (CL(QCOL), r, 1000))

    c = p.cell(row=TOTR, column=1, value="TOTAL — ALL SKUs"); c.font = F_HEAD; c.fill = FILL_T; c.border = BORDER
    for col in list(range(DCOL0, DCOL0 + NDAYS)) + [QCOL, RCOL]:
        L = CL(col)
        c = p.cell(row=TOTR, column=col, value="=SUM(%s%d:%s%d)" % (L, FIRST, L, LAST))
        c.font = F_HEAD; c.fill = FILL_T; c.border = BORDER; c.alignment = CTR
        c.number_format = "#,##0.00" if col == RCOL else "#,##0"
    p.column_dimensions["A"].width = 34
    p.column_dimensions[CL(QCOL)].width = 11; p.column_dimensions[CL(RCOL)].width = 11
    p.column_dimensions[CL(SCOL)].hidden = True
    p.freeze_panes = "B5"
    p.conditional_formatting.add(
        "%s%d:%s%d" % (CL(DCOL0), FIRST, CL(DCOL0 + NDAYS - 1), LAST),
        ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                       end_type="max", end_color="4472C4"))
    p.sheet_view.showGridLines = False
    return p

for pl in PLATFORMS:
    platform_sheet(pl)

# =====================================================================
# 5. Daily Orders
# =====================================================================
do = wb.create_sheet("Daily Orders", 2)
WID = 12
title_row(do, "DAILY ORDERS  —  volume by date and marketplace", WID)
do.cell(row=2, column=1, value="Units = quantity sold on the tracked SKUs.   Order Lines = number of order lines in the export "
        "(all SKUs).   The Check columns should read 0.").font = Font(italic=True, size=9, color="808080")
heads = ["Date", "Amazon\nUnits", "Flipkart\nUnits", "Myntra\nUnits", "TOTAL\nUnits",
         "Amazon\nOrder Lines", "Flipkart\nOrder Lines", "Myntra\nOrder Lines", "TOTAL\nOrder Lines",
         "Check:\nUnits all SKUs", "Check:\nUnlisted SKUs", "Day of\nWeek"]
for i, h in enumerate(heads, start=1):
    c = do.cell(row=4, column=i, value=h); c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
do.row_dimensions[4].height = 34
for d in range(NDAYS):
    r = 5 + d
    c = do.cell(row=r, column=1, value='=IF(N(DSTART)=0,"",DSTART)' if d == 0 else '=IF(N(A%d)=0,"",A%d+1)' % (r-1, r-1))
    c.number_format = DATEFMT; c.border = BORDER; c.font = F_B; c.alignment = CTR
    for j, pl in enumerate(PLATFORMS):                      # units, straight from platform TOTAL row
        c = do.cell(row=r, column=2 + j, value="=%s!%s%d" % (pl, CL(DCOL0 + d), TOTR))
        c.border = BORDER; c.number_format = "#,##0"; c.alignment = CTR
    c = do.cell(row=r, column=5, value="=SUM(B%d:D%d)" % (r, r))
    c.border = BORDER; c.number_format = "#,##0"; c.font = F_B; c.fill = FILL_L; c.alignment = CTR
    for j, pl in enumerate(PLATFORMS):                      # order lines
        c = do.cell(row=r, column=6 + j, value='=COUNTIFS(WDATE,$A%d,WPLAT,"%s")' % (r, pl))
        c.border = BORDER; c.number_format = "#,##0"; c.alignment = CTR
    c = do.cell(row=r, column=9, value="=SUM(F%d:H%d)" % (r, r))
    c.border = BORDER; c.number_format = "#,##0"; c.font = F_B; c.fill = FILL_L; c.alignment = CTR
    c = do.cell(row=r, column=10, value="=SUMIFS(WQTY,WDATE,$A%d)" % r)
    c.border = BORDER; c.number_format = "#,##0"; c.alignment = CTR; c.font = Font(size=9, color="808080")
    c = do.cell(row=r, column=11, value="=J%d-E%d" % (r, r))
    c.border = BORDER; c.number_format = "#,##0"; c.alignment = CTR; c.font = Font(size=9, color="808080")
    c = do.cell(row=r, column=12, value='=TEXT($A%d,"ddd")' % r)
    c.border = BORDER; c.alignment = CTR; c.font = Font(size=9, color="808080")
tr = 5 + NDAYS
c = do.cell(row=tr, column=1, value="TOTAL"); c.font = F_HEAD; c.fill = FILL_T; c.border = BORDER
for col in range(2, 12):
    L = CL(col)
    c = do.cell(row=tr, column=col, value="=SUM(%s5:%s%d)" % (L, L, tr - 1))
    c.font = F_HEAD; c.fill = FILL_T; c.border = BORDER; c.number_format = "#,##0"; c.alignment = CTR
ar = tr + 1
c = do.cell(row=ar, column=1, value="AVERAGE PER DAY"); c.font = F_B; c.fill = FILL_G; c.border = BORDER
for col in range(2, 12):
    L = CL(col)
    c = do.cell(row=ar, column=col, value="=IF(Settings!$B$9=0,0,%s%d/Settings!$B$9)" % (L, tr))
    c.font = F_B; c.fill = FILL_G; c.border = BORDER; c.number_format = "#,##0.0"; c.alignment = CTR
for col, w in [(1, 13)] + [(i, 12) for i in range(2, 12)] + [(12, 9)]:
    do.column_dimensions[CL(col)].width = w
do.freeze_panes = "B5"
do.conditional_formatting.add("B5:D%d" % (tr - 1),
    ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="70AD47"))
do.sheet_view.showGridLines = False

# =====================================================================
# 6. Combined Daily  (date block -> Amazon / Flipkart / Myntra / day total)
# =====================================================================
cd = wb.create_sheet("Combined Daily", 3)
SK0 = 3                       # first SKU column = C
TOTCOL = SK0 + NS             # grand total column
title_row(cd, "COMBINED DAILY SALES  —  each date, then Amazon / Flipkart / Myntra (units)", TOTCOL)
cd.cell(row=2, column=1, value="Every SKU appears for every date and marketplace. No sales shows as 0.").font = Font(italic=True, size=9, color="808080")
for i, h in enumerate(["Date", "Marketplace"], start=1):
    c = cd.cell(row=4, column=i, value=h); c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER; c.alignment = CTR
for i in range(NS):
    c = cd.cell(row=4, column=SK0 + i, value="='SKU Master'!$A$%d" % (4 + i))
    c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="bottom", textRotation=90)
    cd.column_dimensions[CL(SK0 + i)].width = 5.5
c = cd.cell(row=4, column=TOTCOL, value="TOTAL\nAll SKUs"); c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
cd.row_dimensions[4].height = 150
row = 5
for d in range(NDAYS):
    base = row
    for k, pl in enumerate(PLATFORMS + ["TOTAL (day)"]):
        r = base + k
        istot = (k == 3)
        if k == 0:
            c = cd.cell(row=r, column=1, value='=IF(N(DSTART)=0,"",DSTART)' if d == 0 else '=IF(N(A%d)=0,"",A%d+1)' % (base-4, base-4))
            c.number_format = DATEFMT
        else:
            c = cd.cell(row=r, column=1, value="=$A$%d" % base); c.number_format = DATEFMT
            c.font = Font(size=9, color="A6A6A6")
        c.border = BORDER; c.alignment = CTR
        if k == 0: c.font = F_B
        c = cd.cell(row=r, column=2, value=pl); c.border = BORDER; c.alignment = CTR
        c.font = F_HEAD if istot else Font(size=10)
        if istot: c.fill = FILL_T
        else: c.fill = FILL_L if k % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for i in range(NS):
            col = SK0 + i
            if istot:
                v = "=SUM(%s%d:%s%d)" % (CL(col), base, CL(col), base + 2)
            else:
                v = "=%s!%s%d" % (pl, CL(DCOL0 + d), FIRST + i)
            c = cd.cell(row=r, column=col, value=v); c.border = BORDER
            c.number_format = "0;;\"0\""; c.font = Font(size=8); c.alignment = CTR
            if istot: c.font = Font(size=8, bold=True, color="FFFFFF"); c.fill = FILL_T
        v = "=SUM(%s%d:%s%d)" % (CL(SK0), r, CL(SK0 + NS - 1), r)
        c = cd.cell(row=r, column=TOTCOL, value=v); c.border = BORDER
        c.number_format = "#,##0"; c.alignment = CTR
        c.font = Font(size=9, bold=True, color="FFFFFF") if istot else F_B
        if istot: c.fill = FILL_T
        else: c.fill = FILL_A
    row = base + 4
cd.column_dimensions["A"].width = 12; cd.column_dimensions["B"].width = 14
cd.column_dimensions[CL(TOTCOL)].width = 11
cd.freeze_panes = "C5"
cd.sheet_view.showGridLines = False

# =====================================================================
# 7. Best Sellers  (one per platform + all platforms)
# =====================================================================
def best_sellers(label, sheets):
    bs = wb.create_sheet("Best Sellers - " + label)
    title_row(bs, "BEST SELLERS  —  %s  (all %d SKUs, ranked)" % (label.upper(), NS), 7)
    bs.cell(row=2, column=1, value="Ranked on units in the reporting window. Re-ranks itself whenever new data is pasted."
            ).font = Font(italic=True, size=9, color="808080")
    c = bs.cell(row=3, column=1, value="Window:"); c.font = F_B
    c = bs.cell(row=3, column=2, value="=DSTART"); c.number_format = DATEFMT; c.font = F_B; c.alignment = CTR
    bs.cell(row=3, column=3, value="to").alignment = CTR
    c = bs.cell(row=3, column=4, value="=DEND"); c.number_format = DATEFMT; c.font = F_B; c.alignment = CTR
    for i, h in enumerate(["Rank", "SKU", "Product", "Size", "Units\n(window)", "Avg / Day", "% of total"], start=1):
        c = bs.cell(row=4, column=i, value=h); c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bs.row_dimensions[4].height = 30
    # single-platform -> read that sheet's score column; all-platforms -> local hidden score in col I/J
    if len(sheets) == 1:
        s = sheets[0]
        SRC_SCORE = "%s!$%s$%d:$%s$%d" % (s, CL(SCOL), FIRST, CL(SCOL), LAST)
        SRC_SKU   = "%s!$A$%d:$A$%d"   % (s, FIRST, LAST)
        SRC_TOT   = "%s!$%s$%d"        % (s, CL(QCOL), TOTR)
    else:
        for i in range(NS):
            r = 5 + i
            bs.cell(row=r, column=9,  value="=" + "+".join("%s!$%s$%d" % (s, CL(QCOL), FIRST + i) for s in sheets))
            bs.cell(row=r, column=10, value="=I%d+(1000-ROW())/1000000" % r)
        bs.column_dimensions["I"].hidden = True; bs.column_dimensions["J"].hidden = True
        SRC_SCORE = "$J$5:$J$%d" % (4 + NS)
        SRC_SKU   = "'SKU Master'!$A$4:$A$%d" % (3 + NS)
        SRC_TOT   = "SUM($I$5:$I$%d)" % (4 + NS)
    for i in range(NS):
        r = 5 + i
        c = bs.cell(row=r, column=1, value=i + 1); c.border = BORDER; c.alignment = CTR; c.font = F_B
        c = bs.cell(row=r, column=2,
                    value="=INDEX(%s,MATCH(LARGE(%s,$A%d),%s,0))" % (SRC_SKU, SRC_SCORE, r, SRC_SCORE))
        c.border = BORDER; c.font = Font(size=9)
        c = bs.cell(row=r, column=3, value="=IFERROR(INDEX('SKU Master'!$B$4:$B$%d,MATCH($B%d,SKULIST,0)),\"\")" % (3 + NS, r))
        c.border = BORDER; c.font = Font(size=9)
        c = bs.cell(row=r, column=4, value="=IFERROR(INDEX('SKU Master'!$D$4:$D$%d,MATCH($B%d,SKULIST,0)),\"\")" % (3 + NS, r))
        c.border = BORDER; c.font = Font(size=9); c.alignment = CTR
        c = bs.cell(row=r, column=5, value="=INT(LARGE(%s,$A%d))" % (SRC_SCORE, r))
        c.border = BORDER; c.number_format = "#,##0"; c.font = F_B; c.fill = FILL_L; c.alignment = CTR
        c = bs.cell(row=r, column=6, value="=IF(Settings!$B$9=0,0,$E%d/Settings!$B$9)" % r)
        c.border = BORDER; c.number_format = "#,##0.00"; c.alignment = CTR
        c = bs.cell(row=r, column=7, value="=IF(%s=0,0,$E%d/%s)" % (SRC_TOT, r, SRC_TOT))
        c.border = BORDER; c.number_format = "0.0%"; c.alignment = CTR
    tr2 = 5 + NS
    c = bs.cell(row=tr2, column=1, value="TOTAL"); c.font = F_HEAD; c.fill = FILL_T; c.border = BORDER
    bs.merge_cells(start_row=tr2, start_column=1, end_row=tr2, end_column=4)
    for col, fmt in ((5, "#,##0"), (6, "#,##0.00"), (7, "0.0%")):
        c = bs.cell(row=tr2, column=col, value="=SUM(%s5:%s%d)" % (CL(col), CL(col), tr2 - 1))
        c.font = F_HEAD; c.fill = FILL_T; c.border = BORDER; c.number_format = fmt; c.alignment = CTR
    for col, w in ((1, 7), (2, 34), (3, 20), (4, 8), (5, 11), (6, 11), (7, 11)):
        bs.column_dimensions[CL(col)].width = w
    bs.freeze_panes = "A5"
    bs.conditional_formatting.add("E5:E%d" % (tr2 - 1),
        ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="ED7D31"))
    bs.sheet_view.showGridLines = False

for pl in PLATFORMS:
    best_sellers(pl, [pl])
best_sellers("All Platforms", PLATFORMS)

# =====================================================================
# 8. Raw Data  (headers + the supplied export pre-loaded)
# =====================================================================
rd = wb.create_sheet("Raw Data")
HEADER_FALLBACK = os.path.join(HERE, "raw_header.txt")
src = CSVF if CSVF else HEADER_FALLBACK
with open(src, newline="", encoding="utf-8-sig") as fh:
    rdr = csv.reader(fh)
    header = next(rdr)
    for i, h in enumerate(header, start=1):
        c = rd.cell(row=1, column=i, value=h); c.font = F_HEAD; c.fill = FILL_H; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    IDX_DATE, IDX_SUB, IDX_ITEM = 13, 30, 31        # 0-based: Order Date, Suborder Qty, Item Qty
    n = 0
    for j, row_ in enumerate(rdr if CSVF else []):
        n += 1
        r = 2 + j
        for i, v in enumerate(row_):
            if v == "": continue
            if i in (IDX_SUB, IDX_ITEM):
                try: v = int(v)
                except ValueError: pass
            rd.cell(row=r, column=i + 1, value=v)
rd.row_dimensions[1].height = 42
rd.freeze_panes = "A2"
for i in range(1, len(header) + 1):
    rd.column_dimensions[CL(i)].width = 15
print("raw rows loaded:", n)

# =====================================================================
# 9. Work  (hidden helper)
# =====================================================================
wk = wb.create_sheet("Work")
for i, h in enumerate(["Date", "Platform", "SKU", "Units"], start=1):
    c = wk.cell(row=1, column=i, value=h); c.font = F_HEAD; c.fill = FILL_H
wk.cell(row=1, column=6, value="Helper sheet - do not edit. Normalises Raw Data. "
        "Raw Data columns used: C = MP Name, N = Order Date, AE = Suborder Quantity, AF = Item Quantity, AG = SKU."
        ).font = Font(italic=True, size=9, color="808080")
RD = lambda col: "INDEX('Raw Data'!$%s:$%s,ROW())" % (col, col)
# whole-column INDEX keeps these immune to rows being inserted or deleted in Raw Data
for r in range(2, WORK_LAST + 1):
    wk.cell(row=r, column=1, value=(
        '=IF({sku}="","",IF(ISNUMBER({dt}),INT({dt}),'
        'IFERROR(DATE(VALUE(LEFT({dt},4)),VALUE(MID({dt},6,2)),VALUE(MID({dt},9,2))),"")))'
        ).format(sku=RD("AG"), dt=RD("N")))
    wk.cell(row=r, column=2, value=(
        '=IF({sku}="","",IF(ISNUMBER(SEARCH("amazon",{mp})),"Amazon",'
        'IF(ISNUMBER(SEARCH("flipkart",{mp})),"Flipkart",'
        'IF(ISNUMBER(SEARCH("myntra",{mp})),"Myntra","Other"))))'
        ).format(sku=RD("AG"), mp=RD("C")))
    wk.cell(row=r, column=3, value='=IF({sku}="","",LOWER(TRIM({sku})))'.format(sku=RD("AG")))
    wk.cell(row=r, column=4, value=(
        '=IF({sku}="","",MAX(IFERROR({ae}+0,0),IFERROR({af}+0,0)))'
        ).format(sku=RD("AG"), ae=RD("AE"), af=RD("AF")))
for cc, w in (("A", 12), ("B", 12), ("C", 34), ("D", 8)):
    wk.column_dimensions[cc].width = w
wk.sheet_state = "hidden"

# =====================================================================
# named ranges
# =====================================================================
for nm, ref in [
    ("WDATE", "Work!$A$2:$A$%d" % WORK_LAST),
    ("WPLAT", "Work!$B$2:$B$%d" % WORK_LAST),
    ("WSKU",  "Work!$C$2:$C$%d" % WORK_LAST),
    ("WQTY",  "Work!$D$2:$D$%d" % WORK_LAST),
    ("DEND",   "Settings!$B$14"),
    ("DSTART", "Settings!$B$15"),
    ("SKULIST", "'SKU Master'!$A$4:$A$%d" % (3 + NS)),
]:
    wb.defined_names.add(DefinedName(nm, attr_text=ref))

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
