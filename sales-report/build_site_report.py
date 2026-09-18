"""
Build the website (Shopify / OMS) sales report workbook from the site orders export.

    python3 build_site_report.py <site_orders.csv> [output.xlsx]

The site export is one row per ORDER with up to 16 product blocks side by side
(Product SKU (n) / Product Quantity (n)), so quantities are read across all 16
blocks rather than down a single column. Every figure is a live formula reading
the 'Raw Data' sheet, so pasting a fresh export refreshes the whole report.
"""
import csv, json, os, re, sys
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as CL
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import ColorScaleRule

HERE = os.path.dirname(os.path.abspath(__file__))
CSVF = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] else None
OUT  = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "Website_Sales_Report.xlsx")
SKUF = os.path.join(HERE, "skus.txt")
UNLF = os.path.join(HERE, "site_unlisted_skus.json")
ALIF = os.path.join(HERE, "site_sku_aliases.json")

RAWCAP  = 25000            # Raw Data / Work rows 2..RAWCAP+1
NDAYS   = 15
NBLOCK  = 16               # product blocks in the export
PENDCAP = 3000             # capacity of the pending-order list
COL_DATE, COL_OID, COL_STATUS = 1, 2, 155          # A, B, EY
SKU_COL = lambda n: 38 + 7 * (n - 1)
QTY_COL = lambda n: 40 + 7 * (n - 1)
PENDING = ("created", "draft")

skus = [l.strip() for l in open(SKUF) if l.strip()]
NS = len(skus)
# {master_sku: [alternate codes counted into it]} - folds duplicate listings into one row
ALIAS = json.load(open(ALIF)) if os.path.exists(ALIF) else {}
_aliased = {a for v in ALIAS.values() for a in v}
unlisted = [u for u in (json.load(open(UNLF)) if os.path.exists(UNLF) else []) if u not in _aliased]
NU = len(unlisted)

NAVY="1F3864"; BLUE="2E5C8A"; LIGHT="DCE6F1"; ACCENT="FFF2CC"; GREEN="E2EFDA"; RUST="C55A11"
F_TITLE=Font(bold=True,size=14,color="FFFFFF"); F_HEAD=Font(bold=True,size=10,color="FFFFFF")
F_SUB=Font(bold=True,size=10,color=NAVY); F_B=Font(bold=True,size=10)
FILL_T=PatternFill("solid",fgColor=NAVY); FILL_H=PatternFill("solid",fgColor=BLUE)
FILL_L=PatternFill("solid",fgColor=LIGHT); FILL_A=PatternFill("solid",fgColor=ACCENT)
FILL_G=PatternFill("solid",fgColor=GREEN); FILL_R=PatternFill("solid",fgColor=RUST)
thin=Side(style="thin",color="BFBFBF"); BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)
CTR=Alignment(horizontal="center",vertical="center"); DATEFMT="dd-mmm-yy"
NOTE=Font(italic=True,size=9,color="808080")

def title_row(ws,text,width):
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=width)
    c=ws.cell(row=1,column=1,value=text); c.font=F_TITLE; c.fill=FILL_T
    c.alignment=Alignment(horizontal="left",vertical="center"); ws.row_dimensions[1].height=24

def band(ws,row,text,width):
    c=ws.cell(row=row,column=1,value=text); c.font=F_HEAD; c.fill=FILL_H
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=width)

# sum of one SUMIFS per product block, for each SKU code folded into this row
def blocks(crit, tail):
    return "+".join("SUMIFS(QTY_%d,SKU_%d,%s%s)" % (n, n, c, tail)
                    for c in ([crit] if isinstance(crit, str) else crit)
                    for n in range(1, NBLOCK + 1))

def codes_for(row_ref, sku_name):
    """Criteria for a sales row: its own cell, plus any alias codes as literals."""
    return [row_ref] + ['"%s"' % a for a in ALIAS.get(sku_name, [])]

def assert_name_safe(nm):
    """Excel silently reads a defined name that looks like a cell reference AS that
    reference, which turns every formula using it into #VALUE!. 'SKU1' and 'QTY1'
    are real cells (columns 13151 and 12037), so such names must never be used."""
    if re.fullmatch(r"[A-Za-z]{1,3}\d{1,7}", nm) or re.fullmatch(r"[Rr]\d+[Cc]\d+", nm):
        raise SystemExit("defined name %r looks like a cell reference - Excel will "
                         "misread it and every formula using it returns #VALUE!" % nm)

wb = Workbook()

# ---------------------------------------------------------------- README
ws = wb.active; ws.title = "README"
title_row(ws,"WEBSITE SALES REPORT  —  how to use",2)
lines=[
 ("",""),
 ("STEP 1  —  Paste the export",""),
 ("","Open 'Raw Data'. Select cell A2. Paste the site orders export (data rows only, WITHOUT the header row)."),
 ("","Header row 1 is already in place and must not be changed - every formula depends on those column positions."),
 ("","Before pasting fresh data clear the old rows: select row 2 down to the last filled row and press Delete."),
 ("",""),
 ("DATE FORMAT - READ THIS",""),
 ("","The export writes Order Date as dd/mm/yyyy text. The workbook reads both that text and real Excel dates,"),
 ("","so it works either way - PROVIDED your Excel is set to a dd/mm/yyyy locale (India / UK)."),
 ("","On a US (mm/dd/yyyy) locale Excel silently reads 05/09/2026 as 5 September instead of 9 May when it"),
 ("","converts on paste. Check 'Settings': if the earliest and latest dates look wrong, paste the Order Date"),
 ("","column as Text, or switch Windows regional format to English (India)."),
 ("",""),
 ("HOW QUANTITIES ARE READ",""),
 ("","The export is one row per ORDER, with up to 16 product blocks across the row:"),
 ("","  Product SKU (1) / Product Quantity (1) ... Product SKU (16) / Product Quantity (16)."),
 ("","All 16 blocks are read, so multi-item orders are counted in full."),
 ("","Every order counts regardless of Shipment Status - cancelled and RTO included - as with the marketplace report."),
 ("","Sales are booked on Order Date."),
 ("",""),
 ("THE 15-DAY WINDOW",""),
 ("","The workbook finds the LATEST order date in the pasted data, drops it (always a partial day),"),
 ("","and reports the 15 full days before it. An export covering 1 Sep - 18 Sep reports 3 Sep - 17 Sep."),
 ("","Change the length, the trailing days dropped, or force an exact end date on 'Settings'."),
 ("",""),
 ("PENDING ORDERS",""),
 ("","An order is PENDING when its Shipment Status (column EY) is 'created' or 'draft'."),
 ("","'Pending Orders' gives the order count and the SKU-wise quantity behind it."),
 ("","'Pending Order List' lists those orders one per row with their SKUs and quantities."),
 ("","Both cover EVERY pending order in the pasted data, not just the 15-day window - an older order still"),
 ("","needs shipping. The window-only figures sit beside them for reference."),
 ("",""),
 ("SKUs NOT IN YOUR LIST",""),
 ("","Some export rows carry SKUs that are not in your 182-SKU master - typos, retired codes, or a blank"),
 ("","placeholder. They are NOT folded into the 182 SKU rows. Instead:"),
 ("","  - 'Website Sales' lists them in a separate block below the TOTAL row;"),
 ("","  - 'Daily Orders' shows a per-day 'Not in SKU list' units column so nothing is ever lost silently."),
 ("","The unlisted block is seeded from the codes seen in the supplied export. If a NEW bad code appears it"),
 ("","will not get its own row, but it WILL show in the 'Not in SKU list' column - that number is the alarm."),
 ("",""),
 ("THE SHEETS",""),
 ("Settings","Date window controls and a live summary."),
 ("Daily Orders","Orders and units per day, with averages and the unlisted-SKU check."),
 ("Website Sales","Date-wise x SKU-wise units, 15-day Total and Average per Day, plus the unlisted block."),
 ("Best Sellers","All 182 SKUs ranked high to low for the window."),
 ("Pending Orders","Pending order count and the SKU-wise quantity behind it."),
 ("Pending Order List","One row per pending order with its SKUs and quantities."),
 ("SKU Master","The 182 SKUs driving every sheet."),
 ("Raw Data","The paste area."),
 ("Work","Hidden helper (date, status, pending counter). Do not edit."),
 ("",""),
 ("NOTES",""),
 ("","Every SKU is listed on every sheet. A SKU with no sales in the window shows 0, not a blank."),
 ("","Figures are UNITS (quantity). 'Daily Orders' also shows the order count beside them."),
 ("","'Raw Data' and 'Work' cover %s rows (~23 days at current volume). To paste more, copy the last" % "{:,}".format(RAWCAP)),
 ("","Work row further down and widen the named ranges."),
]
r=2
for a,b in lines:
    if a and not b:
        c=ws.cell(row=r,column=1,value=a); c.font=Font(bold=True,size=11,color=NAVY); c.fill=FILL_L
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    else:
        if a: ws.cell(row=r,column=1,value=a).font=F_B
        if b: ws.cell(row=r,column=2,value=b)
    r+=1
ws.column_dimensions["A"].width=24; ws.column_dimensions["B"].width=118
ws.sheet_view.showGridLines=False

# ---------------------------------------------------------------- Settings
st=wb.create_sheet("Settings")
title_row(st,"SETTINGS  —  reporting window",4)
SALES_TOT=5+NS+NU+1     # TOTAL row on Website Sales (set below to match)
rows=[
 (3,"DETECTED FROM RAW DATA",None,None,None),
 (4,"Earliest order date in Raw Data",'=IF(COUNT(WDATE)=0,"",MIN(WDATE))',DATEFMT,"Sanity-check these two against the export."),
 (5,"Latest order date in Raw Data",'=IF(COUNT(WDATE)=0,"",MAX(WDATE))',DATEFMT,"If they look wrong, see the date-format note on README."),
 (6,"Orders detected",'=COUNT(WDATE)',"#,##0",None),
 (7,"Orders whose date could not be read","=COUNTA('Raw Data'!$A$2:$A$%d)-COUNT(WDATE)"%(RAWCAP+1),"#,##0","Should be 0. Anything else means the date format was not understood."),
 (9,"YOU CAN CHANGE THESE",None,None,None),
 (10,"Number of days to report",NDAYS,"0","15 = a fortnight plus a day."),
 (11,"Trailing days to drop",1,"0","The export's last day is partial, so 1 drops it."),
 (12,"Manual END date (blank = automatic)",None,DATEFMT,"Fill only to pin the window to an exact end date."),
 (14,"WINDOW IN USE",None,None,None),
 (15,"END date",'=IF($B$12<>"",$B$12,IF(N($B$5)=0,"",$B$5-$B$11))',DATEFMT,None),
 (16,"START date",'=IF(N($B$15)=0,"",$B$15-$B$10+1)',DATEFMT,None),
 (18,"WINDOW TOTALS",None,None,None),
 (19,"Orders in window","=SUM('Daily Orders'!B5:B%d)"%(4+NDAYS),"#,##0",None),
 (20,"Units - tracked SKUs","=SUM('Daily Orders'!C5:C%d)"%(4+NDAYS),"#,##0",None),
 (21,"Units - not in SKU list","=SUM('Daily Orders'!E5:E%d)"%(4+NDAYS),"#,##0","Shown separately on Website Sales."),
 (22,"Units - all","=B20+B21","#,##0",None),
 (23,"Average orders per day",'=IF($B$10=0,0,B19/$B$10)',"#,##0.0",None),
 (24,"Average units per day",'=IF($B$10=0,0,B22/$B$10)',"#,##0.0",None),
 (25,"Average units per order",'=IF(B19=0,0,B22/B19)',"#,##0.00",None),
 (27,"PENDING (whole Raw Data, not window-limited)",None,None,None),
 (28,'Orders with status "created"','=COUNTIF(WSTAT,"created")',"#,##0",None),
 (29,'Orders with status "draft"','=COUNTIF(WSTAT,"draft")',"#,##0",None),
 (30,"TOTAL pending orders","=B28+B29","#,##0",None),
 (31,"Pending units - tracked SKUs","='Pending Orders'!D%d"%(7+NS+NU),"#,##0",None),
]
for rr,label,val,fmt,note in rows:
    if val is None and fmt is None and note is None:
        band(st,rr,label,4); continue
    st.cell(row=rr,column=1,value=label).font=F_B
    c=st.cell(row=rr,column=2,value=val)
    if fmt: c.number_format=fmt
    c.border=BORDER; c.alignment=CTR
    if rr in (10,11,12): c.fill=FILL_A
    if rr in (28,29,30,31): c.fill=PatternFill("solid",fgColor="FCE4D6")
    if note: st.cell(row=rr,column=3,value=note).font=NOTE
st.column_dimensions["A"].width=38; st.column_dimensions["B"].width=16
st.column_dimensions["C"].width=58; st.sheet_view.showGridLines=False

# ---------------------------------------------------------------- SKU Master
sm=wb.create_sheet("SKU Master")
title_row(sm,"SKU MASTER  —  the list every sheet is built from",4)
for i,h in enumerate(["SKU","Product","Colour","Size","Also counts (merged duplicate listings)"],start=1):
    c=sm.cell(row=3,column=i,value=h); c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER; c.alignment=CTR
SIZEMAP={"ss":"S","sm":"M","sl":"L","sxl":"XL","sxxl":"XXL","sxxxl":"3XL","s4xl":"4XL"}
PRODMAP={"jellybra":"JellyLift Bra","stayputstrapless":"StayPut Strapless","instatuck":"InstaTuck","widestr":"Wide Strap Bra"}
for i,s in enumerate(skus):
    p=s.split("-")
    sm.cell(row=4+i,column=1,value=s).border=BORDER
    sm.cell(row=4+i,column=2,value=PRODMAP.get(p[0],p[0])).border=BORDER
    sm.cell(row=4+i,column=3,value=p[2][1:] if len(p)>2 else "").border=BORDER
    c=sm.cell(row=4+i,column=4,value=SIZEMAP.get(p[3],p[3]) if len(p)>3 else ""); c.border=BORDER; c.alignment=CTR
    if s in ALIAS:
        c=sm.cell(row=4+i,column=5,value=", ".join(ALIAS[s])); c.border=BORDER
        c.font=Font(size=9); c.fill=PatternFill("solid",fgColor="E2EFDA")
        sm.cell(row=4+i,column=1).fill=PatternFill("solid",fgColor="E2EFDA")
sm.cell(row=2,column=1,value="Same 182 SKUs as the marketplace report. Green rows fold a duplicate site listing into the master SKU "
        "(edit site_sku_aliases.json to change).").font=NOTE
for cc,w in (("A",34),("B",20),("C",18),("D",8),("E",42)): sm.column_dimensions[cc].width=w
sm.freeze_panes="A4"; sm.auto_filter.ref="A3:E%d"%(3+NS); sm.sheet_view.showGridLines=False

# ---------------------------------------------------------------- Website Sales
sales=wb.create_sheet("Website Sales")
FIRST=5; LAST=4+NS; UFIRST=LAST+2; ULAST=UFIRST+NU-1; TOTR=ULAST+2
DCOL0=2; QCOL=17; RCOL=18; SCOL=20
title_row(sales,"WEBSITE  —  DATE-WISE / SKU-WISE SALES  (units)",RCOL)
sales.cell(row=2,column=1,value="Source:").font=F_B
c=sales.cell(row=2,column=2,value="Website"); c.font=F_SUB; c.fill=FILL_A; c.border=BORDER; c.alignment=CTR
sales.cell(row=2,column=4,value="All 16 product blocks are read. Every shipment status counts. Sales booked on Order Date.").font=NOTE
sales.cell(row=3,column=1,value="Window:").font=F_B
c=sales.cell(row=3,column=2,value="=DSTART"); c.number_format=DATEFMT; c.font=F_B; c.alignment=CTR
sales.cell(row=3,column=3,value="to").alignment=CTR
c=sales.cell(row=3,column=4,value="=DEND"); c.number_format=DATEFMT; c.font=F_B; c.alignment=CTR
sales.cell(row=3,column=5,value="Days:").font=F_B
c=sales.cell(row=3,column=6,value="=Settings!$B$10"); c.font=F_B; c.alignment=CTR
c=sales.cell(row=4,column=1,value="SKU"); c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER; c.alignment=CTR
for d in range(NDAYS):
    col=DCOL0+d
    c=sales.cell(row=4,column=col,value='=IF(N(DSTART)=0,"",DSTART)' if d==0 else '=IF(N(%s4)=0,"",%s4+1)'%(CL(col-1),CL(col-1)))
    c.number_format="dd-mmm"; c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER
    c.alignment=Alignment(horizontal="center",vertical="center",textRotation=90)
    sales.column_dimensions[CL(col)].width=7
for col,lbl in ((QCOL,"Total (window)"),(RCOL,"Avg / Day")):
    c=sales.cell(row=4,column=col,value=lbl); c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
sales.row_dimensions[4].height=62

def sku_row(r, name_formula, tie_row, sku_name=None):
    c=sales.cell(row=r,column=1,value=name_formula); c.border=BORDER; c.font=Font(size=9)
    crit=codes_for("$A%d"%r, sku_name)
    for d in range(NDAYS):
        col=CL(DCOL0+d)
        c=sales.cell(row=r,column=DCOL0+d,value="="+blocks(crit,",WDATE,%s$4"%col))
        c.border=BORDER; c.number_format='0;;"0"'; c.font=Font(size=9); c.alignment=CTR
    c=sales.cell(row=r,column=QCOL,value="=SUM(%s%d:%s%d)"%(CL(DCOL0),r,CL(DCOL0+NDAYS-1),r))
    c.border=BORDER; c.font=F_B; c.fill=FILL_L; c.number_format="#,##0"; c.alignment=CTR
    c=sales.cell(row=r,column=RCOL,value='=IF($F$3=0,0,%s%d/$F$3)'%(CL(QCOL),r))
    c.border=BORDER; c.number_format="#,##0.00"; c.fill=FILL_L; c.alignment=CTR
    if tie_row is not None:
        sales.cell(row=r,column=SCOL,value="=%s%d+(1000-ROW())/1000000"%(CL(QCOL),r))

for i in range(NS):
    r=FIRST+i
    sku_row(r,"='SKU Master'!$A$%d"%(4+i),True,skus[i])
    if skus[i] in ALIAS:
        sales.cell(row=r,column=1).fill=PatternFill("solid",fgColor="E2EFDA")
        sales.cell(row=r,column=1).comment=None
band(sales,LAST+1,"SKUs FOUND IN THE EXPORT THAT ARE NOT IN YOUR 182-SKU LIST  —  reported separately, never folded in above",RCOL)
for i,u in enumerate(unlisted):
    r=UFIRST+i
    sku_row(r,u,None,None)
    sales.cell(row=r,column=1).fill=PatternFill("solid",fgColor="FCE4D6")
c=sales.cell(row=TOTR,column=1,value="TOTAL — tracked + unlisted"); c.font=F_HEAD; c.fill=FILL_T; c.border=BORDER
for col in list(range(DCOL0,DCOL0+NDAYS))+[QCOL,RCOL]:
    L=CL(col)
    c=sales.cell(row=TOTR,column=col,value="=SUM(%s%d:%s%d)+SUM(%s%d:%s%d)"%(L,FIRST,L,LAST,L,UFIRST,L,ULAST))
    c.font=F_HEAD; c.fill=FILL_T; c.border=BORDER; c.alignment=CTR
    c.number_format="#,##0.00" if col==RCOL else "#,##0"
c=sales.cell(row=TOTR+1,column=1,value="of which: tracked SKUs only"); c.font=F_B; c.fill=FILL_G; c.border=BORDER
for col in list(range(DCOL0,DCOL0+NDAYS))+[QCOL,RCOL]:
    L=CL(col)
    c=sales.cell(row=TOTR+1,column=col,value="=SUM(%s%d:%s%d)"%(L,FIRST,L,LAST))
    c.font=F_B; c.fill=FILL_G; c.border=BORDER; c.alignment=CTR
    c.number_format="#,##0.00" if col==RCOL else "#,##0"
sales.column_dimensions["A"].width=34
sales.column_dimensions[CL(QCOL)].width=11; sales.column_dimensions[CL(RCOL)].width=11
sales.column_dimensions[CL(SCOL)].hidden=True
sales.freeze_panes="B5"
sales.conditional_formatting.add("%s%d:%s%d"%(CL(DCOL0),FIRST,CL(DCOL0+NDAYS-1),LAST),
    ColorScaleRule(start_type="num",start_value=0,start_color="FFFFFF",end_type="max",end_color="4472C4"))
sales.sheet_view.showGridLines=False
SALES_TOTAL_ROW=TOTR

# ---------------------------------------------------------------- Daily Orders
do=wb.create_sheet("Daily Orders",2)
title_row(do,"DAILY ORDERS  —  website volume by date",8)
do.cell(row=2,column=1,value="Orders = rows in the export for that date. Units = quantity across all 16 product blocks.").font=NOTE
heads=["Date","Orders","Units\n(tracked SKUs)","Units\n(all SKUs)","Not in\nSKU list","Units\nper order","Day of\nWeek","Pending orders\n(created+draft)"]
for i,h in enumerate(heads,start=1):
    c=do.cell(row=4,column=i,value=h); c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
do.row_dimensions[4].height=34
for d in range(NDAYS):
    r=5+d
    c=do.cell(row=r,column=1,value='=IF(N(DSTART)=0,"",DSTART)' if d==0 else '=IF(N(A%d)=0,"",A%d+1)'%(r-1,r-1))
    c.number_format=DATEFMT; c.border=BORDER; c.font=F_B; c.alignment=CTR
    c=do.cell(row=r,column=2,value="=COUNTIFS(WDATE,$A%d)"%r); c.border=BORDER; c.number_format="#,##0"; c.alignment=CTR
    c=do.cell(row=r,column=3,value="='Website Sales'!%s%d"%(CL(DCOL0+d),SALES_TOTAL_ROW+1))
    c.border=BORDER; c.number_format="#,##0"; c.font=F_B; c.fill=FILL_L; c.alignment=CTR
    c=do.cell(row=r,column=4,value="="+blocks('"<>"',',WDATE,$A%d'%r)); c.border=BORDER; c.number_format="#,##0"; c.alignment=CTR
    c=do.cell(row=r,column=5,value="=D%d-C%d"%(r,r)); c.border=BORDER; c.number_format="#,##0"; c.alignment=CTR; c.font=Font(size=9,color="808080")
    c=do.cell(row=r,column=6,value="=IF(B%d=0,0,D%d/B%d)"%(r,r,r)); c.border=BORDER; c.number_format="#,##0.00"; c.alignment=CTR
    c=do.cell(row=r,column=7,value='=IF(N($A%d)=0,"",TEXT($A%d,"ddd"))'%(r,r)); c.border=BORDER; c.alignment=CTR; c.font=Font(size=9,color="808080")
    c=do.cell(row=r,column=8,value='=COUNTIFS(WDATE,$A%d,WSTAT,"created")+COUNTIFS(WDATE,$A%d,WSTAT,"draft")'%(r,r))
    c.border=BORDER; c.number_format="#,##0"; c.alignment=CTR; c.font=Font(size=9,color=RUST)
tr=5+NDAYS
c=do.cell(row=tr,column=1,value="TOTAL"); c.font=F_HEAD; c.fill=FILL_T; c.border=BORDER
for col in range(2,9):
    L=CL(col)
    v="=IF(B%d=0,0,D%d/B%d)"%(tr,tr,tr) if col==6 else ("" if col==7 else "=SUM(%s5:%s%d)"%(L,L,tr-1))
    if col==7: continue
    c=do.cell(row=tr,column=col,value=v); c.font=F_HEAD; c.fill=FILL_T; c.border=BORDER; c.alignment=CTR
    c.number_format="#,##0.00" if col==6 else "#,##0"
ar=tr+1
c=do.cell(row=ar,column=1,value="AVERAGE PER DAY"); c.font=F_B; c.fill=FILL_G; c.border=BORDER
for col in (2,3,4,5,8):
    L=CL(col)
    c=do.cell(row=ar,column=col,value="=IF(Settings!$B$10=0,0,%s%d/Settings!$B$10)"%(L,tr))
    c.font=F_B; c.fill=FILL_G; c.border=BORDER; c.number_format="#,##0.0"; c.alignment=CTR
for col,w in ((1,13),(2,11),(3,15),(4,13),(5,12),(6,11),(7,9),(8,16)): do.column_dimensions[CL(col)].width=w
do.freeze_panes="B5"
do.conditional_formatting.add("B5:B%d"%(tr-1),ColorScaleRule(start_type="min",start_color="FFFFFF",end_type="max",end_color="70AD47"))
do.sheet_view.showGridLines=False

# ---------------------------------------------------------------- Best Sellers
bs=wb.create_sheet("Best Sellers")
title_row(bs,"BEST SELLERS  —  WEBSITE  (all %d SKUs, ranked)"%NS,7)
bs.cell(row=2,column=1,value="Ranked on units in the reporting window. Re-ranks itself whenever new data is pasted. Unlisted SKUs are excluded.").font=NOTE
bs.cell(row=3,column=1,value="Window:").font=F_B
c=bs.cell(row=3,column=2,value="=DSTART"); c.number_format=DATEFMT; c.font=F_B; c.alignment=CTR
bs.cell(row=3,column=3,value="to").alignment=CTR
c=bs.cell(row=3,column=4,value="=DEND"); c.number_format=DATEFMT; c.font=F_B; c.alignment=CTR
for i,h in enumerate(["Rank","SKU","Product","Size","Units\n(window)","Avg / Day","% of total"],start=1):
    c=bs.cell(row=4,column=i,value=h); c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
bs.row_dimensions[4].height=30
SCORE="'Website Sales'!$%s$%d:$%s$%d"%(CL(SCOL),FIRST,CL(SCOL),LAST)
SRCSKU="'Website Sales'!$A$%d:$A$%d"%(FIRST,LAST)
TOTREF="'Website Sales'!$%s$%d"%(CL(QCOL),SALES_TOTAL_ROW+1)
for i in range(NS):
    r=5+i
    c=bs.cell(row=r,column=1,value=i+1); c.border=BORDER; c.alignment=CTR; c.font=F_B
    c=bs.cell(row=r,column=2,value="=INDEX(%s,MATCH(LARGE(%s,$A%d),%s,0))"%(SRCSKU,SCORE,r,SCORE)); c.border=BORDER; c.font=Font(size=9)
    c=bs.cell(row=r,column=3,value='=IFERROR(INDEX(\'SKU Master\'!$B$4:$B$%d,MATCH($B%d,SKULIST,0)),"")'%(3+NS,r)); c.border=BORDER; c.font=Font(size=9)
    c=bs.cell(row=r,column=4,value='=IFERROR(INDEX(\'SKU Master\'!$D$4:$D$%d,MATCH($B%d,SKULIST,0)),"")'%(3+NS,r)); c.border=BORDER; c.font=Font(size=9); c.alignment=CTR
    c=bs.cell(row=r,column=5,value="=INT(LARGE(%s,$A%d))"%(SCORE,r)); c.border=BORDER; c.number_format="#,##0"; c.font=F_B; c.fill=FILL_L; c.alignment=CTR
    c=bs.cell(row=r,column=6,value="=IF(Settings!$B$10=0,0,$E%d/Settings!$B$10)"%r); c.border=BORDER; c.number_format="#,##0.00"; c.alignment=CTR
    c=bs.cell(row=r,column=7,value="=IF(%s=0,0,$E%d/%s)"%(TOTREF,r,TOTREF)); c.border=BORDER; c.number_format="0.0%"; c.alignment=CTR
tr2=5+NS
c=bs.cell(row=tr2,column=1,value="TOTAL"); c.font=F_HEAD; c.fill=FILL_T; c.border=BORDER
bs.merge_cells(start_row=tr2,start_column=1,end_row=tr2,end_column=4)
for col,fmt in ((5,"#,##0"),(6,"#,##0.00"),(7,"0.0%")):
    c=bs.cell(row=tr2,column=col,value="=SUM(%s5:%s%d)"%(CL(col),CL(col),tr2-1))
    c.font=F_HEAD; c.fill=FILL_T; c.border=BORDER; c.number_format=fmt; c.alignment=CTR
for col,w in ((1,7),(2,34),(3,20),(4,8),(5,11),(6,11),(7,11)): bs.column_dimensions[CL(col)].width=w
bs.freeze_panes="A5"
bs.conditional_formatting.add("E5:E%d"%(tr2-1),ColorScaleRule(start_type="min",start_color="FFFFFF",end_type="max",end_color="ED7D31"))
bs.sheet_view.showGridLines=False

# ---------------------------------------------------------------- Pending Orders
po=wb.create_sheet("Pending Orders")
title_row(po,"PENDING ORDERS  —  Shipment Status 'created' or 'draft'",6)
po.cell(row=2,column=1,value="Covers EVERY pending order in Raw Data, not only the 15-day window - an older unshipped order still needs picking.").font=NOTE
summ=[(3,'Orders with status "created"','=COUNTIF(WSTAT,"created")'),
      (4,'Orders with status "draft"','=COUNTIF(WSTAT,"draft")'),
      (5,'TOTAL PENDING ORDERS','=B3+B4')]
for rr,lbl,val in summ:
    c=po.cell(row=rr,column=1,value=lbl); c.font=F_B if rr<5 else Font(bold=True,size=11,color="FFFFFF")
    if rr==5: c.fill=FILL_R
    c=po.cell(row=rr,column=2,value=val); c.number_format="#,##0"; c.border=BORDER; c.alignment=CTR
    c.font=F_B if rr<5 else Font(bold=True,size=12,color="FFFFFF")
    if rr==5: c.fill=FILL_R
    else: c.fill=PatternFill("solid",fgColor="FCE4D6")
PR0=7
for i,h in enumerate(["SKU",'Units\n"created"','Units\n"draft"',"TOTAL\npending units","Product","Size"],start=1):
    c=po.cell(row=PR0-1,column=i,value=h); c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
po.row_dimensions[PR0-1].height=32
allsk=[("='SKU Master'!$A$%d"%(4+i),False,skus[i]) for i in range(NS)]+[(u,True,None) for u in unlisted]
for i,(nm,isu,sname) in enumerate(allsk):
    r=PR0+i
    c=po.cell(row=r,column=1,value=nm); c.border=BORDER; c.font=Font(size=9)
    if isu: c.fill=PatternFill("solid",fgColor="FCE4D6")
    elif sname in ALIAS: c.fill=PatternFill("solid",fgColor="E2EFDA")
    crit=codes_for("$A%d"%r, sname)
    c=po.cell(row=r,column=2,value="="+blocks(crit,',WSTAT,"created"')); c.border=BORDER; c.number_format='0;;"0"'; c.alignment=CTR; c.font=Font(size=9)
    c=po.cell(row=r,column=3,value="="+blocks(crit,',WSTAT,"draft"')); c.border=BORDER; c.number_format='0;;"0"'; c.alignment=CTR; c.font=Font(size=9)
    c=po.cell(row=r,column=4,value="=B%d+C%d"%(r,r)); c.border=BORDER; c.number_format="#,##0"; c.font=F_B; c.fill=FILL_L; c.alignment=CTR
    c=po.cell(row=r,column=5,value='=IFERROR(INDEX(\'SKU Master\'!$B$4:$B$%d,MATCH($A%d,SKULIST,0)),"not in SKU list")'%(3+NS,r)); c.border=BORDER; c.font=Font(size=9)
    c=po.cell(row=r,column=6,value='=IFERROR(INDEX(\'SKU Master\'!$D$4:$D$%d,MATCH($A%d,SKULIST,0)),"")'%(3+NS,r)); c.border=BORDER; c.font=Font(size=9); c.alignment=CTR
ptr=PR0+len(allsk)
c=po.cell(row=ptr,column=1,value="TOTAL"); c.font=F_HEAD; c.fill=FILL_T; c.border=BORDER
for col in (2,3,4):
    L=CL(col)
    c=po.cell(row=ptr,column=col,value="=SUM(%s%d:%s%d)"%(L,PR0,L,ptr-1))
    c.font=F_HEAD; c.fill=FILL_T; c.border=BORDER; c.number_format="#,##0"; c.alignment=CTR
for col,w in ((1,34),(2,12),(3,12),(4,14),(5,20),(6,8)): po.column_dimensions[CL(col)].width=w
po.freeze_panes="A7"; po.auto_filter.ref="A%d:F%d"%(PR0-1,ptr-1); po.sheet_view.showGridLines=False

# ---------------------------------------------------------------- Pending Order List
pl=wb.create_sheet("Pending Order List")
NSHOW=4
WID=5+NSHOW*2
title_row(pl,"PENDING ORDER LIST  —  one row per pending order",WID)
pl.cell(row=2,column=1,value="Every order whose Shipment Status is 'created' or 'draft', oldest row first. "
        "The first %d product lines are shown; 'Lines' and 'Total Units' always cover all 16."%NSHOW).font=NOTE
c=pl.cell(row=3,column=1,value="Pending orders:"); c.font=F_B
c=pl.cell(row=3,column=2,value="=COUNTIF(WSTAT,\"created\")+COUNTIF(WSTAT,\"draft\")")
c.font=Font(bold=True,size=12,color="FFFFFF"); c.fill=FILL_R; c.border=BORDER; c.alignment=CTR; c.number_format="#,##0"
pl.cell(row=3,column=4,value="List capacity: %s rows."%"{:,}".format(PENDCAP)).font=NOTE
heads=["#","Order Date","Order ID","Status","Total\nUnits"]
for n in range(1,NSHOW+1): heads+=["SKU (%d)"%n,"Qty\n(%d)"%n]
for i,h in enumerate(heads,start=1):
    c=pl.cell(row=4,column=i,value=h); c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
pl.row_dimensions[4].height=30
HID=WID+2   # hidden: matched raw row number
for k in range(1,PENDCAP+1):
    r=4+k
    pl.cell(row=r,column=HID,value='=IFERROR(MATCH(%d,WPEND,0)+1,"")'%k)
    rr="$%s%d"%(CL(HID),r)
    c=pl.cell(row=r,column=1,value='=IF(%s="","",%d)'%(rr,k)); c.border=BORDER; c.alignment=CTR; c.font=Font(size=9)
    c=pl.cell(row=r,column=2,value='=IF(%s="","",INDEX(WDATE,%s-1))'%(rr,rr))
    c.border=BORDER; c.number_format=DATEFMT; c.alignment=CTR; c.font=Font(size=9)
    c=pl.cell(row=r,column=3,value='=IF(%s="","",INDEX(\'Raw Data\'!$%s:$%s,%s))'%(rr,CL(COL_OID),CL(COL_OID),rr))
    c.border=BORDER; c.font=Font(size=9)
    c=pl.cell(row=r,column=4,value='=IF(%s="","",INDEX(WSTAT,%s-1))'%(rr,rr)); c.border=BORDER; c.alignment=CTR; c.font=Font(size=9)
    tot="+".join('N(INDEX(\'Raw Data\'!$%s:$%s,%s))'%(CL(QTY_COL(n)),CL(QTY_COL(n)),rr) for n in range(1,NBLOCK+1))
    c=pl.cell(row=r,column=5,value='=IF(%s="","",%s)'%(rr,tot))
    c.border=BORDER; c.number_format="#,##0"; c.font=F_B; c.fill=FILL_L; c.alignment=CTR
    for n in range(1,NSHOW+1):
        cs=6+(n-1)*2
        c=pl.cell(row=r,column=cs,value='=IF(%s="","",INDEX(\'Raw Data\'!$%s:$%s,%s))'%(rr,CL(SKU_COL(n)),CL(SKU_COL(n)),rr))
        c.border=BORDER; c.font=Font(size=8)
        c=pl.cell(row=r,column=cs+1,value='=IF(%s="","",INDEX(\'Raw Data\'!$%s:$%s,%s))'%(rr,CL(QTY_COL(n)),CL(QTY_COL(n)),rr))
        c.border=BORDER; c.alignment=CTR; c.font=Font(size=8); c.number_format="0;;"
pl.column_dimensions["A"].width=6; pl.column_dimensions["B"].width=12
pl.column_dimensions["C"].width=16; pl.column_dimensions["D"].width=10; pl.column_dimensions["E"].width=9
for n in range(1,NSHOW+1):
    pl.column_dimensions[CL(6+(n-1)*2)].width=30; pl.column_dimensions[CL(7+(n-1)*2)].width=6
pl.column_dimensions[CL(HID)].hidden=True
pl.freeze_panes="A5"; pl.auto_filter.ref="A4:%s%d"%(CL(WID),4+PENDCAP); pl.sheet_view.showGridLines=False

# ---------------------------------------------------------------- Raw Data
rd=wb.create_sheet("Raw Data")
src=CSVF if CSVF else os.path.join(HERE,"site_raw_header.txt")
with open(src,newline="",encoding="utf-8-sig") as fh:
    rdr=csv.reader(fh); header=next(rdr)
    for i,h in enumerate(header,start=1):
        c=rd.cell(row=1,column=i,value=h); c.font=F_HEAD; c.fill=FILL_H; c.border=BORDER
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    QCOLS={QTY_COL(n) for n in range(1,NBLOCK+1)}
    n=0
    for j,row_ in enumerate(rdr if CSVF else []):
        n+=1
        for i,v in enumerate(row_):
            if v=="": continue
            if (i+1) in QCOLS:
                try: v=int(v)
                except ValueError: pass
            rd.cell(row=2+j,column=i+1,value=v)
rd.row_dimensions[1].height=42; rd.freeze_panes="A2"
for i in range(1,len(header)+1): rd.column_dimensions[CL(i)].width=15
print("raw rows loaded:",n)
if n>RAWCAP: print("!! WARNING: %d rows exceeds capacity %d"%(n,RAWCAP))

# ---------------------------------------------------------------- Work
wk=wb.create_sheet("Work")
for i,h in enumerate(["Order Date","Shipment Status","Pending counter"],start=1):
    c=wk.cell(row=1,column=i,value=h); c.font=F_HEAD; c.fill=FILL_H
wk.cell(row=1,column=5,value="Helper sheet - do not edit. Raw Data columns used: A = Order Date, B = Order ID, "
        "EY = Shipment Status, and the 16 Product SKU / Product Quantity blocks.").font=NOTE
RD=lambda col:"INDEX('Raw Data'!$%s:$%s,ROW())"%(col,col)
D=RD(CL(COL_DATE)); S=RD(CL(COL_STATUS))
for r in range(2,RAWCAP+2):
    wk.cell(row=r,column=1,value=(
        '=IF({d}="","",IF(ISNUMBER({d}),INT({d}),'
        'IFERROR(DATE(VALUE(MID({d},7,4)),VALUE(MID({d},4,2)),VALUE(MID({d},1,2))),"")))').format(d=D))
    wk.cell(row=r,column=2,value='=IF({d}="","",LOWER(TRIM({s})))'.format(d=D,s=S))
    wk.cell(row=r,column=3,value='=IF(OR($B{r}="created",$B{r}="draft"),C{p}+1,C{p})'.format(r=r,p=r-1))
wk.cell(row=1,column=3).value="Pending counter"
wk["C1"]=0; wk["C1"].font=F_HEAD; wk["C1"].fill=FILL_H
for cc,w in (("A",12),("B",18),("C",14)): wk.column_dimensions[cc].width=w
wk.sheet_state="hidden"

# ---------------------------------------------------------------- names
names=[("WDATE","Work!$A$2:$A$%d"%(RAWCAP+1)),
       ("WSTAT","Work!$B$2:$B$%d"%(RAWCAP+1)),
       ("WPEND","Work!$C$2:$C$%d"%(RAWCAP+1)),
       ("DEND","Settings!$B$15"),("DSTART","Settings!$B$16"),
       ("SKULIST","'SKU Master'!$A$4:$A$%d"%(3+NS))]
for nn in range(1,NBLOCK+1):
    names.append(("SKU_%d"%nn,"'Raw Data'!$%s$2:$%s$%d"%(CL(SKU_COL(nn)),CL(SKU_COL(nn)),RAWCAP+1)))
    names.append(("QTY_%d"%nn,"'Raw Data'!$%s$2:$%s$%d"%(CL(QTY_COL(nn)),CL(QTY_COL(nn)),RAWCAP+1)))
for nm,ref in names:
    assert_name_safe(nm)
    wb.defined_names.add(DefinedName(nm,attr_text=ref))

wb.calculation.fullCalcOnLoad=True
wb.save(OUT)
print("saved",OUT)
