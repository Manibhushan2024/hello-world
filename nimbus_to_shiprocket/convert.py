#!/usr/bin/env python3
"""
Nimbus Post -> Shiprocket order-report converter.

Converts a Nimbus Post "order_b2c_report" CSV into the column layout of a
Shiprocket "secure_..._reports" CSV, and (optionally) appends the converted rows
below an existing Shiprocket report so both carriers' data live in one file.

Key differences handled:
  * Nimbus has ONE row per order with up to 10 product columns
    ("Product SKU (1)" ... "Product SKU (10)").  Shiprocket has ONE row per
    product line.  Every Nimbus order is exploded into one row per product,
    exactly like Shiprocket does.
  * Dates: Nimbus uses dd/mm/yyyy, Shiprocket uses "YYYY-MM-DD HH:MM:SS".
  * Statuses: Nimbus "in_transit" style -> Shiprocket "IN TRANSIT" style.
  * Zones: Nimbus "b" -> Shiprocket "z_b".
  * Payment method: Nimbus "COD"/"Prepaid" -> Shiprocket "cod"/"prepaid".
  * Order Total: Nimbus has no order-total column; it is computed as
      sum(qty * unit price) - Total Discount + Shipping + COD + Other charges
    which matches Nimbus' own "Collectable Amount" on every COD order.
  * Dimensions: three Nimbus columns -> Shiprocket "LxBxH".

Usage:
    python convert.py --nimbus nimbus.csv --shiprocket shiprocket.csv -o combined.csv
    python convert.py --nimbus nimbus.csv -o nimbus_as_shiprocket.csv     # convert only
    python convert.py --nimbus nimbus.csv --shiprocket shiprocket.csv --in-place

Both inputs may be .csv or .xlsx (the format Nimbus/Shiprocket download as).
Only the Python standard library is used; no installs needed.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

# --------------------------------------------------------------------------- #
# Shiprocket report layout (118 columns, in order).  Taken verbatim from a real
# Shiprocket export so the converted rows line up with an existing report.
# --------------------------------------------------------------------------- #
SHIPROCKET_COLUMNS = [
    "Order ID", "Forward ID", "Shiprocket Created At", "Channel", "Status",
    "Channel SKU", "Master SKU", "Product Name", "Product Category",
    "Product Quantity", "Channel Created At", "Customer Name", "Customer Email",
    "Customer Mobile", "Customer Alternate Phone", "Address Line 1",
    "Address Line 2", "Address City", "Address State", "Address Pincode",
    "Payment Method", "Product Price", "Order Total", "Tax", "Tax %",
    "Discount Value", "Product HSN", "Weight (KG)", "Dimensions (CM)",
    "Charged Weight", "Courier Company", "AWB Code", "SRX Premium LM AWB",
    "Shipping Bill URL", "AWB Assigned Date", "Manifested Date",
    "Pickup Location ID", "Pickup Address Name", "Pickup Scheduled Date",
    "Order Picked Up Date", "Pickup First Attempt Date", "Pickedup Timestamp",
    "Order Shipped Date", "EDD", "Delayed Reason", "Order Delivered Date",
    "RTO Address", "RTO Initiated Date", "RTO Delivered Date",
    "COD Remittance Date", "COD Payble Amount", "Remitted Amount", "CRF ID",
    "UTR No", "Zone", "COD Charges", "Freight Total Amount",
    "Customer_invoice_id", "Shipping Charges", "Pickup Exception Reason",
    "NPR1 Date", "NPR1 Reason", "NPR2 Date", "NPR2 Reason", "Latest NPR Date",
    "Latest NPR Reason", "First Out For Delivery Date",
    "First_Pickup_Scheduled_Date", "Buyer's Lat/long", "Order Type",
    "Order Tags", "Invoice Date", "Pickup Code", "Eway Bill Nos",
    "Last Updated AT", "Partial COD Collected", "Partial COD Value",
    "RTO Score Charges", "RTO Score Tax", "Delivery Boost Charges",
    "Delivery Boost Tax", "WhatsApp Tracking Charges", "WhatsApp Tracking Tax",
    "Brand Boost Charges", "Brand Boost Tax", "NDR 1 Attempt Date",
    "NDR 1 Remark", "NDR 2 Attempt Date", "NDR 2 Remark", "NDR 3 Attempt Date",
    "NDR 3 Remark", "Latest NDR Date", "Latest NDR Reason",
    "Bridge Call Recording", "Hub Address", "RAD Score", "RAD Datetimestamp",
    "BAG ID", "Pickup Pincode", "Verifier User ID", "Verifier Email",
    "Verifier Name", "Attempt Count", "Pickup Generated Date", "RTO WayBill",
    "RTO Risk", "RTO Reason", "Order Risk", "Address Risk", "Address Score",
    "Lost Date", "Latest OFD Date", "Master Courier", "Is Reverse",
    "Promise EDD", "Updated New EDD", "Exchange Order Type",
    "Cancellation Reason",
]

# Nimbus Post order_b2c_report layout (145 columns).  Used to recover the
# column names when a header cell has been edited by hand in Excel.
NIMBUS_COLUMNS = [
    "Order Date", "Order ID", "Channel Name", "Pickup Address ID",
    "Pickup Warehouse Address", "RTO Warehouse Address",
    "Is Document (Yes/No)", "Payment Method (COD/Prepaid)",
    "Buyer's Full Name", "Buyer's Email", "Buyer's Mobile No.",
    "Shipping Complete Address", "Shipping Address Landmark",
    "Shipping Address State", "Shipping Address City",
    "Shipping Address Pincode", "Billing Full Name", "Billing Email",
    "Billing Mobile No.", "Billing Complete Address", "Billing Landmark",
    "Billing State", "Billing City", "Billing Pincode", "Order Tags",
    "Reseller Name", "Shipment Weight (Kgs)", "Shipment Length (cm)",
    "Shipment Breadth (cm)", "Shipment Height (cm)", "Partial COD (Yes/No)",
    "Collectable Amount", "Shipping Charges", "COD Charges", "Total Discount",
    "Other Charges", "Verified Order (Yes/No)", "Product SKU (1)",
    "Product Name (1)*", "Product Quantity (1)*", "Product Unit Price (1)*",
    "Product Discount (1)", "Product HSN Code (1)", "Product Tax % (1)",
    "Product SKU (2)", "Product Name (2)", "Product Quantity (2)",
    "Product Unit Price (2)", "Product Discount (2)", "Product HSN Code (2)",
    "Product Tax % (2)", "Product SKU (3)", "Product Name (3)",
    "Product Quantity (3)", "Product Unit Price (3)", "Product Discount (3)",
    "Product HSN Code (3)", "Product Tax % (3)", "Product SKU (4)",
    "Product Name (4)", "Product Quantity (4)", "Product Unit Price (4)",
    "Product Discount (4)", "Product HSN Code (4)", "Product Tax % (4)",
    "Product SKU (5)", "Product Name (5)", "Product Quantity (5)",
    "Product Unit Price (5)", "Product Discount (5)", "Product HSN Code (5)",
    "Product Tax % (5)", "Product SKU (6)", "Product Name (6)",
    "Product Quantity (6)", "Product Unit Price (6)", "Product Discount (6)",
    "Product HSN Code (6)", "Product Tax % (6)", "Product SKU (7)",
    "Product Name (7)", "Product Quantity (7)", "Product Unit Price (7)",
    "Product Discount (7)", "Product HSN Code (7)", "Product Tax % (7)",
    "Product SKU (8)", "Product Name (8)", "Product Quantity (8)",
    "Product Unit Price (8)", "Product Discount (8)", "Product HSN Code (8)",
    "Product Tax % (8)", "Product SKU (9)", "Product Name (9)",
    "Product Quantity (9)", "Product Unit Price (9)", "Product Discount (9)",
    "Product HSN Code (9)", "Product Tax % (9)", "Product SKU (10)",
    "Product Name (10)", "Product Quantity (10)", "Product Unit Price (10)",
    "Product Discount (10)", "Product HSN Code (10)", "Product Tax % (10)",
    "Courier Name", "Courier Assigned Date", "AWB", "Zone", "Shipment Status",
    "Picked Date", "Shipped Date", "EDD", "Delivered Date",
    "RTO Delivered Date", "Weight Slab",
    "Total Freight Charges (inclusive of COD)", "Invoice ID", "Charged Weight",
    "Remittance ID", "Remitted Date", "Pickup Warehouse Nickname",
    "Contact Person Name", "Contact Number", "Email Address",
    "Complete address", "Landmark", "Pincode", "City", "State", "Country",
    "Is Primary (Yes/No)", "Is RTO warehouse same (Yes/No)",
    "RTO Warehouse Nickname", "RTO Warehouse Contact Person Name",
    "RTO Warehouse Contact Number", "RTO Warehouse Email Address",
    "RTO Warehouse Complete address", "RTO Warehouse Landmark",
    "RTO Warehouse Pincode", "RTO Warehouse City", "RTO Warehouse State",
    "RTO Warehouse Country",
]

# Nimbus shipment_status -> Shiprocket Status.
STATUS_MAP = {
    "created": "NEW ORDER",
    "pickup_pending": "PICKUP SCHEDULED",
    "pickup_scheduled": "PICKUP SCHEDULED",
    "out_for_pickup": "OUT FOR PICKUP",
    "pickup_failed": "PICKUP EXCEPTION",
    "pickup_completed": "SHIPPED",
    "in_transit": "IN TRANSIT",
    "reached_at_destination": "REACHED DESTINATION HUB",
    "out_for_delivery": "OUT FOR DELIVERY",
    "ndr": "UNDELIVERED",
    "delivered": "DELIVERED",
    "rto_initiated": "RTO INITIATED",
    "rto_in_transit": "RTO IN TRANSIT",
    "rto_delivered": "RTO DELIVERED",
    "cancelled": "CANCELED",
    "cancellation_requested": "CANCELLATION REQUESTED",
    "lost": "LOST",
    "damaged": "DAMAGED",
    "exception": "EXCEPTION",
}

# First word(s) of a Nimbus courier name -> Shiprocket "Master Courier".
MASTER_COURIER_PREFIXES = [
    ("amazon", "Amazon"),
    ("delhivery", "Delhivery"),
    ("bluedart", "Blue Dart"),
    ("blue dart", "Blue Dart"),
    ("xb ", "Xpressbees"),
    ("xpressbees", "Xpressbees"),
    ("shadowfax", "Shadowfax"),
    ("ekart", "Ekart"),
    ("dtdc", "DTDC"),
    ("ecom", "Ecom Express"),
    ("india post", "India Post"),
]

# Shiprocket writes "N/A" (not blank) in these columns when there is no value.
NA_COLUMNS = {
    "Order Picked Up Date", "Pickedup Timestamp", "Order Delivered Date",
    "RTO Initiated Date", "RTO Delivered Date", "First Out For Delivery Date",
    "Eway Bill Nos", "RTO Score Charges", "RTO Score Tax",
    "Delivery Boost Charges", "Delivery Boost Tax", "WhatsApp Tracking Charges",
    "WhatsApp Tracking Tax", "Brand Boost Charges", "Brand Boost Tax",
    "NDR 1 Attempt Date", "NDR 1 Remark", "NDR 2 Attempt Date", "NDR 2 Remark",
    "NDR 3 Attempt Date", "NDR 3 Remark", "Latest NDR Date", "Latest NDR Reason",
    "Lost Date", "Latest OFD Date",
}

MAX_PRODUCTS = 10

# Excel stores dates as "days since 1899-12-30"; 20000..80000 covers 1954..2119.
EXCEL_EPOCH = datetime(1899, 12, 30)
EXCEL_SERIAL_RE = re.compile(r"^([2-7]\d{4})(\.\d+)?$")


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
def nimbus_date(value: str, with_time: bool = True) -> str:
    """dd/mm/yyyy -> 'YYYY-MM-DD 00:00:00' (or 'YYYY-MM-DD').  '' stays ''."""
    value = (value or "").strip()
    if not value:
        return ""
    if EXCEL_SERIAL_RE.match(value):  # .xlsx stores dates as day counts
        dt = EXCEL_EPOCH + timedelta(days=float(value))
        return dt.strftime("%Y-%m-%d %H:%M:%S" if with_time else "%Y-%m-%d")
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S",
                "%Y-%m-%d %H:%M:%S", "%d/%m/%y"):
        try:
            dt = datetime.strptime(value, fmt)
            break
        except ValueError:
            continue
    else:
        return value  # unknown format: pass through untouched
    return dt.strftime("%Y-%m-%d %H:%M:%S" if with_time else "%Y-%m-%d")


def to_float(value: str, default: float = 0.0) -> float:
    value = (value or "").strip().replace(",", "")
    if not value:
        return default
    try:
        return float(value)
    except ValueError:
        return default


def money(value: float) -> str:
    return f"{value:.2f}"


def nimbus_status(value: str) -> str:
    key = (value or "").strip().lower()
    if not key:
        return ""
    return STATUS_MAP.get(key, key.replace("_", " ").upper())


def nimbus_zone(value: str) -> str:
    z = (value or "").strip().lower()
    if not z:
        return ""
    return z if z.startswith("z_") else f"z_{z}"


def master_courier(courier_name: str) -> str:
    name = (courier_name or "").strip().lower()
    if not name:
        return ""
    for prefix, master in MASTER_COURIER_PREFIXES:
        if name.startswith(prefix):
            return master
    return courier_name.split()[0]


def product_col(row: dict, base: str, idx: int) -> str:
    """Nimbus product column names carry a '*' on the required first-product
    columns, e.g. 'Product Name (1)*'.  Look up both spellings."""
    for name in (f"{base} ({idx})", f"{base} ({idx})*"):
        if name in row:
            return (row[name] or "").strip()
    return ""


def nimbus_products(row: dict) -> list[dict]:
    products = []
    for i in range(1, MAX_PRODUCTS + 1):
        sku = product_col(row, "Product SKU", i)
        name = product_col(row, "Product Name", i)
        if not sku and not name:
            continue
        products.append({
            "sku": sku,
            "name": name,
            "qty": product_col(row, "Product Quantity", i),
            "price": product_col(row, "Product Unit Price", i),
            "discount": product_col(row, "Product Discount", i),
            "hsn": product_col(row, "Product HSN Code", i),
            "tax_pct": product_col(row, "Product Tax %", i),
        })
    return products


def order_total(row: dict, products: list[dict]) -> float:
    subtotal = sum(to_float(p["qty"], 1) * to_float(p["price"]) for p in products)
    return (subtotal
            - to_float(row.get("Total Discount"))
            + to_float(row.get("Shipping Charges"))
            + to_float(row.get("COD Charges"))
            + to_float(row.get("Other Charges")))


# --------------------------------------------------------------------------- #
# Core conversion
# --------------------------------------------------------------------------- #
def convert_nimbus_row(row: dict) -> list[dict]:
    """Turn one Nimbus order row into 1..N Shiprocket rows (one per product)."""
    g = lambda k: (row.get(k) or "").strip()  # noqa: E731

    products = nimbus_products(row)
    if not products:
        # An order with no product lines still deserves a row.
        products = [{"sku": "", "name": "", "qty": "", "price": "",
                     "discount": "", "hsn": "", "tax_pct": ""}]

    total = money(order_total(row, products))
    order_date = nimbus_date(g("Order Date"))
    payment = g("Payment Method (COD/Prepaid)").lower()
    is_reverse = "Yes" if payment == "reverse" else "No"
    picked = nimbus_date(g("Picked Date"))
    remitted_date = nimbus_date(g("Remitted Date"), with_time=False)
    collectable = to_float(g("Collectable Amount"))
    dims = "x".join(d for d in (g("Shipment Length (cm)"),
                                g("Shipment Breadth (cm)"),
                                g("Shipment Height (cm)")) if d)

    common = {
        "Order ID": g("Order ID"),
        "Shiprocket Created At": order_date,
        "Channel": g("Channel Name"),
        "Status": nimbus_status(g("Shipment Status")),
        "Channel Created At": order_date,
        "Customer Name": g("Buyer's Full Name"),
        "Customer Email": g("Buyer's Email"),
        "Customer Mobile": g("Buyer's Mobile No."),
        "Address Line 1": g("Shipping Complete Address"),
        "Address Line 2": g("Shipping Address Landmark"),
        "Address City": g("Shipping Address City"),
        "Address State": g("Shipping Address State"),
        "Address Pincode": g("Shipping Address Pincode"),
        "Payment Method": payment,
        "Order Total": total,
        "Discount Value": money(to_float(g("Total Discount"))),
        "Weight (KG)": g("Shipment Weight (Kgs)"),
        "Dimensions (CM)": dims,
        "Charged Weight": g("Charged Weight"),
        "Courier Company": g("Courier Name"),
        "AWB Code": g("AWB"),
        "AWB Assigned Date": nimbus_date(g("Courier Assigned Date")),
        "Pickup Location ID": g("Pickup Address ID"),
        "Pickup Address Name": g("Pickup Warehouse Nickname"),
        "Order Picked Up Date": picked,
        "Pickedup Timestamp": picked,
        "Order Shipped Date": nimbus_date(g("Shipped Date")),
        "EDD": nimbus_date(g("EDD")),
        "Order Delivered Date": nimbus_date(g("Delivered Date")),
        "RTO Address": g("RTO Warehouse Address"),
        "RTO Delivered Date": nimbus_date(g("RTO Delivered Date")),
        "COD Remittance Date": remitted_date,
        "COD Payble Amount": money(to_float(g("COD Charges"))),
        "Remitted Amount": money(collectable) if remitted_date else "0.00",
        "CRF ID": g("Remittance ID"),
        "Zone": nimbus_zone(g("Zone")),
        "COD Charges": money(to_float(g("COD Charges"))),
        "Freight Total Amount": money(to_float(g("Total Freight Charges (inclusive of COD)"))),
        "Customer_invoice_id": g("Invoice ID"),
        "Shipping Charges": money(to_float(g("Shipping Charges"))),
        "Order Tags": g("Order Tags"),
        "Pickup Code": g("Pickup Warehouse Nickname"),
        "Pickup Pincode": g("Pincode"),
        "Master Courier": master_courier(g("Courier Name")),
        "Is Reverse": is_reverse,
    }

    out_rows = []
    for p in products:
        qty = to_float(p["qty"], 1)
        price = to_float(p["price"])
        tax_pct = to_float(p["tax_pct"])
        # Shiprocket reports tax as the GST component of a tax-inclusive price.
        tax = price * qty * tax_pct / (100 + tax_pct) if tax_pct else 0.0

        out = {col: "" for col in SHIPROCKET_COLUMNS}
        out.update(common)
        out.update({
            "Channel SKU": p["sku"],
            "Master SKU": p["sku"],
            "Product Name": p["name"],
            "Product Quantity": p["qty"],
            "Product Price": money(price) if p["price"] else "",
            "Tax": money(tax),
            "Tax %": money(tax_pct),
            "Product HSN": p["hsn"],
        })
        for col in NA_COLUMNS:
            if out[col] == "":
                out[col] = "N/A"
        out_rows.append(out)
    return out_rows


def is_blank_row(row: dict) -> bool:
    return not any((v or "").strip() for v in row.values())


def read_csv(path: str) -> tuple[list[str], list[dict]]:
    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), list(reader)


_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _xlsx_number(text: str) -> str:
    """Render a numeric cell the way the CSV export would (no float noise,
    no scientific notation, integers without '.0')."""
    try:
        f = float(text)
    except ValueError:
        return text
    if f.is_integer():
        return f"{f:.0f}"
    return repr(f)


def _fix_mojibake(text: str) -> str:
    """Excel sometimes opens a UTF-8 CSV as Windows-1252, turning '™' into
    'â„¢' and Hindi into 'à¤šà¤‚...'.  Undo that when the round-trip works."""
    if not text or not any(ch in text for ch in "Ã¢à"):
        return text
    try:
        return text.encode("cp1252").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


def _col_index(ref: str) -> int:
    """'A1' -> 0, 'DF2' -> 109."""
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def read_xlsx(path: str) -> tuple[list[str], list[dict]]:
    """Read the first worksheet of an .xlsx into (header, rows-of-strings)
    using only the standard library, so no openpyxl/pandas is required."""
    with zipfile.ZipFile(path) as z:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            for _, el in ET.iterparse(z.open("xl/sharedStrings.xml")):
                if el.tag == f"{_NS}si":
                    shared.append("".join(t.text or "" for t in el.iter(f"{_NS}t")))
                    el.clear()

        # First sheet listed in workbook.xml, resolved through the rels file.
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        first = wb.find(f"{_NS}sheets/{_NS}sheet")
        rid = first.get(f"{_REL_NS}id") if first is not None else None
        target = "worksheets/sheet1.xml"
        if rid:
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            for rel in rels:
                if rel.get("Id") == rid:
                    target = rel.get("Target").lstrip("/")
                    if target.startswith("xl/"):
                        target = target[3:]
        sheet_path = f"xl/{target}"

        header: list[str] = []
        rows: list[dict] = []
        for _, el in ET.iterparse(z.open(sheet_path)):
            if el.tag != f"{_NS}row":
                continue
            cells: dict[int, str] = {}
            for c in el.findall(f"{_NS}c"):
                idx = _col_index(c.get("r", ""))
                ctype = c.get("t", "n")
                v = c.find(f"{_NS}v")
                if ctype == "s":
                    val = _fix_mojibake(shared[int(v.text)]) if v is not None else ""
                elif ctype == "inlineStr":
                    val = _fix_mojibake("".join(t.text or "" for t in c.iter(f"{_NS}t")))
                elif ctype == "b":
                    val = "Yes" if (v is not None and v.text == "1") else "No"
                elif ctype in ("str", "e"):
                    val = v.text or "" if v is not None else ""
                else:
                    val = _xlsx_number(v.text or "") if v is not None else ""
                cells[idx] = (val or "").strip()
            el.clear()
            if not cells:
                continue
            width = max(cells) + 1
            values = [cells.get(i, "") for i in range(width)]
            if not header:
                header = values
                continue
            values = values[:len(header)] + [""] * (len(header) - len(values))
            rows.append(dict(zip(header, values)))
        return header, rows


def read_table(path: str) -> tuple[list[str], list[dict]]:
    """Read a .csv or .xlsx report into (header, rows-of-strings)."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        return read_xlsx(path)
    return read_csv(path)


def repair_nimbus_header(header: list[str], rows: list[dict]) -> tuple[list[str], list[dict]]:
    """If the file has Nimbus' column count but some header cells were edited
    (e.g. a stray 'x' typed into a header in Excel), restore the standard
    names by position so the data underneath is not lost."""
    if len(header) != len(NIMBUS_COLUMNS):
        return header, rows
    known = set(NIMBUS_COLUMNS)
    fixes = {h: c for h, c in zip(header, NIMBUS_COLUMNS) if h != c and h not in known}
    if not fixes:
        return header, rows
    for bad, good in fixes.items():
        print(f"warning: header {bad!r} is not a Nimbus column; treating it as {good!r}")
    new_header = [fixes.get(h, h) for h in header]
    return new_header, [{fixes.get(k, k): v for k, v in r.items()} for r in rows]


def convert_nimbus_file(path: str) -> tuple[list[dict], int]:
    """Returns (shiprocket-style rows, number of Nimbus orders converted)."""
    header, rows = read_table(path)
    header, rows = repair_nimbus_header(header, rows)
    if "Order ID" not in header or "Shipment Status" not in header:
        sys.exit(f"error: {path} does not look like a Nimbus Post order report "
                 f"(missing 'Order ID' / 'Shipment Status' columns)")
    out, orders = [], 0
    for row in rows:
        if is_blank_row(row) or not (row.get("Order ID") or "").strip():
            continue  # Nimbus exports end with a block of empty rows
        out.extend(convert_nimbus_row(row))
        orders += 1
    return out, orders


def write_csv(path: str, columns: list[str], rows: list[dict],
              source_column: bool = False) -> None:
    cols = columns + (["Source"] if source_column else [])
    tmp_fd, tmp_path = tempfile.mkstemp(dir=os.path.dirname(os.path.abspath(path)) or ".",
                                        suffix=".csv.tmp")
    with os.fdopen(tmp_fd, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore",
                                lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    os.replace(tmp_path, path)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Convert a Nimbus Post order report to Shiprocket's report "
                    "layout and append it below an existing Shiprocket report.")
    ap.add_argument("--nimbus", "-n", required=True,
                    help="Nimbus Post order_b2c_report (.csv or .xlsx)")
    ap.add_argument("--shiprocket", "-s",
                    help="Shiprocket report (.csv or .xlsx) to append below (optional; "
                         "omit to only convert the Nimbus file)")
    ap.add_argument("--output", "-o",
                    help="Output CSV path (default: <shiprocket>_combined.csv, or "
                         "<nimbus>_as_shiprocket.csv when no Shiprocket file is given)")
    ap.add_argument("--in-place", action="store_true",
                    help="Write the combined data back into the Shiprocket file itself")
    ap.add_argument("--skip-existing", action="store_true",
                    help="Skip Nimbus orders whose Order ID already exists in the "
                         "Shiprocket file")
    ap.add_argument("--source-column", action="store_true",
                    help="Add a trailing 'Source' column (Shiprocket / NimbusPost) "
                         "so the two carriers can be told apart")
    args = ap.parse_args(argv)

    if args.in_place and not args.shiprocket:
        ap.error("--in-place requires --shiprocket")
    if args.in_place and args.output:
        ap.error("use either --in-place or --output, not both")

    nimbus_rows, n_orders = convert_nimbus_file(args.nimbus)
    for r in nimbus_rows:
        r["Source"] = "NimbusPost"

    ship_rows: list[dict] = []
    columns = SHIPROCKET_COLUMNS
    if args.shiprocket:
        ship_header, ship_rows = read_table(args.shiprocket)
        if "Order ID" not in ship_header or "AWB Code" not in ship_header:
            sys.exit(f"error: {args.shiprocket} does not look like a Shiprocket report")
        # Follow the real file's column order if Shiprocket ever changes it.
        columns = list(ship_header)
        for col in SHIPROCKET_COLUMNS:
            if col not in columns:
                columns.append(col)
        for r in ship_rows:
            r["Source"] = "Shiprocket"
        if args.skip_existing:
            existing = {(r.get("Order ID") or "").strip() for r in ship_rows}
            before = len(nimbus_rows)
            nimbus_rows = [r for r in nimbus_rows if r["Order ID"] not in existing]
            print(f"skipped {before - len(nimbus_rows)} Nimbus row(s) already in Shiprocket file")

    if args.in_place:
        output = args.shiprocket
    elif args.output:
        output = args.output
    elif args.shiprocket:
        stem, _ = os.path.splitext(args.shiprocket)
        output = f"{stem}_combined.csv"
    else:
        stem, _ = os.path.splitext(args.nimbus)
        output = f"{stem}_as_shiprocket.csv"

    write_csv(output, columns, ship_rows + nimbus_rows, args.source_column)

    print(f"Nimbus orders converted : {n_orders} -> {len(nimbus_rows)} product rows")
    if args.shiprocket:
        print(f"Shiprocket rows kept    : {len(ship_rows)}")
    print(f"Total rows written      : {len(ship_rows) + len(nimbus_rows)}")
    print(f"Output                  : {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
