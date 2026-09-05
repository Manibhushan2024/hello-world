"""Small self-contained tests for the Nimbus -> Shiprocket converter.

Run with:  python3 nimbus_to_shiprocket/test_convert.py -v
"""
import csv
import io
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(__file__))
import convert  # noqa: E402


def nimbus_row(**overrides):
    base = {
        "Order Date": "26/08/2026", "Order ID": "BC-1", "Channel Name": "Shopify - X",
        "Pickup Address ID": "WH-001", "RTO Warehouse Address": "Gali 15, Delhi, 110042",
        "Payment Method (COD/Prepaid)": "COD", "Buyer's Full Name": "A B",
        "Buyer's Email": "a@b.c", "Buyer's Mobile No.": "9999999999",
        "Shipping Complete Address": "House 1", "Shipping Address Landmark": "Near park",
        "Shipping Address State": "DELHI", "Shipping Address City": "DELHI",
        "Shipping Address Pincode": "110001", "Order Tags": "fastrr",
        "Shipment Weight (Kgs)": "0.3", "Shipment Length (cm)": "44",
        "Shipment Breadth (cm)": "13", "Shipment Height (cm)": "4",
        "Collectable Amount": "2798.2", "Shipping Charges": "100", "COD Charges": "0",
        "Total Discount": "299.8", "Other Charges": "0",
        "Product SKU (1)": "sku-a", "Product Name (1)*": "Item A",
        "Product Quantity (1)*": "1", "Product Unit Price (1)*": "1499",
        "Product HSN Code (1)": "62121000", "Product Tax % (1)": "0",
        "Product SKU (2)": "sku-b", "Product Name (2)": "Item B",
        "Product Quantity (2)": "1", "Product Unit Price (2)": "1499",
        "Product HSN Code (2)": "62121000", "Product Tax % (2)": "5",
        "Courier Name": "Amazon 500 g", "Courier Assigned Date": "29/08/2026",
        "AWB": "372126000000", "Zone": "b", "Shipment Status": "delivered",
        "Picked Date": "31/08/2026", "Shipped Date": "01/09/2026", "EDD": "02/09/2026",
        "Delivered Date": "01/09/2026", "RTO Delivered Date": "",
        "Total Freight Charges (inclusive of COD)": "61.92", "Charged Weight": "0.458",
        "Remittance ID": "", "Remitted Date": "",
        "Pickup Warehouse Nickname": "Samaypur", "Pincode": "110042",
    }
    base.update(overrides)
    return base


class ConvertRowTests(unittest.TestCase):
    def test_one_row_per_product(self):
        rows = convert.convert_nimbus_row(nimbus_row())
        self.assertEqual([r["Channel SKU"] for r in rows], ["sku-a", "sku-b"])
        self.assertEqual([r["Master SKU"] for r in rows], ["sku-a", "sku-b"])
        for r in rows:
            self.assertEqual(set(r), set(convert.SHIPROCKET_COLUMNS))

    def test_order_total_matches_collectable_amount(self):
        rows = convert.convert_nimbus_row(nimbus_row())
        self.assertEqual(rows[0]["Order Total"], "2798.20")
        self.assertEqual(rows[1]["Order Total"], "2798.20")
        self.assertEqual(rows[0]["Discount Value"], "299.80")

    def test_field_conversions(self):
        r = convert.convert_nimbus_row(nimbus_row())[0]
        self.assertEqual(r["Shiprocket Created At"], "2026-08-26 00:00:00")
        self.assertEqual(r["Order Delivered Date"], "2026-09-01 00:00:00")
        self.assertEqual(r["RTO Delivered Date"], "N/A")
        self.assertEqual(r["Status"], "DELIVERED")
        self.assertEqual(r["Payment Method"], "cod")
        self.assertEqual(r["Zone"], "z_b")
        self.assertEqual(r["Dimensions (CM)"], "44x13x4")
        self.assertEqual(r["Master Courier"], "Amazon")
        self.assertEqual(r["Courier Company"], "Amazon 500 g")
        self.assertEqual(r["Address Line 2"], "Near park")
        self.assertEqual(r["Is Reverse"], "No")
        self.assertEqual(r["Product Price"], "1499.00")
        self.assertEqual(r["Tax %"], "0.00")

    def test_inclusive_tax(self):
        r = convert.convert_nimbus_row(nimbus_row())[1]
        self.assertEqual(r["Tax %"], "5.00")
        self.assertEqual(r["Tax"], "71.38")  # 1499 * 5 / 105

    def test_status_mapping_and_fallback(self):
        self.assertEqual(convert.nimbus_status("pickup_failed"), "PICKUP EXCEPTION")
        self.assertEqual(convert.nimbus_status("rto_in_transit"), "RTO IN TRANSIT")
        self.assertEqual(convert.nimbus_status("something_new"), "SOMETHING NEW")
        self.assertEqual(convert.nimbus_status(""), "")

    def test_reverse_order(self):
        r = convert.convert_nimbus_row(nimbus_row(**{"Payment Method (COD/Prepaid)": "Reverse"}))[0]
        self.assertEqual(r["Payment Method"], "reverse")
        self.assertEqual(r["Is Reverse"], "Yes")

    def test_remittance(self):
        r = convert.convert_nimbus_row(nimbus_row(**{"Remitted Date": "03/09/2026",
                                                      "Remittance ID": "R1"}))[0]
        self.assertEqual(r["COD Remittance Date"], "2026-09-03")
        self.assertEqual(r["Remitted Amount"], "2798.20")
        self.assertEqual(r["CRF ID"], "R1")


class XlsxTests(unittest.TestCase):
    def _write_xlsx(self, path, header, rows):
        """Minimal .xlsx writer (inline strings + numbers) for tests."""
        import zipfile
        from xml.sax.saxutils import escape

        def col(i):
            s = ""
            i += 1
            while i:
                i, r = divmod(i - 1, 26)
                s = chr(65 + r) + s
            return s

        def cell(ref, v):
            if isinstance(v, (int, float)):
                return f'<c r="{ref}"><v>{v}</v></c>'
            return f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(v))}</t></is></c>'

        xml_rows = []
        for r_idx, row in enumerate([header] + rows, start=1):
            cells = "".join(cell(f"{col(c)}{r_idx}", v) for c, v in enumerate(row) if v != "")
            xml_rows.append(f'<row r="{r_idx}">{cells}</row>')
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("xl/workbook.xml",
                       f'<workbook xmlns="{ns}" xmlns:r="{rns}"><sheets>'
                       f'<sheet name="Orders" sheetId="1" r:id="rId1"/></sheets></workbook>')
            z.writestr("xl/_rels/workbook.xml.rels",
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>')
            z.writestr("xl/worksheets/sheet1.xml",
                       f'<worksheet xmlns="{ns}"><sheetData>{"".join(xml_rows)}</sheetData></worksheet>')

    def test_read_xlsx_numbers_dates_and_mojibake(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "n.xlsx")
            self._write_xlsx(path, ["Order Date", "AWB", "Charged Weight", "Name", "Pincode"],
                             [[46260, 372125664635, 0.22900000000000001, "InstaTuck\u00e2\u201e\u00a2 Beige", 110042]])
            header, rows = convert.read_xlsx(path)
            self.assertEqual(header, ["Order Date", "AWB", "Charged Weight", "Name", "Pincode"])
            self.assertEqual(rows[0]["AWB"], "372125664635")          # no 3.72E+11
            self.assertEqual(rows[0]["Charged Weight"], "0.229")       # no float noise
            self.assertEqual(rows[0]["Name"], "InstaTuck\u2122 Beige")  # mojibake undone
            self.assertEqual(convert.nimbus_date(rows[0]["Order Date"]), "2026-08-26 00:00:00")
            self.assertEqual(convert.nimbus_date(rows[0]["Pincode"]), "110042")  # not a date

    def test_repair_edited_header(self):
        header = list(convert.NIMBUS_COLUMNS)
        header[2] = "x"  # someone typed over 'Channel Name' in Excel
        rows = [{h: "" for h in header}]
        rows[0]["x"] = "Shopify"
        fixed_header, fixed_rows = convert.repair_nimbus_header(header, rows)
        self.assertEqual(fixed_header, convert.NIMBUS_COLUMNS)
        self.assertEqual(fixed_rows[0]["Channel Name"], "Shopify")

    def test_repair_leaves_foreign_layout_alone(self):
        header, rows = convert.repair_nimbus_header(["a", "b"], [{"a": "1", "b": "2"}])
        self.assertEqual(header, ["a", "b"])


class EndToEndTests(unittest.TestCase):
    def _write(self, path, columns, rows):
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=columns)
            w.writeheader()
            w.writerows(rows)

    def test_append_below_shiprocket(self):
        with tempfile.TemporaryDirectory() as d:
            nimbus = os.path.join(d, "nimbus.csv")
            ship = os.path.join(d, "ship.csv")
            out = os.path.join(d, "out.csv")
            n1 = nimbus_row()
            blank = {k: "" for k in n1}
            self._write(nimbus, list(n1), [n1, blank, blank])
            s1 = {c: "" for c in convert.SHIPROCKET_COLUMNS}
            s1.update({"Order ID": "ABBC-1", "AWB Code": "123", "Product Name": "Existing"})
            self._write(ship, convert.SHIPROCKET_COLUMNS, [s1])

            convert.main(["--nimbus", nimbus, "--shiprocket", ship, "-o", out])

            with open(out, newline="", encoding="utf-8") as fh:
                reader = csv.DictReader(fh)
                self.assertEqual(reader.fieldnames, convert.SHIPROCKET_COLUMNS)
                rows = list(reader)
            self.assertEqual([r["Order ID"] for r in rows], ["ABBC-1", "BC-1", "BC-1"])
            self.assertEqual(rows[0]["Product Name"], "Existing")  # untouched

    def test_skip_existing_and_source_column(self):
        with tempfile.TemporaryDirectory() as d:
            nimbus = os.path.join(d, "nimbus.csv")
            ship = os.path.join(d, "ship.csv")
            out = os.path.join(d, "out.csv")
            n1, n2 = nimbus_row(), nimbus_row(**{"Order ID": "BC-2"})
            self._write(nimbus, list(n1), [n1, n2])
            s1 = {c: "" for c in convert.SHIPROCKET_COLUMNS}
            s1.update({"Order ID": "BC-1", "AWB Code": "123"})
            self._write(ship, convert.SHIPROCKET_COLUMNS, [s1])

            convert.main(["-n", nimbus, "-s", ship, "-o", out,
                          "--skip-existing", "--source-column"])

            with open(out, newline="", encoding="utf-8") as fh:
                rows = list(csv.DictReader(fh))
            self.assertEqual([(r["Order ID"], r["Source"]) for r in rows],
                             [("BC-1", "Shiprocket"), ("BC-2", "NimbusPost"), ("BC-2", "NimbusPost")])

    def test_rejects_wrong_file(self):
        with tempfile.TemporaryDirectory() as d:
            bad = os.path.join(d, "bad.csv")
            self._write(bad, ["a", "b"], [{"a": "1", "b": "2"}])
            with self.assertRaises(SystemExit):
                convert.main(["-n", bad])


if __name__ == "__main__":
    unittest.main()
