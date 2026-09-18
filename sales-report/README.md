# Marketplace Sales Report

Builds an Excel workbook that turns the combined EasyEcom marketplace export
(Amazon + Flipkart + Myntra) into a date-wise / SKU-wise sales report.

Every figure in the workbook is a live formula reading the `Raw Data` sheet, so a
fresh export can be pasted in and the whole report refreshes itself — no rebuild
needed.

## Usage

```bash
pip install openpyxl
python3 build_report.py <export.csv> [output.xlsx]   # pre-loaded with that export
python3 build_report.py "" blank.xlsx                # empty template, ready to paste into
```

## Reporting window

The workbook finds the latest order date in the pasted data, drops it (the export's
final day is always partial) and reports the 15 full days before it. An export
covering 1 Sep – 18 Sep therefore reports 3 Sep – 17 Sep. The length, the number of
trailing days dropped, and a hard-pinned end date are all editable on `Settings`.

## Counting rules

- Every order line counts regardless of `Order Status` — Shipped, Printed, Assigned,
  Upcoming and **Cancelled** all count as a sale. No status filter is applied.
- Units per line = `MAX(Item Quantity, Suborder Quantity)`, which guards against rows
  where one of the two fields is 0.
- Sales are booked on `Order Date`.
- Platform comes from `MP Name`: contains `Amazon` / `Flipkart` / `Myntra`. Anything
  else is tagged `Other` and surfaces in the check columns on `Daily Orders`.

## Sheets

| Sheet | Contents |
|---|---|
| `README` | In-workbook usage notes |
| `Settings` | Window controls and live totals |
| `Daily Orders` | Units and order lines per day per marketplace, plus daily averages |
| `Combined Daily` | One block per date; under it Amazon / Flipkart / Myntra and a day total. SKUs across the top |
| `Amazon`, `Flipkart`, `Myntra` | Date-wise × SKU-wise units, 15-day total, average per day |
| `Best Sellers - *` | All SKUs ranked per platform, plus an all-platforms ranking |
| `SKU Master` | The SKU list driving every sheet |
| `Raw Data` | Paste area (export header pre-filled) |
| `Work` | Hidden helper normalising date, platform, SKU and quantity |

Every SKU appears on every sheet; a SKU with no sales shows `0`, never a blank.

## Columns consumed from the export

`C` MP Name · `N` Order Date · `AE` Suborder Quantity · `AF` Item Quantity · `AG` SKU

The `Work` sheet reaches these through whole-column `INDEX(...,ROW())` references, so
inserting or deleting rows in `Raw Data` cannot break it.

## Files

- `build_report.py` — workbook generator
- `skus.txt` — the 182 tracked SKUs, one per line
- `raw_header.txt` — export header row, used when building an empty template

## Capacity

`Work` covers 30,000 export rows (~55 days at current volumes). To go beyond that,
copy the last `Work` row further down and extend the `WDATE` / `WPLAT` / `WSKU` /
`WQTY` named ranges.
