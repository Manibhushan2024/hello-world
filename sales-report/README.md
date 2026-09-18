# Sales Reports

Two Excel report generators over the same 182-SKU master list:

- **`build_report.py`** — marketplace (Amazon + Flipkart + Myntra), from the combined EasyEcom export.
- **`build_site_report.py`** — website (Shopify / OMS), from the site orders export.

## Marketplace report

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


---

# Website report (`build_site_report.py`)

```bash
python3 build_site_report.py <site_orders.csv> [output.xlsx]
python3 build_site_report.py "" blank.xlsx      # empty template
```

Same 15-day window logic and the same 182-SKU master as the marketplace report.

## Reading the site export

The site export is **one row per order** with up to 16 product blocks laid out
across the row, so quantities are read sideways rather than down a column:

| Block | SKU column | Quantity column |
|---|---|---|
| 1 | `AL` (38) | `AN` (40) |
| n | `38 + 7(n-1)` | `40 + 7(n-1)` |
| 16 | `EM` (143) | `EO` (145) |

`Order Date` is column `A`, `Order ID` is `B`, `Shipment Status` is `EY` (155).
All 16 blocks are summed, so multi-item orders count in full.

### Dates

The site export writes `Order Date` as `dd/mm/yyyy` **text**. The workbook reads
both that text and real Excel dates. On a US (`mm/dd/yyyy`) locale Excel converts
ambiguous values wrongly on paste, so `Settings` surfaces the earliest and latest
dates read plus an "orders whose date could not be read" counter as a tripwire.

## Pending orders

An order is pending when `Shipment Status` is `created` or `draft`. The report
gives the pending order count, the SKU-wise quantity behind it, and an
order-by-order list. Pending figures cover **every** pending order in the pasted
data, not just the 15-day window — an older unshipped order still needs picking.

## SKUs outside the master list

Site rows sometimes carry SKUs absent from the 182-SKU master (typos, retired
codes, a `'-` placeholder). Two mechanisms handle them.

**Merged duplicate listings.** `site_sku_aliases.json` maps a master SKU to the
alternate codes counted into it:

```json
{"jellybra-q1-cleoblue-ss": ["jellybra-q1-cleoblue-s"]}
```

The site carries both codes for the same size, so they are summed into the one
master row. Merged rows are highlighted on `SKU Master` with an "Also counts"
column, and the alternate code is dropped from the unlisted block so nothing is
counted twice.

**Everything else** is *never* folded into the 182 rows. It gets its own block
below the TOTAL row on `Website Sales`, and `Daily Orders` carries a per-day
"not in SKU list" units column so nothing is lost silently. The block is seeded
from `site_unlisted_skus.json`; a new bad code will not get its own row but will
still show in that column.

## Sheets

| Sheet | Contents |
|---|---|
| `Settings` | Window controls, live totals, date-read tripwire |
| `Daily Orders` | Orders and units per day, averages, pending count, unlisted check |
| `Website Sales` | Date-wise × SKU-wise units, 15-day total, average per day, unlisted block |
| `Best Sellers` | All 182 SKUs ranked for the window |
| `Pending Orders` | Pending count and SKU-wise quantity |
| `Pending Order List` | One row per pending order with its SKUs and quantities |
| `SKU Master`, `Raw Data`, `Work` | SKU list, paste area, hidden helper |

## Capacity

`Raw Data` / `Work` cover 25,000 orders (~23 days at current volume) and the
pending list holds 3,000 orders. Extend by widening the named ranges.

## Files

- `build_site_report.py` — website workbook generator
- `site_raw_header.txt` — site export header row, for building an empty template
- `site_unlisted_skus.json` — SKUs seen in the export that are not in `skus.txt`
- `site_sku_aliases.json` — alternate site codes folded into a master SKU
