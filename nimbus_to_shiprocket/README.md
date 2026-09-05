# Nimbus Post → Shiprocket report converter

Turns a Nimbus Post `order_b2c_report_*.csv` into the exact column layout of a
Shiprocket `secure_*_reports_*.csv`, and appends the converted rows **below** the
Shiprocket data so both carriers live in one file.

Pure Python 3 (standard library only). Nothing to install.

## Quick start

```bash
# 1. Convert Nimbus and append it below the Shiprocket report -> <shiprocket>_combined.csv
python3 nimbus_to_shiprocket/convert.py --nimbus nimbus.csv --shiprocket shiprocket.csv

# 2. Same, but choose the output name
python3 nimbus_to_shiprocket/convert.py -n nimbus.csv -s shiprocket.csv -o combined.csv

# 3. Only convert Nimbus (no Shiprocket file) -> <nimbus>_as_shiprocket.csv
python3 nimbus_to_shiprocket/convert.py -n nimbus.csv

# 4. Overwrite the Shiprocket file itself with the combined data
python3 nimbus_to_shiprocket/convert.py -n nimbus.csv -s shiprocket.csv --in-place
```

Optional flags:

| Flag | What it does |
|---|---|
| `--skip-existing` | Drop Nimbus orders whose Order ID is already in the Shiprocket file |
| `--source-column` | Add a trailing `Source` column (`Shiprocket` / `NimbusPost`) |

The original Shiprocket rows are copied through untouched, in their original
order, followed by the converted Nimbus rows.

## What the conversion does

| Topic | Nimbus | Shiprocket | Handling |
|---|---|---|---|
| Rows | 1 row per **order**, products in `Product SKU (1)…(10)` | 1 row per **product line** | Each Nimbus order becomes one row per product; order-level fields repeat on every line, like Shiprocket does |
| Dates | `26/08/2026` | `2026-08-26 00:00:00` | Converted; time is `00:00:00` because Nimbus only exports the day |
| Status | `in_transit`, `pickup_failed`, … | `IN TRANSIT`, `PICKUP EXCEPTION`, … | Mapped via `STATUS_MAP` in `convert.py`; unknown values are upper-cased |
| Payment | `COD` / `Prepaid` / `Reverse` | `cod` / `prepaid` | Lower-cased; `Reverse` also sets `Is Reverse = Yes` |
| Zone | `b` | `z_b` | Prefixed |
| Order Total | not exported | `2798.20` | `Σ(qty × unit price) − Total Discount + Shipping + COD + Other charges`. This reproduces Nimbus' own *Collectable Amount* on every COD order in the sample file |
| Dimensions | 3 columns | `22x13x4` | Joined as `LxBxH` |
| Tax | `Product Tax %` | `Tax`, `Tax %` | Tax = GST share of the tax-inclusive price, the way Shiprocket reports it |
| Master Courier | `Amazon 250`, `XB DS 500g`, … | `Amazon`, `Xpressbees`, … | Derived from the courier name prefix |
| Empty cells | blank | `N/A` in some columns | The same columns Shiprocket fills with `N/A` are filled with `N/A` |

Full column-by-column mapping is in `convert_nimbus_row()` inside `convert.py`.
Shiprocket columns that have no Nimbus equivalent (Forward ID, NDR/NPR details,
risk scores, UTR, etc.) are left blank.

## Things to know

* **Trailing blank rows** in Nimbus exports are skipped automatically.
* **AWB numbers in scientific notation** (`3.72126E+11`): if the Nimbus file was
  opened and re-saved in Excel before running the converter, Excel has already
  destroyed those AWB values and they cannot be recovered. Run the converter on
  the CSV exactly as downloaded from Nimbus to keep full AWBs.
* **Discount Value** carries the order-level *Total Discount* from Nimbus,
  repeated on every product line of that order.
* The converter validates that each input really is a Nimbus / Shiprocket report
  and exits with a clear error otherwise.

## Tests

```bash
python3 nimbus_to_shiprocket/test_convert.py -v
```
