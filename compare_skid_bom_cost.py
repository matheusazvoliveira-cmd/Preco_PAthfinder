"""
Compare 'SKID BOM.xlsx' against '7-PAC Item Costs.xlsx', the same check that
compare_bom_cost.py performs for the Pathfinder BOM, adapted to the SKID BOM's
layout.

Matching key:
  - SKID BOM: Part Number column (position varies per sheet, see SHEET_LAYOUTS)
  - PAC:      column B -> 'product_number' (sheet 'Item Costs')

Unlike the Pathfinder BOM, the SKID BOM has no existing cost column, so every
line is priced from the PAC (Brazil, BRL) reference only:
  1. Notes/Name/Category say the item is included elsewhere or supplied with
     no cost (e.g. "Included in ...", "No cost", "via customer") -> 0,
     source=INCLUDED / NO_COST.
  2. Otherwise, use the PAC reference cost (Brazil, BRL) if a match was
     found -> source=PAC (BRL), converted to USD using BRL_PER_USD.
  3. No match anywhere -> 0, source=MISSING (needs manual follow-up).

The SKID BOM workbook has 5 sheets with inconsistent layouts (some have a
leading ID column, some don't, some lack a category column). SHEET_LAYOUTS
below maps each sheet to its Part Number / Name / Qty / Category columns.

Output:
  - 'SKID BOM - Priced.xlsx' with one sheet per source sheet (original data
    plus PAC match / resolved cost columns) and a combined Totals sheet.
  - A console summary with PAC match rate and totals per sheet and overall.
"""
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter

from compare_bom_cost import (
    BRL_PER_USD,
    INCLUDED_MARKERS,
    NO_COST_MARKERS,
    normalize,
    load_pac_costs,
    find_pac_reference,
)

BASE = Path(__file__).parent
SKID_FILE = BASE / "SKID BOM.xlsx"
OUT_FILE = BASE / "SKID BOM - Priced.xlsx"

# column index (0-based, as returned by openpyxl values_only rows) of each
# field per sheet - the workbook's sheets don't share a common layout.
SHEET_LAYOUTS = {
    "Skid":        {"start_row": 1, "id_col": 0,    "pn_col": 1, "name_col": 2, "qty_col": 3, "cat_col": 4},
    "Air skid":    {"start_row": 1, "id_col": None, "pn_col": 0, "name_col": 1, "qty_col": 2, "cat_col": 3},
    "battery kit": {"start_row": 1, "id_col": None, "pn_col": 0, "name_col": 1, "qty_col": 2, "cat_col": 3},
    "Frame kit":   {"start_row": 1, "id_col": None, "pn_col": 0, "name_col": 1, "qty_col": 2, "cat_col": None},
    "QUESTION":    {"start_row": 0, "id_col": 0,    "pn_col": 1, "name_col": 2, "qty_col": 3, "cat_col": 4},
}


def as_qty(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def resolve_cost(name, category, pac_ref):
    """Decide which cost feeds the 'Resolved Cost' column (no BOM cost exists
    in the SKID BOM, so PAC is the only real source)."""
    text_blob = f"{name or ''} {category or ''}".lower()

    if any(marker in text_blob for marker in INCLUDED_MARKERS):
        return {"cost": 0, "source": "INCLUDED", "currency": "-"}

    if any(marker in text_blob for marker in NO_COST_MARKERS):
        return {"cost": 0, "source": "NO_COST", "currency": "-"}

    if pac_ref is not None:
        return {"cost": pac_ref["cost"], "source": "PAC", "currency": pac_ref["currency"]}

    return {"cost": 0, "source": "MISSING", "currency": "-"}


def main():
    wb_skid = openpyxl.load_workbook(SKID_FILE, data_only=True)
    pac_lookup = load_pac_costs()

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    totals_by_sheet_source_currency = {}
    grand_pac_found = 0
    grand_rows = 0
    grand_missing = []
    price_diffs_usd = []

    for sheet_name, layout in SHEET_LAYOUTS.items():
        if sheet_name not in wb_skid.sheetnames:
            continue
        ws = wb_skid[sheet_name]
        rows = list(ws.iter_rows(values_only=True))[layout["start_row"]:]

        cols_order = []
        if layout["id_col"] is not None:
            cols_order.append(("ID", layout["id_col"]))
        cols_order.append(("Part Number", layout["pn_col"]))
        cols_order.append(("Name", layout["name_col"]))
        cols_order.append(("Qty", layout["qty_col"]))
        if layout["cat_col"] is not None:
            cols_order.append(("Category", layout["cat_col"]))

        extra_cols = [
            "PAC Ref Cost (BRL, Brazil, per unit)",
            "PAC Ref Currency",
            "PAC Ref Cost (converted to USD, per unit)",
            "Extended Cost (USD)",
            "Cost Source",
            "PAC Ref Matches (plants)",
            "PAC Ref Description",
        ]

        out_ws = out_wb.create_sheet(sheet_name)
        out_ws.append(["PAC Matched"] + [label for label, _ in cols_order] + extra_cols)

        sheet_rows = 0
        sheet_pac_found = 0
        sheet_missing = []

        for row in rows:
            if row is None or all(v is None for v in row):
                continue
            sheet_rows += 1
            grand_rows += 1

            name = row[layout["name_col"]] if layout["name_col"] < len(row) else None
            category = row[layout["cat_col"]] if layout["cat_col"] is not None and layout["cat_col"] < len(row) else None
            raw_pn = row[layout["pn_col"]] if layout["pn_col"] < len(row) else None
            part_number = normalize(raw_pn)

            pac_ref = find_pac_reference(part_number, pac_lookup)
            if pac_ref is not None:
                sheet_pac_found += 1
                grand_pac_found += 1
            else:
                sheet_missing.append((raw_pn, name))
                grand_missing.append((sheet_name, raw_pn, name))

            result = resolve_cost(name, category, pac_ref)
            unit_cost_brl = result["cost"] or 0
            source = result["source"]
            currency = result["currency"]

            unit_cost_usd = None
            ext_cost_usd = None
            qty = as_qty(row[layout["qty_col"]] if layout["qty_col"] < len(row) else None)
            if source == "PAC" and currency == "BRL":
                unit_cost_usd = unit_cost_brl / BRL_PER_USD
                if qty is not None:
                    ext_cost_usd = unit_cost_usd * qty
                    price_diffs_usd.append(ext_cost_usd)

            key = (sheet_name, source, currency)
            totals_by_sheet_source_currency[key] = totals_by_sheet_source_currency.get(key, 0.0) + (
                ext_cost_usd if ext_cost_usd is not None else 0.0
            )

            out_ws.append(
                ["Yes" if pac_ref else "No"]
                + [row[idx] if idx < len(row) else None for _, idx in cols_order]
                + [
                    pac_ref["cost"] if pac_ref else None,
                    pac_ref["currency"] if pac_ref else None,
                    unit_cost_usd,
                    ext_cost_usd,
                    source,
                    pac_ref["matches"] if pac_ref else 0,
                    pac_ref["desc"] if pac_ref else None,
                ]
            )

        for col_idx in range(1, len(cols_order) + len(extra_cols) + 2):
            out_ws.column_dimensions[get_column_letter(col_idx)].width = 20

        print(f"[{sheet_name}] rows: {sheet_rows}  PAC matched: {sheet_pac_found}  unmatched: {len(sheet_missing)}")
        for pn, name in sheet_missing:
            print(f"    - {pn!r} | {name}")

    # Combined totals sheet
    tot_ws = out_wb.create_sheet("Totals")
    tot_ws.append(["Sheet", "Source", "Currency", "Extended Total (USD, PAC lines only)"])
    grand_total_usd = 0.0
    for (sheet_name, source, currency), total in sorted(totals_by_sheet_source_currency.items()):
        tot_ws.append([sheet_name, source, currency, total])
        if source == "PAC":
            grand_total_usd += total
    tot_ws.append([])
    tot_ws.append(["Grand total (USD, all sheets, PAC-priced lines only)", "", "", grand_total_usd])
    tot_ws.append(["BRL_PER_USD rate used", "", "", BRL_PER_USD])

    out_wb.save(OUT_FILE)

    print(f"\nSKID BOM rows processed (all sheets): {grand_rows}")
    print(f"PAC (Brazil) reference found for: {grand_pac_found} / {grand_rows} part numbers")
    print(f"Unmatched in PAC (no reference at all): {len(grand_missing)}")
    print(f"\nGrand total (USD, PAC-priced lines only): {grand_total_usd:,.2f}")
    print(f"Saved priced workbook to: {OUT_FILE}")


if __name__ == "__main__":
    main()
