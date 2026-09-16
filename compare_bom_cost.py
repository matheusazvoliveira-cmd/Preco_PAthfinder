"""
Compare 'Pathfinder BOM.xlsx' against '7-PAC Item Costs.xlsx' to price the full
Pathfinder Box Car bill of material.

Matching key:
  - BOM:  column E  -> 'Part Number'
  - PAC:  column B  -> 'product_number'  (sheet 'Item Costs')

IMPORTANT: the BOM cost is the US cost and the PAC cost is the Brazil plant
cost (currency_code = BRL in the file). They are two different cost bases for
two different markets and are NEVER summed or validated against each other -
they are reported side by side for reference only.

The PAC lookup is now performed for EVERY BOM line (not only the ones missing
a BOM cost), so the true PAC match rate across all 281 parts is visible, even
when the BOM already carries its own (US) cost for that line.

Cost source priority per BOM line (drives the 'Resolved Cost'/US total only):
  1. BOM already has a non-zero 'Total Amount' -> use it as-is (source=BOM).
  2. Notes/Description say the item is included in another line or supplied
     with no cost (e.g. "Included in ...", "No cost", "via customer") -> 0,
     source=INCLUDED / NO_COST.
  3. Otherwise, use the PAC reference cost (Brazil, BRL) if a match was
     found -> source=PAC (BRL).
  4. No match anywhere -> 0, source=MISSING (needs manual follow-up).

Output:
  - 'Pathfinder BOM - Priced.xlsx' with the original BOM columns plus a
    PAC reference (found for every row where possible), the resolved
    cost/source used for the US total, and a totals sheet.
  - A console summary with grand totals split by source/currency, plus the
    overall PAC match rate across all 281 rows (regardless of BOM source).
  - For rows with both a BOM (US) cost and a PAC (BRL) match, the PAC cost
    is converted to USD (using BRL_PER_USD below) and compared to the BOM
    cost as a percentage difference.
"""
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent
BOM_FILE = BASE / "Pathfinder BOM.xlsx"
PAC_FILE = BASE / "7-PAC Item Costs.xlsx"
OUT_FILE = BASE / "Pathfinder BOM - Priced.xlsx"

# Update to the current market rate as needed (BRL per 1 USD).
BRL_PER_USD = 5.16

INCLUDED_MARKERS = ("included in",)
NO_COST_MARKERS = ("no cost", "via customer")


def normalize(pn):
    if pn is None:
        return None
    return str(pn).strip().upper()


def load_pac_costs():
    """Return dict: normalized product_number -> list of PAC rows (dicts)."""
    wb = openpyxl.load_workbook(PAC_FILE, read_only=True, data_only=True)
    ws = wb["Item Costs"]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {name: i for i, name in enumerate(header)}

    lookup = {}
    for r in rows:
        pn = normalize(r[idx["product_number"]])
        if not pn:
            continue
        resource_plus_overhead = r[idx["resource_plus_overhead"]] or 0
        material_plus_moh = r[idx["material_plus_moh"]] or 0
        aco_cost = r[idx["aco_cost"]]
        if aco_cost is None:
            aco_cost = resource_plus_overhead + material_plus_moh
        entry = {
            "inv_org_code": r[idx["inv_org_code"]],
            "product_desc": r[idx["product_desc"]],
            "currency_code": r[idx["currency_code"]],
            "aco_cost": aco_cost,
        }
        lookup.setdefault(pn, []).append(entry)
    wb.close()
    return lookup


def find_pac_reference(part_number, pac_lookup):
    """PAC (Brazil, BRL) reference lookup - run for every row regardless of
    whether the BOM already has its own (US) cost."""
    if not part_number or part_number not in pac_lookup:
        return None
    matches = pac_lookup[part_number]
    # Multiple plants (inv_org_code) can price the same part differently;
    # report the average unit cost and how many plant entries were averaged.
    avg_cost = sum(m["aco_cost"] for m in matches) / len(matches)
    currencies = {m["currency_code"] for m in matches}
    return {
        "cost": avg_cost,
        "currency": currencies.pop() if len(currencies) == 1 else "/".join(sorted(currencies)),
        "matches": len(matches),
        "desc": matches[0]["product_desc"],
    }


def resolve_cost(bom_row, pac_ref, header_idx):
    """Decide which cost feeds the 'Resolved Cost' (US total) column."""
    total_amount = bom_row[header_idx["Total Amount"]]
    notes = bom_row[header_idx["Notes"]] or ""
    description = bom_row[header_idx["Description"]] or ""
    text_blob = f"{notes} {description}".lower()

    if total_amount not in (None, 0):
        return {"cost": total_amount, "source": "BOM", "currency": "BOM-currency"}

    if any(marker in text_blob for marker in INCLUDED_MARKERS):
        return {"cost": 0, "source": "INCLUDED", "currency": "-"}

    if any(marker in text_blob for marker in NO_COST_MARKERS):
        return {"cost": 0, "source": "NO_COST", "currency": "-"}

    if pac_ref is not None:
        return {"cost": pac_ref["cost"], "source": "PAC", "currency": pac_ref["currency"]}

    return {"cost": 0, "source": "MISSING", "currency": "-"}


def main():
    wb_bom = openpyxl.load_workbook(BOM_FILE, data_only=True)
    ws_bom = wb_bom["Sheet1"]
    rows = list(ws_bom.iter_rows(values_only=True))
    header = list(rows[0])
    header_idx = {name: i for i, name in enumerate(header)}
    data_rows = rows[1:]

    pac_lookup = load_pac_costs()

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Priced BOM"
    extra_cols = [
        "Resolved Cost (US total)",
        "Cost Source",
        "PAC Ref Cost (BRL, Brazil)",
        "PAC Ref Currency",
        "PAC Ref Cost (converted to USD)",
        "Price Diff % (PAC USD vs BOM US)",
        "PAC Ref Matches (plants)",
        "PAC Ref Description",
    ]
    out_ws.append(["PAC Matched"] + header + extra_cols)

    totals_by_source_currency = {}
    grand_total_bom_currency = 0.0
    unmatched = []
    pac_ref_found = 0
    pac_ref_missing_parts = []
    price_diff_pcts = []

    for row in data_rows:
        part_number = normalize(row[header_idx["Part Number"]])
        pac_ref = find_pac_reference(part_number, pac_lookup)
        if pac_ref is not None:
            pac_ref_found += 1
        elif part_number:
            pac_ref_missing_parts.append((row[header_idx["Part Number"]], row[header_idx["Component"]]))

        result = resolve_cost(row, pac_ref, header_idx)
        cost = result["cost"] or 0
        source = result["source"]
        currency = result["currency"]

        key = (source, currency)
        totals_by_source_currency[key] = totals_by_source_currency.get(key, 0.0) + cost

        if source == "BOM":
            grand_total_bom_currency += cost

        if source == "MISSING":
            unmatched.append((row[header_idx["Part Number"]], row[header_idx["Component"]]))

        bom_us_cost = row[header_idx["Total Amount"]]
        pac_cost_usd = None
        price_diff_pct = None
        if pac_ref is not None and pac_ref["currency"] == "BRL":
            pac_cost_usd = pac_ref["cost"] / BRL_PER_USD
            if bom_us_cost not in (None, 0):
                price_diff_pct = (pac_cost_usd - bom_us_cost) / bom_us_cost * 100
                price_diff_pcts.append(price_diff_pct)

        out_ws.append(
            ["Yes" if pac_ref else "No"]
            + list(row)
            + [
                cost,
                source,
                pac_ref["cost"] if pac_ref else None,
                pac_ref["currency"] if pac_ref else None,
                pac_cost_usd,
                price_diff_pct,
                pac_ref["matches"] if pac_ref else 0,
                pac_ref["desc"] if pac_ref else None,
            ]
        )

    for col_idx in range(1, len(header) + len(extra_cols) + 2):
        out_ws.column_dimensions[get_column_letter(col_idx)].width = 18

    # Totals sheet
    tot_ws = out_wb.create_sheet("Totals")
    tot_ws.append(["Source", "Currency", "Total"])
    for (source, currency), total in sorted(totals_by_source_currency.items()):
        tot_ws.append([source, currency, total])
    tot_ws.append([])
    tot_ws.append(["Grand total priced in BOM's own currency (source=BOM only)", "", grand_total_bom_currency])
    tot_ws.append([])
    tot_ws.append(["BRL_PER_USD rate used", "", BRL_PER_USD])
    if price_diff_pcts:
        avg_price_diff_pct = sum(price_diff_pcts) / len(price_diff_pcts)
        tot_ws.append(["Average Price Diff % (PAC USD vs BOM US, matched rows only)", "", avg_price_diff_pct])

    out_wb.save(OUT_FILE)

    print(f"BOM rows processed: {len(data_rows)}")
    print(f"PAC (Brazil) reference found for: {pac_ref_found} / {len(data_rows)} part numbers")
    print(f"Unmatched in PAC (no reference at all): {len(pac_ref_missing_parts)}")
    print(f"Rows with NO cost anywhere (BOM nor PAC nor included/no-cost): {len(unmatched)}")
    for pn, comp in unmatched:
        print(f"  - {pn} | {comp}")
    print("\nResolved (US) totals by source/currency:")
    for (source, currency), total in sorted(totals_by_source_currency.items()):
        print(f"  {source:10s} {currency:12s} {total:,.2f}")
    if price_diff_pcts:
        avg_price_diff_pct = sum(price_diff_pcts) / len(price_diff_pcts)
        print(
            f"\nPrice Diff % (PAC in USD vs BOM in USD, rate={BRL_PER_USD} BRL/USD): "
            f"avg {avg_price_diff_pct:+.1f}% across {len(price_diff_pcts)} rows with both costs"
        )
    print(f"\nSaved priced workbook to: {OUT_FILE}")


if __name__ == "__main__":
    main()
