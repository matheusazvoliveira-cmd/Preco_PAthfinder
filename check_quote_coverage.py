"""
Check missing BOM items against Quotation workbook ('4 - Cotacao - Planilha Modelo - Vagao.xlsm'),
excluding Antenna Kit & PTC Kit sub-items (covered by Master kit prices) and Headlight items.

Output: 'Quote Coverage Report.xlsx'
"""
import subprocess
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent
QUOTE_FILE = BASE / "4 - Cotação - Planilha Modelo - Vagao.xlsm"
OUT_FILE = BASE / "Quote Coverage Report.xlsx"


def safe_copy(src):
    dst = Path(tempfile.gettempdir()) / src.name
    subprocess.run(
        ["powershell", "-NoProfile", "-Command", f'Copy-Item -Path "{src}" -Destination "{dst}" -Force'],
        check=True,
    )
    return dst


def norm(pn):
    if pn is None:
        return ""
    return str(pn).strip().upper()


def main():
    # 1. Load Quotation workbook (Approval Summary & Material Pricing)
    tmp_quote = safe_copy(QUOTE_FILE)
    wb_quote = openpyxl.load_workbook(tmp_quote, data_only=True, read_only=True)

    ws_app = wb_quote["Approval Summary"]
    rows_app = list(ws_app.iter_rows(values_only=True))
    header_app = list(rows_app[2])
    idx_app = {name: i for i, name in enumerate(header_app) if name}
    quote_prices = {}
    for r in rows_app[3:]:
        pn = norm(r[idx_app["PN"]]) if "PN" in idx_app else ""
        if pn:
            quote_prices[pn] = r[idx_app["Unit. Raw Cost"]]

    ws_mp = wb_quote["Material Pricing"]
    rows_mp = list(ws_mp.iter_rows(values_only=True))
    header_mp = list(rows_mp[2])
    idx_mp = {name: i for i, name in enumerate(header_mp) if name}
    mp_prices = {}
    for r in rows_mp[3:]:
        if r is None:
            continue
        pn = norm(r[idx_mp["PART NUMBER WABTEC"]]) if "PART NUMBER WABTEC" in idx_mp else ""
        if pn:
            mp_prices[pn] = r[idx_mp.get("RAW COST UNITÁRIO USD")]

    wb_quote.close()

    # 2. Load Pathfinder BOM
    wb1 = openpyxl.load_workbook(safe_copy(BASE / "Pathfinder BOM - Priced.xlsx"), data_only=True, read_only=True)
    ws1 = wb1["Priced BOM"]
    rows1 = list(ws1.iter_rows(values_only=True))
    header1 = list(rows1[0])
    idx1 = {name: i for i, name in enumerate(header1)}

    pathfinder_items = []
    for row in rows1[1:]:
        sub = str(row[idx1["Sub-System"]] or "").strip()
        pn = str(row[idx1["Part Number"]] or "").strip()
        comp = str(row[idx1["Component"]] or "").strip()
        desc = str(row[idx1["Description"]] or "").strip()
        notes = str(row[idx1["Notes"]] or "").strip()
        cost = row[idx1["Resolved Cost (US total)"]]
        source = str(row[idx1["Cost Source"]] or "").strip()
        pathfinder_items.append({
            "bomsheet": "Pathfinder BOM",
            "subsystem": sub,
            "pn": pn,
            "component": comp,
            "desc": desc,
            "notes": notes,
            "cost": cost,
            "source": source,
        })
    wb1.close()

    # 3. Load SKID BOM
    skid_items = []
    wb2 = openpyxl.load_workbook(safe_copy(BASE / "SKID BOM - Priced.xlsx"), data_only=True, read_only=True)
    for sheet in wb2.sheetnames:
        if sheet in ("Totals", "QUESTION"):
            continue
        ws = wb2[sheet]
        rows = list(ws.iter_rows(values_only=True))
        header = list(rows[0])
        idx = {name: i for i, name in enumerate(header)}
        if "Part Number" not in idx:
            continue
        for row in rows[1:]:
            if row is None or all(v is None for v in row):
                continue
            pn = str(row[idx["Part Number"]] or "").strip()
            name = str(row[idx.get("Name")] or "").strip() if idx.get("Name") is not None else ""
            cat = str(row[idx.get("Category")] or "").strip() if idx.get("Category") is not None else ""
            source = str(row[idx.get("Cost Source")] or "").strip()
            cost = row[idx.get("Extended Cost (USD)")] if idx.get("Extended Cost (USD)") is not None else row[idx.get("PAC Ref Cost (converted to USD, per unit)")]
            skid_items.append({
                "bomsheet": f"SKID ({sheet})",
                "subsystem": cat,
                "pn": pn,
                "component": name,
                "desc": "",
                "notes": "",
                "cost": cost,
                "source": source,
            })
    wb2.close()

    def is_excluded(item):
        sub_upper = item["subsystem"].upper()
        if sub_upper in ("ANTENNA KIT", "ANTENNA KIT MASTER", "PTC KIT", "PTC KIT MASTER"):
            return True, f"Covered by {sub_upper} Master price"
        blob = f"{item['subsystem']} {item['pn']} {item['component']} {item['desc']} {item['notes']}".lower()
        if "headlight" in blob or "hdlt" in blob or "smart box-v3-hlc" in blob:
            return True, "Headlight related item"
        return False, ""

    all_items = pathfinder_items + skid_items

    unpriced_list = []
    excluded_list = []
    seen_keys = set()

    for item in all_items:
        excluded, reason = is_excluded(item)
        if excluded:
            excluded_list.append((item, reason))
            continue

        pn_norm = norm(item["pn"])

        if item["source"] in ("INCLUDED", "NO_COST", "BOM", "PAC"):
            continue
        if item["cost"] not in (None, 0):
            continue

        q_cost = quote_prices.get(pn_norm)
        mp_cost = mp_prices.get(pn_norm)
        if q_cost not in (None, 0) or mp_cost not in (None, 0):
            continue

        key = (pn_norm, item["component"])
        if key not in seen_keys:
            seen_keys.add(key)
            unpriced_list.append(item)

    out_wb = openpyxl.Workbook()

    summary_ws = out_wb.active
    summary_ws.title = "Summary"
    summary_ws.append(["Metric", "Value"])
    summary_ws.append(["Total Pathfinder BOM Rows", len(pathfinder_items)])
    summary_ws.append(["Total SKID BOM Rows (excl QUESTION sheet)", len(skid_items)])
    summary_ws.append(["Excluded Rows (Antenna Kit / PTC Kit sub-items & Headlight)", len(excluded_list)])
    summary_ws.append(["TRULY UNPRICED DISTINCT ITEMS REMAINING", len(unpriced_list)])

    still_ws = out_wb.create_sheet("Still Unpriced")
    still_ws.append(["BOM Source Sheet", "Sub-System / Category", "Part Number", "Component Description"])
    for item in unpriced_list:
        still_ws.append([item["bomsheet"], item["subsystem"], item["pn"], item["component"]])

    ex_ws = out_wb.create_sheet("Excluded Kits & Headlight")
    ex_ws.append(["BOM Source Sheet", "Sub-System / Category", "Part Number", "Component Description", "Exclusion Reason"])
    for item, reason in excluded_list:
        ex_ws.append([item["bomsheet"], item["subsystem"], item["pn"], item["component"], reason])

    for ws in out_wb.worksheets:
        for col_idx in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = 35

    try:
        out_wb.save(OUT_FILE)
        print(f"\nSaved updated report to: {OUT_FILE}")
    except PermissionError:
        alt_file = BASE / "Quote Coverage Report - Updated.xlsx"
        out_wb.save(alt_file)
        print(f"\nPermissionError on primary file (it's open in Excel). Saved to: {alt_file}")

    print(f"Total truly unpriced distinct items remaining: {len(unpriced_list)}")


if __name__ == "__main__":
    main()
