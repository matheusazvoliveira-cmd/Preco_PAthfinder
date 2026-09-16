"""
Populate all available USD prices from Pathfinder BOM & SKID BOM (or PAC fallback converted to USD)
into '4 - Cotação - Planilha Modelo - Vagao.xlsm' (sheets 'Material Pricing' and 'Approval Summary').
"""
import os
import tempfile
from pathlib import Path

import openpyxl

BASE = Path(r"c:\Users\matheus.deoliveira\OneDrive - Wabtec Corporation\Documents\wabtec\Advanced Technology\MRS\Proposal")
QUOTE_FILE = BASE / "4 - Cotação - Planilha Modelo - Vagao.xlsm"
OUT_FILE = BASE / "4 - Cotação - Planilha Modelo - Vagao - Atualizada.xlsm"


def safe_copy_file(src):
    dst = Path(tempfile.gettempdir()) / src.name
    cmd = f'powershell -NoProfile -Command "Copy-Item -LiteralPath \'{src}\' -Destination \'{dst}\' -Force"'
    os.system(cmd)
    return dst


def norm(pn):
    if pn is None:
        return ""
    return str(pn).strip().upper()


def main():
    print("Step 1: Reading prices from BOM & PAC files...", flush=True)

    # 1. Collect USD prices from Pathfinder BOM - Priced.xlsx
    wb_p1 = openpyxl.load_workbook(safe_copy_file(BASE / "Pathfinder BOM - Priced.xlsx"), data_only=True, read_only=True)
    ws_p1 = wb_p1["Priced BOM"]
    rows_p1 = list(ws_p1.iter_rows(values_only=True))
    header_p1 = list(rows_p1[0])
    idx_p1 = {name: i for i, name in enumerate(header_p1)}

    bom_usd_map = {}

    for r in rows_p1[1:]:
        pn = norm(r[idx_p1["Part Number"]])
        if not pn:
            continue
        sub = str(r[idx_p1["Sub-System"]] or "").strip()
        if sub.upper() in ("ANTENNA KIT", "ANTENNA KIT MASTER", "PTC KIT", "PTC KIT MASTER"):
            continue
        comp = str(r[idx_p1["Component"]] or "").strip()
        desc = str(r[idx_p1["Description"]] or "").strip()
        notes = str(r[idx_p1["Notes"]] or "").strip()
        blob = f"{sub} {pn} {comp} {desc} {notes}".lower()
        if "headlight" in blob or "hdlt" in blob or "smart box-v3-hlc" in blob:
            continue

        source = r[idx_p1["Cost Source"]]
        res_cost = r[idx_p1["Resolved Cost (US total)"]]
        pac_usd = r[idx_p1["PAC Ref Cost (converted to USD)"]]

        usd_val = None
        if source == "BOM" and res_cost not in (None, 0):
            usd_val = res_cost
        elif source == "PAC" and pac_usd not in (None, 0):
            usd_val = pac_usd

        if usd_val not in (None, 0):
            bom_usd_map[pn] = {
                'cost': round(float(usd_val), 4),
                'source': source,
                'component': comp,
            }
    wb_p1.close()

    # 2. Collect USD prices from SKID BOM - Priced.xlsx
    wb_p2 = openpyxl.load_workbook(safe_copy_file(BASE / "SKID BOM - Priced.xlsx"), data_only=True, read_only=True)
    for sheet in wb_p2.sheetnames:
        if sheet in ("Totals", "QUESTION"):
            continue
        ws_p2 = wb_p2[sheet]
        rows_p2 = list(ws_p2.iter_rows(values_only=True))
        header_p2 = list(rows_p2[0])
        idx_p2 = {name: i for i, name in enumerate(header_p2)}
        if "Part Number" not in idx_p2:
            continue
        for r in rows_p2[1:]:
            if not r or all(v is None for v in r):
                continue
            pn = norm(r[idx_p2["Part Number"]])
            if not pn:
                continue
            cat = str(r[idx_p2.get("Category")] or "").strip() if idx_p2.get("Category") is not None else ""
            if cat.upper() in ("ANTENNA KIT", "ANTENNA KIT MASTER", "PTC KIT", "PTC KIT MASTER"):
                continue
            name = str(r[idx_p2.get("Name")] or "").strip()
            blob = f"{cat} {pn} {name}".lower()
            if "headlight" in blob or "hdlt" in blob or "smart box-v3-hlc" in blob:
                continue
            source = r[idx_p2.get("Cost Source")]
            pac_usd = r[idx_p2.get("PAC Ref Cost (converted to USD, per unit)")]
            if source == "PAC" and pac_usd not in (None, 0):
                if pn not in bom_usd_map:
                    bom_usd_map[pn] = {
                        'cost': round(float(pac_usd), 4),
                        'source': 'PAC',
                        'component': name,
                    }
    wb_p2.close()

    print(f"Total BOM/PAC items with valid USD prices: {len(bom_usd_map)}")

    # 3. Open Quotation workbook with keep_vba=True
    tmp_quote = safe_copy_file(QUOTE_FILE)
    print("Loading Quotation workbook...")
    wb_q = openpyxl.load_workbook(tmp_quote)
    print("Workbook loaded successfully!")

    ws_mp = wb_q["Material Pricing"]
    ws_app = wb_q["Approval Summary"]

    # Map existing PNs in Material Pricing
    existing_mp = {}
    for r in range(4, ws_mp.max_row + 1):
        pn_val = ws_mp.cell(r, 2).value
        pn = norm(pn_val)
        if pn:
            existing_mp[pn] = r

    # Perform updates on existing Material Pricing rows
    for u in updates_existing:
        ws_mp.cell(u['row'], 15).value = u['cost']

    # Perform inserts starting at row 96
    next_mp_row = 96
    next_app_row = 96
    next_item_num = 93  # Row 95 was item 92

    # Insert rows into Approval Summary before row 96 (TOTAL row) to avoid MergedCell conflict
    if inserts_new:
        print("Step 4: Inserting rows into Approval Summary worksheet...", flush=True)
        ws_app.insert_rows(96, amount=len(inserts_new))

    inserted_count = 0
    for item in inserts_new:
            ws_mp.cell(next_mp_row, 2).value = pn
            ws_mp.cell(next_mp_row, 3).value = f"=IFERROR(_xlfn.XLOOKUP(B{next_mp_row},'Cadastro de Itens'!B:B,'Cadastro de Itens'!F:F),\"{data['component']}\")"
            ws_mp.cell(next_mp_row, 4).value = f"=IFERROR(_xlfn.XLOOKUP(Table1[[#This Row],[PART NUMBER WABTEC]],'Cadastro de Itens'!B:B,'Cadastro de Itens'!D:D),\"\")"
            ws_mp.cell(next_mp_row, 5).value = f"=IFERROR(_xlfn.XLOOKUP(Table1[[#This Row],[PART NUMBER WABTEC]],'Cadastro de Itens'!B:B,'Cadastro de Itens'!N:N),\"\")"
            ws_mp.cell(next_mp_row, 6).value = "WRE"
            ws_mp.cell(next_mp_row, 7).value = "MG"
            ws_mp.cell(next_mp_row, 8).value = "MG"
            ws_mp.cell(next_mp_row, 9).value = f"=IFERROR(IF(_xlfn.XLOOKUP(Table1[[#This Row],[PART NUMBER WABTEC]],'Tipo Fiscal'!B:B,'Tipo Fiscal'!I:I)=\"C\",\"FGI\",\"UX\"),\"PREENCHA\")"
            ws_mp.cell(next_mp_row, 10).value = f"=G{next_mp_row}&H{next_mp_row}"
            ws_mp.cell(next_mp_row, 11).value = f"=G{next_mp_row}&H{next_mp_row}&E{next_mp_row}"
            ws_mp.cell(next_mp_row, 12).value = f"=J{next_mp_row}&D{next_mp_row}"
            ws_mp.cell(next_mp_row, 14).value = f"=Table1[[#This Row],[LEAD TIME ORACLE]]+60"
            ws_mp.cell(next_mp_row, 15).value = data['cost']
            ws_mp.cell(next_mp_row, 16).value = f"=IFERROR(_xlfn.XLOOKUP(Table1[[#This Row],[WBT OF ORIGIN]],RATES!$B$15:$B$18,RATES!$C$15:$C$18),\"\")"
            ws_mp.cell(next_mp_row, 17).value = f"=Table1[[#This Row],[RAW COST UNITÁRIO USD]]*Table1[[#This Row],[TRANSFER PRICE %]]"
            ws_mp.cell(next_mp_row, 18).value = f"=O{next_mp_row}*(1+P{next_mp_row})"
            ws_mp.cell(next_mp_row, 19).value = "=RATES!$C$11"

            # Insert corresponding row in Approval Summary
            ws_app.cell(next_app_row, 1).value = next_item_num
            ws_app.cell(next_app_row, 2).value = f"='Material Pricing'!B{next_mp_row}"
            ws_app.cell(next_app_row, 3).value = f"=_xlfn.XLOOKUP(B{next_app_row},'Material Pricing'!B:B,'Material Pricing'!C:C,\"\")"
            ws_app.cell(next_app_row, 4).value = 1
            ws_app.cell(next_app_row, 5).value = f"=_xlfn.XLOOKUP(B{next_app_row},'Material Pricing'!B:B,'Material Pricing'!O:O,\"\")"
            ws_app.cell(next_app_row, 6).value = f"=IFERROR(E{next_app_row}*D{next_app_row},\"\")"
            ws_app.cell(next_app_row, 7).value = f"=_xlfn.XLOOKUP(B{next_app_row},'Material Pricing'!B:B,'Material Pricing'!R:R,\"\")"
            ws_app.cell(next_app_row, 8).value = f"=IFERROR(G{next_app_row}*$D{next_app_row},\"\")"
            ws_app.cell(next_app_row, 9).value = f"=IFERROR((H{next_app_row}-F{next_app_row})/H{next_app_row},0)"
            ws_app.cell(next_app_row, 10).value = f"=IFERROR(Table1[[#This Row],[FRETE NACIONAL UNITÁRIO USD]]+Table1[[#This Row],[II UNITÁRIO USD]]+Table1[[#This Row],[FRETE INTERNACIONAL UNITÁRIO USD]],\"\")"
            ws_app.cell(next_app_row, 11).value = f"=IFERROR(J{next_app_row}*$D{next_app_row},\"\")"
            ws_app.cell(next_app_row, 12).value = f"=IFERROR(((N{next_app_row}-H{next_app_row})-K{next_app_row})/(N{next_app_row}-H{next_app_row}),0)"
            ws_app.cell(next_app_row, 13).value = f"=_xlfn.XLOOKUP(B{next_app_row},'Material Pricing'!B:B,'Material Pricing'!AB:AB,\"\")"
            ws_app.cell(next_app_row, 14).value = f"=IFERROR(M{next_app_row}*D{next_app_row},\"\")"
            ws_app.cell(next_app_row, 15).value = f"=IFERROR(((H{next_app_row}-F{next_app_row})+(N{next_app_row}-K{next_app_row}-H{next_app_row}))/N{next_app_row},0)"

            next_mp_row += 1
            next_app_row += 1
            next_item_num += 1
            inserted_count += 1

    print(f"\nUpdated {len(updates_existing)} existing rows with USD costs.")
    print(f"Inserted {inserted_count} new rows into Material Pricing and Approval Summary.")

    # Save output
    wb_q.save(OUT_FILE)
    print(f"Saved updated quotation workbook to: {OUT_FILE}")

    # Try saving directly to main file if not locked
    try:
        wb_q.save(QUOTE_FILE)
        print(f"Also updated main quotation file directly: {QUOTE_FILE}")
    except PermissionError:
        print(f"(Main file '{QUOTE_FILE.name}' is open in Excel, so changes were saved to '{OUT_FILE.name}')")


if __name__ == "__main__":
    main()
