"""Adds a "Simulation" / "Simulação" tab to the final summary workbooks.

Switches (yellow cells) drive one "selected scenario" column; a fixed 4-scenario
comparison sits next to it so the message is visible without touching anything:
  Base | No import duty (II) | 1 battery set | No II + 1 battery set

The base numbers are read from the workbook's own P&L sheet (rows 10-15 and B3:B7),
so the simulation always reconciles with the P&L the team already reviewed.
Only the battery share is derived here (from the Final Cost List in the working file).

Run: PYTHONIOENCODING=utf-8 python add_simulation_sheet.py EN|PT
"""
import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent
LANG = sys.argv[1] if len(sys.argv) > 1 else "EN"
TARGET = BASE / f"Pathfinder_Cost_Summary_{LANG}.xlsx"
SRC = BASE / "Lista_Final_Pathfinder_SKID_Custos_English_Updated.xlsx"
REF_SRC = BASE / "Planilha de Custos de Importacao - Pathfinder completo- Set 26.xlsx"

OLD_INC, OLD_EXC = 1.526333104, 1.302332
TP_HAD, TP_NONE = 0.13, 0.172
INTL_FREIGHT = 0.012295642058762843
DEFAULT_II = 0.18
BATT_PN = "84A222059P5"

T = {
    "EN": dict(
        sheet="Simulation", pl_hint="P&L",
        title="SIMULATION - what really makes the Pathfinder Iron Ore Car expensive",
        sub="Change the two yellow switches; the comparison table on the right never depends on them.",
        sw_title="SWITCHES", sw_ii="Remove import duty (II) from everything?  (1 = yes, 0 = no)",
        sw_bat="Battery sets  (2 = current BOM, 1 = a single set)",
        base_title="BASE VALUES (read from the P&L tab)",
        fx="Exchange rate", tdm="Total Deal Margin (TDM)", pc="PIS/COFINS", icms="ICMS", ipi="IPI",
        us_ref="US domestic reference price (USD, input)",
        b_exb="Imported items, excluding batteries", b_bat="Batteries - value per set (current BOM has 2)",
        lines=["Raw cost - imported items", "Raw cost - PAC items", "TP - Transfer Price", "International freight",
               "II - Import duty", "Domestic freight", "Supermetal - refurbishment / integration"],
        rows=["FINAL COST (with TP):", "Base for margin (Final Cost - TP)", "Net price w/o TP (Base / (1 - TDM))",
              "NET PRICE ('OUR PRICE'):", "+ PIS/COFINS (gross-up)", "+ ICMS (gross-up)", "+ IPI (on top)",
              "FINAL CLIENT PRICE:"],
        cols=["Selected (switches)", "Base (current)", "No import duty (II)", "1 battery set", "No II + 1 battery set"],
        usd="USD", brl="BRL",
        drv_title="WHAT DRIVES THE FINAL CLIENT PRICE  (base case)",
        drv_rows=["Removing ALL import duty (II) lowers the Final Client Price by",
                  "Local output taxes (PIS/COFINS + ICMS + IPI) add on top of Net Price",
                  "Local taxes are how many times bigger than the II effect?",
                  "Final Client Price with NO II and 1 battery set, vs. US reference"],
        notes=["II (import duty) sits inside the cost, so it is multiplied by the margin (1/(1-TDM)) and then by the local tax cascade - removing it is worth about 2.7x its face value in the final price.",
               "Even so, the local output taxes (PIS/COFINS, ICMS, IPI) stack on the whole price and outweigh the II effect - that is the point of this tab.",
               "Battery sets: II, TP and freight of the batteries scale with the quantity; the row values come from the Final Cost List (84A222059P5, group without historical TP -> 17.2%, no confirmed NCM -> II 18%)."],
        ref_title="6) REFERENCE: FINISHED US UNIT IMPORTED AS-IS (buy the ready-built US$900k Pathfinder and land it in Brazil)",
        ref_rows=["US price (finished product, unchanged)", "+ International freight + insurance",
                  "+ Customs / import expenses", "+ All import taxes (II, IPI, PIS/COFINS, ICMS)",
                  "LANDED IN BRAZIL:"],
        ref_fx_lab="Exchange rate used for this reference (differs from the scenario FX above)",
        ref_note="Source: read LIVE from 'Planilha de Custos de Importacao - Pathfinder completo- Set 26.xlsx', sheet "
                 "'Pathfinder Pronto' (cells H24, I22, I26, I27, I28, I35, I56). Different scope than the 4 scenarios "
                 "above: this is a finished, already-built unit (no Wabtec Brasil margin, no TP, no Supermetal "
                 "refurbishment - it's a straight import of a complete product), so it is not an apples-to-apples 5th "
                 "column of the scenario table.",
        ref_vs="vs. each scenario's Final Client Price (positive = landed-finished-unit costs more)",
        regime_title="7) SCENARIO: SUPERMETAL SPECIAL REGIME (Material: ICMS 1.5%/IPI 0%; Service: ISS instead of ICMS/IPI)",
        regime_note="Same split and rates as Pathfinder_Cost_Summary_PT_Cenario_B_RegimeEspecial.xlsx (build_scenario_b.py) - "
                    "material R$432,500 (incl. R$20,000 transport) + service R$275,290.60, from 'Pathfinder GDT MRS - Cost "
                    "Rollup - Rev0.xlsx' ('Container in a GDT' sheet). Uses the Base scenario's other costs (2 battery "
                    "sets, II in) - not yet combined with the No-II/battery switches above. ICMS 1.5%/IPI 0% and ISS 5% "
                    "are still pending tax confirmation - see the Fonte column in the Cenario B file.",
        regime_icms="ICMS - Supermetal special regime (material)", regime_ipi="IPI - Supermetal special regime (material)",
        regime_iss="ISS on the service portion (placeholder)",
        summary_title="8) SUMMARY - ALL SCENARIOS SIDE BY SIDE",
        summary_h=["Scenario", "Net Price (Our Price) USD", "Final Client Price USD"],
        summary_rows=["1  Base (2 battery sets, II in)", "2  No import duty (II = 0)", "3  One battery set (II in)",
                      "4  No II + one battery set", "5  Supermetal special regime (Base battery/II)",
                      "6  US reference: finished unit landed in Brazil"],
    ),
    "PT": dict(
        sheet="Simulação", pl_hint="P&L",
        title="SIMULAÇÃO - o que realmente encarece o Pathfinder Iron Ore Car",
        sub="Mude as duas chaves amarelas; a tabela de comparação à direita nunca depende delas.",
        sw_title="CHAVES", sw_ii="Retirar o imposto de importação (II) de tudo?  (1 = sim, 0 = não)",
        sw_bat="Conjuntos de bateria  (2 = BOM atual, 1 = um único conjunto)",
        base_title="VALORES-BASE (lidos da aba P&L)",
        fx="Câmbio", tdm="Total Deal Margin (TDM)", pc="PIS/COFINS", icms="ICMS", ipi="IPI",
        us_ref="Preço de referência doméstico EUA (USD, input)",
        b_exb="Itens importados, exceto baterias", b_bat="Baterias - valor por conjunto (a BOM atual tem 2)",
        lines=["Raw cost - itens importados", "Raw cost - itens PAC", "TP - Transfer Price", "Frete internacional",
               "II - Imposto de importação", "Frete nacional", "Supermetal - reforma / integração"],
        rows=["CUSTO FINAL (com TP):", "Base p/ margem (Custo Final - TP)", "Preço líquido s/ TP (Base / (1 - TDM))",
              "PREÇO LÍQUIDO ('NOSSO PREÇO'):", "+ PIS/COFINS (por dentro)", "+ ICMS (por dentro)", "+ IPI (por fora)",
              "PREÇO FINAL AO CLIENTE:"],
        cols=["Selecionado (chaves)", "Base (atual)", "Sem imposto de importação (II)", "1 conjunto de bateria", "Sem II + 1 conjunto de bateria"],
        usd="USD", brl="BRL",
        drv_title="O QUE PESA NO PREÇO FINAL AO CLIENTE  (caso base)",
        drv_rows=["Retirar TODO o imposto de importação (II) reduz o Preço Final ao Cliente em",
                  "Impostos locais de saída (PIS/COFINS + ICMS + IPI) somados ao Preço Líquido",
                  "Os impostos locais são quantas vezes maiores que o efeito do II?",
                  "Preço Final ao Cliente SEM II e com 1 conjunto de bateria, vs. referência EUA"],
        notes=["O II está dentro do custo, então é multiplicado pela margem (1/(1-TDM)) e depois pela cascata de impostos locais - retirar o II vale cerca de 2,7x o valor dele no preço final.",
               "Mesmo assim, os impostos locais de saída (PIS/COFINS, ICMS, IPI) incidem sobre o preço inteiro e pesam mais que o efeito do II - esse é o ponto desta aba.",
               "Conjuntos de bateria: II, TP e frete das baterias escalam com a quantidade; os valores vêm da Final Cost List (84A222059P5, grupo sem TP histórico -> 17,2%, sem NCM confirmada -> II 18%)."],
        ref_title="6) REFERÊNCIA: UNIDADE PRONTA DOS EUA IMPORTADA COMO ESTÁ (comprar o Pathfinder pronto de US$900 mil e desembaraçar no Brasil)",
        ref_rows=["Preço EUA (produto pronto, inalterado)", "+ Frete internacional + seguro",
                  "+ Despesas aduaneiras / importação", "+ Todos os impostos de importação (II, IPI, PIS/COFINS, ICMS)",
                  "POSTO NO BRASIL:"],
        ref_fx_lab="Câmbio usado nessa referência (diferente do câmbio dos cenários acima)",
        ref_note="Fonte: lido AO VIVO de 'Planilha de Custos de Importacao - Pathfinder completo- Set 26.xlsx', aba "
                 "'Pathfinder Pronto' (celulas H24, I22, I26, I27, I28, I35, I56). Escopo diferente dos 4 cenários "
                 "acima: é uma unidade pronta e já fabricada (sem margem da Wabtec Brasil, sem TP, sem reforma da "
                 "Supermetal - é a importação direta de um produto completo), então não é uma 5ª coluna comparável "
                 "1-pra-1 com a tabela de cenários.",
        ref_vs="vs. o Preço Final ao Cliente de cada cenário (positivo = a unidade pronta importada custa mais)",
        regime_title="7) CENÁRIO: REGIME ESPECIAL SUPERMETAL (Material: ICMS 1,5%/IPI 0%; Serviço: ISS em vez de ICMS/IPI)",
        regime_note="Mesma divisão e taxas do Pathfinder_Cost_Summary_PT_Cenario_B_RegimeEspecial.xlsx (build_scenario_b.py) - "
                    "material R$432.500 (incl. R$20.000 de transporte) + serviço R$275.290,60, vindos de 'Pathfinder GDT "
                    "MRS - Cost Rollup - Rev0.xlsx' (aba 'Container in a GDT'). Usa os demais custos do cenário Base (2 "
                    "conjuntos de bateria, II incluído) - ainda não combinado com as chaves de II/bateria acima. ICMS "
                    "1,5%/IPI 0% e ISS 5% ainda dependem de confirmação do time de tax - ver coluna Fonte no arquivo "
                    "Cenário B.",
        regime_icms="ICMS - regime especial Supermetal (material)", regime_ipi="IPI - regime especial Supermetal (material)",
        regime_iss="ISS sobre a parcela de serviço (placeholder)",
        summary_title="8) RESUMO - TODOS OS CENÁRIOS LADO A LADO",
        summary_h=["Cenário", "Preço Líquido (Nosso Preço) USD", "Preço Final ao Cliente USD"],
        summary_rows=["1  Base (2 conjuntos de bateria, II incluído)", "2  Sem imposto de importação (II = 0)",
                      "3  Um conjunto de bateria (II incluído)", "4  Sem II + um conjunto de bateria",
                      "5  Regime especial Supermetal (bateria/II do Base)",
                      "6  Referência EUA: unidade pronta posta no Brasil"],
    ),
}[LANG]

HF = PatternFill(start_color="FF1F497D", end_color="FF1F497D", fill_type="solid")
HFONT = Font(bold=True, color="FFFFFFFF")
INP = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
KEY = PatternFill(start_color="FFE2EFDA", end_color="FFE2EFDA", fill_type="solid")
BRL = '\\R\\$\\ #,##0.00'
USD = '\\$\\ #,##0.00'
PCT = '0.00%'
BOLD = Font(bold=True)
thin = Side(style="thin", color="FFB0B0B0")
BOR = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------- battery share from the working file (source of truth for item level) ----------
src_ws = openpyxl.load_workbook(SRC, data_only=False)["Final Cost List"]
FX_SRC = 5.10
batt = None
for r in src_ws.iter_rows(min_row=2, values_only=True):
    if r[0] and str(r[0]).strip() == BATT_PN and r[5] and r[7]:
        qty = r[4] or 1
        ratio = r[7] / (r[5] * qty)
        tp_rate = TP_HAD if abs(ratio - OLD_INC) < abs(ratio - OLD_EXC) else TP_NONE
        raw = r[5] * qty * FX_SRC
        tp = raw * tp_rate
        fr = (raw + tp) * INTL_FREIGHT
        ii = (raw + tp + fr) * DEFAULT_II
        batt = dict(qty=qty, raw=raw / qty, tp=tp / qty, fr=fr / qty, ii=ii / qty)
        break
assert batt, "battery row not found in Final Cost List"
print("battery per set:", {k: round(v, 2) for k, v in batt.items()}, "BOM qty:", batt["qty"])

# ---------- read the P&L sheet of the target workbook ----------
wb = openpyxl.load_workbook(TARGET, data_only=False)
pl_name = next(n for n in wb.sheetnames if "P&L" in n)
pl = wb[pl_name]
def pl_val(row):
    v = pl.cell(row=row, column=2).value
    assert isinstance(v, (int, float)), f"P&L B{row} is not a number: {v!r}"
    return float(v)
PLQ = f"'{pl_name}'"
base_vals = [pl_val(r) for r in range(10, 17)]  # raw imp, PAC, TP, intl freight, II, dom freight, Supermetal
print("P&L base rows 10-16:", [round(v, 2) for v in base_vals])
print("P&L params B3..B7:", [pl.cell(row=r, column=2).value for r in range(3, 8)])

if T["sheet"] in wb.sheetnames:
    del wb[T["sheet"]]
ws = wb.create_sheet(T["sheet"])
SQ = f"'{T['sheet']}'"

ws["A1"] = T["title"]
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = T["sub"]
ws["A2"].font = Font(italic=True, color="FF808080", size=9)

# switches
ws["A4"] = T["sw_title"]; ws["A4"].font = BOLD
ws["A5"] = T["sw_ii"]; ws["B5"] = 0
ws["A6"] = T["sw_bat"]; ws["B6"] = batt["qty"]
for c in ("B5", "B6"):
    ws[c].fill, ws[c].border, ws[c].font = INP, BOR, BOLD

# params linked to the P&L tab
ws["A8"] = T["base_title"]; ws["A8"].font = BOLD
params = [(T["fx"], f"={PLQ}!B3", '0.00'), (T["tdm"], f"={PLQ}!B4", PCT), (T["pc"], f"={PLQ}!B5", PCT),
          (T["icms"], f"={PLQ}!B6", PCT), (T["ipi"], f"={PLQ}!B7", PCT), (T["us_ref"], 900000, USD)]
for i, (lab, val, fmt) in enumerate(params, start=9):
    ws.cell(row=i, column=1, value=lab)
    c = ws.cell(row=i, column=2, value=val)
    c.number_format = fmt
    if lab == T["us_ref"]:
        c.fill, c.border = INP, BOR
FX, TDM, PC, ICMS, IPI, USREF = [f"$B${i}" for i in range(9, 15)]

# helper block (bottom of the sheet): value excluding batteries, and battery value per set
HELP0 = 48
ws.cell(row=HELP0 - 2, column=1, value=T["base_title"]).font = BOLD
ws.cell(row=HELP0 - 1, column=2, value=T["b_exb"]).font = BOLD
ws.cell(row=HELP0 - 1, column=3, value=T["b_bat"]).font = BOLD
for c in (2, 3):
    ws.cell(row=HELP0 - 1, column=c).alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[HELP0 - 1].height = 48
line_names = T["lines"]
helper_rows = {}
for i, name in enumerate(line_names):
    r = HELP0 + i
    ws.cell(row=r, column=1, value=name)
    helper_rows[i] = r
# raw(0), tp(2), intl freight(3), II(4) depend on the battery quantity
per_set = {0: batt["raw"], 2: batt["tp"], 3: batt["fr"], 4: batt["ii"]}
for i in range(len(line_names)):
    r = helper_rows[i]
    if i in per_set:
        ws.cell(row=r, column=3, value=round(per_set[i], 2)).number_format = BRL
        ws.cell(row=r, column=2, value=f"={PLQ}!B{10 + i}-{batt['qty']}*C{r}").number_format = BRL
    else:
        ws.cell(row=r, column=2, value=f"={PLQ}!B{10 + i}").number_format = BRL
        ws.cell(row=r, column=3, value=0).number_format = BRL

# ---------- scenario table ----------
hr = 18
ws.cell(row=hr - 1, column=1, value=("SCENARIOS" if LANG == "EN" else "CENÁRIOS")).font = Font(bold=True, size=12)
scen = [  # (noII expr, battery qty expr)
    (f"$B$5", f"$B$6"),
    ("0", str(batt["qty"])),
    ("1", str(batt["qty"])),
    ("0", "1"),
    ("1", "1"),
]
for j, name in enumerate(T["cols"]):
    for k, unit in enumerate((T["brl"], T["usd"])):
        c = ws.cell(row=hr, column=2 + j * 2 + k, value=f"{name} - {unit}")
        c.fill, c.font = HF, HFONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[hr].height = 48
n_lines = len(line_names)
labels = line_names + T["rows"]
for i, lab in enumerate(labels):
    ws.cell(row=hr + 1 + i, column=1, value=lab)
    if lab.endswith(":"):
        ws.cell(row=hr + 1 + i, column=1).font = BOLD
R = lambda i: hr + 1 + i
ROW_COST, ROW_BASE, ROW_NSTP, ROW_NET, ROW_PC, ROW_ICMS, ROW_IPI, ROW_FIN = [n_lines + k for k in range(8)]
for j, (noii, bq) in enumerate(scen):
    bc, uc = get_column_letter(2 + j * 2), get_column_letter(3 + j * 2)
    for i in range(n_lines):
        hr_ = helper_rows[i]
        if i in per_set:
            f = f"=$B${hr_}+$C${hr_}*({bq})"
            if i == 4:
                f = f"=($B${hr_}+$C${hr_}*({bq}))*(1-({noii}))"
        else:
            f = f"=$B${hr_}"
        ws[f"{bc}{R(i)}"] = f
    ws[f"{bc}{R(ROW_COST)}"] = f"=SUM({bc}{R(0)}:{bc}{R(n_lines - 1)})"
    ws[f"{bc}{R(ROW_BASE)}"] = f"={bc}{R(ROW_COST)}-{bc}{R(2)}"
    ws[f"{bc}{R(ROW_NSTP)}"] = f"={bc}{R(ROW_BASE)}/(1-{TDM})"
    ws[f"{bc}{R(ROW_NET)}"] = f"={bc}{R(ROW_NSTP)}+{bc}{R(2)}"
    ws[f"{bc}{R(ROW_PC)}"] = f"={bc}{R(ROW_NET)}/(1-{PC})"
    ws[f"{bc}{R(ROW_ICMS)}"] = f"={bc}{R(ROW_PC)}/(1-{ICMS})"
    ws[f"{bc}{R(ROW_IPI)}"] = f"={bc}{R(ROW_ICMS)}*{IPI}"
    ws[f"{bc}{R(ROW_FIN)}"] = f"={bc}{R(ROW_ICMS)}+{bc}{R(ROW_IPI)}"
    for i in range(len(labels)):
        ws[f"{bc}{R(i)}"].number_format = BRL
        ws[f"{uc}{R(i)}"] = f"={bc}{R(i)}/{FX}"
        ws[f"{uc}{R(i)}"].number_format = USD
    for i in (ROW_COST, ROW_NET, ROW_FIN):
        ws[f"{bc}{R(i)}"].font = ws[f"{uc}{R(i)}"].font = Font(bold=True, size=12)
    if j == 0:
        for i in range(len(labels)):
            ws[f"{bc}{R(i)}"].fill = KEY
            ws[f"{uc}{R(i)}"].fill = KEY

# ---------- drivers ----------
dr = R(len(labels)) + 2
ws.cell(row=dr, column=1, value=T["drv_title"]).font = Font(bold=True, size=12)
base_fin_b, noii_fin_b, both_fin_b = (f"{get_column_letter(2 + j * 2)}{R(ROW_FIN)}" for j in (1, 2, 4))
base_net_b = f"{get_column_letter(4)}{R(ROW_NET)}"
base_net_b = f"D{R(ROW_NET)}"  # base scenario is column index j=1 -> B is selected, D is base
base_fin_b, noii_fin_b, both_fin_b = f"D{R(ROW_FIN)}", f"F{R(ROW_FIN)}", f"J{R(ROW_FIN)}"
drivers = [
    (f"={base_fin_b}-{noii_fin_b}", f"=B{dr + 1}/{FX}"),
    (f"={base_fin_b}-{base_net_b}", f"=B{dr + 2}/{FX}"),
    (f"=B{dr + 2}/B{dr + 1}", None),
    (f"={both_fin_b}/{FX}", f"=B{dr + 4}/{USREF}-1"),
]
for k, (lab, (f1, f2)) in enumerate(zip(T["drv_rows"], drivers), start=1):
    ws.cell(row=dr + k, column=1, value=lab)
    c = ws.cell(row=dr + k, column=2, value=f1)
    c.number_format = BRL if k in (1, 2) else ('0.0"x"' if k == 3 else USD)
    if f2:
        c2 = ws.cell(row=dr + k, column=3, value=f2)
        c2.number_format = USD if k in (1, 2) else PCT
    for cc in (ws.cell(row=dr + k, column=1), c):
        cc.font = BOLD
for n, note in enumerate(T["notes"]):
    c = ws.cell(row=dr + 6 + n, column=1, value="- " + note)
    c.font = Font(italic=True, color="FF808080", size=9)
    ws.merge_cells(start_row=dr + 6 + n, start_column=1, end_row=dr + 6 + n, end_column=8)

# ---------- section 6: finished-US-unit reference (read live from the China Gate "completo" workbook) ----------
_ref_wb = openpyxl.load_workbook(REF_SRC, data_only=True, read_only=True)
_ref = _ref_wb["Pathfinder Pronto"]
assert _ref["H24"].value == 900000, f"US sticker price changed in {REF_SRC.name}: {_ref['H24'].value!r}"
REF_FX = _ref["I22"].value
_i26, _i27, _i28, _i35, _i56 = (_ref[c].value for c in ("I26", "I27", "I28", "I35", "I56"))
assert all(isinstance(v, (int, float)) for v in (REF_FX, _i26, _i27, _i28, _i35, _i56)), \
    f"Non-numeric cell in {REF_SRC.name}!Pathfinder Pronto"
_ref_wb.close()
REF_PROD = _i26 / REF_FX
REF_FRT = (_i27 + _i28) / REF_FX
REF_EXP = _i35 / REF_FX
REF_TAX = (_i56 - _i26 - _i27 - _i28 - _i35) / REF_FX
REF_TOTAL = REF_PROD + REF_FRT + REF_EXP + REF_TAX
print(f"Referencia EUA lida ao vivo de {REF_SRC.name}: FX={REF_FX} produto={REF_PROD:.0f} "
      f"frete+seguro={REF_FRT:.0f} despesas={REF_EXP:.0f} impostos={REF_TAX:.0f} total={REF_TOTAL:.0f}")

rr = max(dr + 6 + len(T["notes"]) + 2, HELP0 + len(line_names) + 3)  # must clear the helper block (HELP0..HELP0+len(line_names))
ws.cell(row=rr, column=1, value=T["ref_title"]).font = Font(bold=True, size=12)
rr += 1

ws.cell(row=rr, column=1, value=T["ref_fx_lab"])
c = ws.cell(row=rr, column=2, value=REF_FX)
c.number_format = '0.00'
c.fill, c.border, c.font = INP, BOR, BOLD
REF_FX_CELL = f"$B${rr}"
rr += 1

ws.cell(row=rr, column=2, value=T["brl"]).font = Font(italic=True, size=8, color="FF808080")
ws.cell(row=rr, column=3, value=T["usd"]).font = Font(italic=True, size=8, color="FF808080")
rr += 1

ref_row0 = rr
for i, lab in enumerate(T["ref_rows"]):
    ws.cell(row=rr, column=1, value=lab)
    if lab.endswith(":"):
        ws.cell(row=rr, column=1).font = BOLD
    rr += 1
ws.cell(row=ref_row0, column=3, value=REF_PROD).number_format = USD
ws.cell(row=ref_row0 + 1, column=3, value=REF_FRT).number_format = USD
ws.cell(row=ref_row0 + 2, column=3, value=REF_EXP).number_format = USD
ws.cell(row=ref_row0 + 3, column=3, value=REF_TAX).number_format = USD
tot_cell = ws.cell(row=ref_row0 + 4, column=3, value=f"=SUM(C{ref_row0}:C{ref_row0 + 3})")
tot_cell.number_format = USD
tot_cell.font = Font(bold=True, size=12)
for i in range(5):
    bcell = ws.cell(row=ref_row0 + i, column=2, value=f"=C{ref_row0 + i}*{REF_FX_CELL}")
    bcell.number_format = BRL
    if i == 4:
        bcell.font = Font(bold=True, size=12)
rr += 1

ws.cell(row=rr, column=1, value=T["ref_vs"]).font = BOLD
rr += 1
for j, scen_name in enumerate(T["cols"][1:], start=1):  # skip "Selected", compare vs the 4 fixed scenarios
    fin_col = get_column_letter(3 + j * 2)  # USD column of each scenario's Final Client Price row
    ws.cell(row=rr, column=1, value=scen_name)
    c = ws.cell(row=rr, column=2, value=f"=C{ref_row0 + 4}-{fin_col}{R(ROW_FIN)}")
    c.number_format = USD
    rr += 1

rr += 1
ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=8)
note_cell = ws.cell(row=rr, column=1, value="- " + T["ref_note"])
note_cell.font = Font(italic=True, color="FFCC0000", size=9)
note_cell.alignment = Alignment(wrap_text=True)
ws.row_dimensions[rr].height = 60
rr += 2

# ---------- section 7: Supermetal special regime scenario (same split/rates as Cenario B) ----------
MAT_TOTAL, SVC_TOTAL = 432500.0, 275290.6  # matches build_scenario_b.py sections 1 and 3
ws.cell(row=rr, column=1, value=T["regime_title"]).font = Font(bold=True, size=12)
rr += 1
regime_inputs = [(T["regime_icms"], 0.015), (T["regime_ipi"], 0.0), (T["regime_iss"], 0.05)]
for lab, val in regime_inputs:
    ws.cell(row=rr, column=1, value=lab)
    c = ws.cell(row=rr, column=2, value=val)
    c.number_format = PCT
    c.fill, c.border, c.font = INP, BOR, BOLD
    rr += 1
ICMS_REGIME_CELL, IPI_REGIME_CELL, ISS_CELL = (f"$B${rr - 3}", f"$B${rr - 2}", f"$B${rr - 1}")

# non-Supermetal part reuses the Base scenario's own cost lines (column D = BRL); row by row, same
# pattern as build_scenario_b.py, so each intermediate step stays inspectable instead of one giant formula.
def calc_row(label, formula, bold=False):
    global rr
    ws.cell(row=rr, column=1, value=label)
    c = ws.cell(row=rr, column=2, value=formula)
    c.number_format = BRL
    ws.cell(row=rr, column=3, value=f"=B{rr}/{FX}").number_format = USD
    if bold:
        c.font = ws.cell(row=rr, column=3).font = Font(bold=True, size=12)
    row_used = rr
    rr += 1
    return row_used

nonsm_cost_row = calc_row("Custo (sem Supermetal)" if LANG == "PT" else "Cost (excl. Supermetal)", "=D19+D20+D21+D22+D23+D24")
nonsm_base_row = calc_row("Base p/ margem" if LANG == "PT" else "Base for margin", f"=B{nonsm_cost_row}-D21")
nonsm_net_row = calc_row("Preço líquido (sem Supermetal)" if LANG == "PT" else "Net price (excl. Supermetal)",
                          f"=B{nonsm_base_row}/(1-{TDM})+D21")
nonsm_pc_row = calc_row("+ PIS/COFINS", f"=B{nonsm_net_row}/(1-{PC})")
nonsm_icms_row = calc_row("+ ICMS (18%)", f"=B{nonsm_pc_row}/(1-{ICMS})")
nonsm_fin_row = calc_row("Preço final (sem Supermetal):" if LANG == "PT" else "Final price (excl. Supermetal):",
                          f"=B{nonsm_icms_row}*(1+{IPI})", bold=True)

mat_net_row = calc_row("Material com margem" if LANG == "PT" else "Material with margin", f"={MAT_TOTAL}/(1-{TDM})")
mat_pc_row = calc_row("+ PIS/COFINS (material)", f"=B{mat_net_row}/(1-{PC})")
mat_icms_row = calc_row("+ ICMS regime especial", f"=B{mat_pc_row}/(1-{ICMS_REGIME_CELL})")
mat_fin_row = calc_row("Preço final (material):" if LANG == "PT" else "Final price (material):",
                        f"=B{mat_icms_row}*(1+{IPI_REGIME_CELL})", bold=True)

svc_net_row = calc_row("Serviço com margem" if LANG == "PT" else "Service with margin", f"={SVC_TOTAL}/(1-{TDM})")
svc_pc_row = calc_row("+ PIS/COFINS (serviço)", f"=B{svc_net_row}/(1-{PC})")
svc_fin_row = calc_row("Preço final (serviço, com ISS):" if LANG == "PT" else "Final price (service, with ISS):",
                        f"=B{svc_pc_row}*(1+{ISS_CELL})", bold=True)

regime_net_row = calc_row(T["rows"][3], f"=B{nonsm_net_row}+B{mat_net_row}+B{svc_net_row}", bold=True)
regime_fin_row = calc_row(T["rows"][-1], f"=B{nonsm_fin_row}+B{mat_fin_row}+B{svc_fin_row}", bold=True)
rr += 1

ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=8)
rnote = ws.cell(row=rr, column=1, value="- " + T["regime_note"])
rnote.font = Font(italic=True, color="FF808080", size=9)
rnote.alignment = Alignment(wrap_text=True)
ws.row_dimensions[rr].height = 60
rr += 2

# ---------- section 8: all-scenarios summary table ----------
ws.cell(row=rr, column=1, value=T["summary_title"]).font = Font(bold=True, size=12)
rr += 1
for col, h in enumerate(T["summary_h"], start=1):
    hc = ws.cell(row=rr, column=col, value=h)
    hc.fill, hc.font = HF, HFONT
    hc.alignment = Alignment(wrap_text=True, vertical="center")
rr += 1
summary_net_refs = [f"D{R(ROW_NET)}", f"F{R(ROW_NET)}", f"H{R(ROW_NET)}", f"J{R(ROW_NET)}",
                    f"B{regime_net_row}", f"C{ref_row0 + 4}"]
summary_fin_refs = [f"D{R(ROW_FIN)}", f"F{R(ROW_FIN)}", f"H{R(ROW_FIN)}", f"J{R(ROW_FIN)}",
                    f"B{regime_fin_row}", f"C{ref_row0 + 4}"]
summary_is_brl = [True, True, True, True, True, False]  # section-6 reference's C column is already USD
for i, name in enumerate(T["summary_rows"]):
    ws.cell(row=rr, column=1, value=name)
    net_ref, fin_ref = summary_net_refs[i], summary_fin_refs[i]
    net_formula = f"={net_ref}/{FX}" if summary_is_brl[i] else f"={net_ref}"
    fin_formula = f"={fin_ref}/{FX}" if summary_is_brl[i] else f"={fin_ref}"
    ws.cell(row=rr, column=2, value=net_formula).number_format = USD
    fc = ws.cell(row=rr, column=3, value=fin_formula)
    fc.number_format = USD
    fc.font = BOLD
    rr += 1

ws.column_dimensions["A"].width = 62
for col in "BCDEFGHIJK":
    ws.column_dimensions[col].width = 19

from openpyxl.worksheet.datavalidation import DataValidation
dv1 = DataValidation(type="list", formula1='"0,1"', allow_blank=False)
dv2 = DataValidation(type="list", formula1='"1,2"', allow_blank=False)
ws.add_data_validation(dv1); ws.add_data_validation(dv2)
dv1.add("B5"); dv2.add("B6")
wb.save(TARGET)
print("saved", TARGET.name, "sheet", T["sheet"])
