# -*- coding: utf-8 -*-
"""Gera o Cenario B (Regime Especial Supermetal + ISS sobre servico) a partir dos
numeros reais do 'Container in a GDT' (Pathfinder GDT MRS - Cost Rollup - Rev0.xlsx)
e do 'Executive Summary (P&L)' (Pathfinder_Cost_Summary_PT.xlsx).
Cada premissa fica comentada em portugues (coluna + cell comment) para revisao do time de tax.
"""
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

OUT = Path(__file__).parent / "Pathfinder_Cost_Summary_PT_Cenario_B_RegimeEspecial.xlsx"

NAVY = "14315B"; HEADER_FILL = "1F3864"; SECTION_FILL = "D9E2F3"
WARN_FILL = "FCE4D6"; GOOD_FILL = "E2EFDA"
thin = Side(style="thin", color="B7B7B7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Cenario B - Regime Especial"
ws.sheet_view.showGridLines = False

col_widths = {"A": 3, "B": 52, "C": 16, "D": 16, "E": 78}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

r = 1
def title(text, size=16, color="FFFFFF", fill=HEADER_FILL, row=None, span="B:E"):
    global r
    rr = row or r
    ws.merge_cells(f"B{rr}:E{rr}")
    c = ws[f"B{rr}"]
    c.value = text
    c.font = Font(name="Aptos Display", size=size, bold=True, color=color)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[rr].height = 26
    if row is None:
        r += 1

def section(text):
    global r
    ws.merge_cells(f"B{r}:E{r}")
    c = ws[f"B{r}"]
    c.value = text
    c.font = Font(name="Aptos Display", size=12, bold=True, color=NAVY)
    c.fill = PatternFill("solid", fgColor=SECTION_FILL)
    for col in "BCDE":
        ws[f"{col}{r}"].border = border
    r += 1

def header_row():
    global r
    labels = ["Descricao", "R$", "US$", "Premissa / Fonte (para o time de tax)"]
    for col, lab in zip("BCDE", labels):
        cell = ws[f"{col}{r}"]
        cell.value = lab
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="41649A")
        cell.border = border
    r += 1

def line(desc, brl_formula_or_value, usd_formula, fonte, bold=False, fill=None, comment=None, brl_is_formula=True):
    global r
    ws[f"B{r}"] = desc
    ws[f"C{r}"] = brl_formula_or_value
    ws[f"D{r}"] = usd_formula
    ws[f"E{r}"] = fonte
    ws[f"C{r}"].number_format = '#,##0.00'
    ws[f"D{r}"].number_format = '#,##0.00'
    if bold:
        for col in "BCD":
            ws[f"{col}{r}"].font = Font(bold=True)
    if fill:
        for col in "BCDE":
            ws[f"{col}{r}"].fill = PatternFill("solid", fgColor=fill)
    for col in "BCDE":
        ws[f"{col}{r}"].border = border
        ws[f"{col}{r}"].alignment = Alignment(wrap_text=True, vertical="top") if col == "E" else Alignment(vertical="center")
    if comment:
        ws[f"B{r}"].comment = Comment(comment, "Analise Pathfinder GDT")
    row_used = r
    r += 1
    return row_used

# ---------------------------------------------------------------------------
title("CENARIO B - REGIME ESPECIAL SUPERMETAL + ISS SOBRE SERVICO")
r += 0
sub = ws[f"B{r}"]
ws.merge_cells(f"B{r}:E{r}")
sub.value = ("Objetivo: separar, dentro do custo 'Supermetal - GDT Integration/Assembly', o que e MATERIAL "
             "(sujeito a ICMS/IPI/PIS-COFINS) do que e SERVICO/mao de obra (sujeito a ISS), em vez de tributar "
             "o valor inteiro como se fosse uma venda de mercadoria. Todas as premissas abaixo estao comentadas "
             "para revisao do time de tax antes de qualquer decisao comercial.")
sub.font = Font(italic=True, size=10, color="444444")
sub.alignment = Alignment(wrap_text=True)
ws.row_dimensions[r].height = 46
r += 2

# ---------------------------------------------------------------------------
section("0) INPUTS EDITAVEIS (mesma metodologia do Executive Summary (P&L))")
header_row()
fx_row = line("Taxa de cambio (USD->BRL)", 5.10, None, "Copiado do Executive Summary (P&L), celula B3 do Pathfinder_Cost_Summary_PT.xlsx.")
tdm_row = line("Margem total do negocio (TDM)", 0.45, None, "Copiado do Executive Summary (P&L), celula B4. E o mesmo alvo de margem usado no cenario atual (Cenario A); nao muda entre cenarios.")
piscofins_row = line("PIS/COFINS (mantido igual ao Cenario A)", 0.1006, None,
                      "PREMISSA A CONFIRMAR COM TAX: assumimos que o PIS/COFINS incide igual sobre mercadoria e sobre servico (mesma aliquota blended 10,06% do Cenario A). Na pratica pode haver diferenca de regime (cumulativo x nao cumulativo) entre a nota fiscal de servico e a nota fiscal de mercadoria - pedir confirmacao.")
icms_regime_row = line("ICMS - Regime Especial Supermetal (materiais)", 0.015, None,
                        "ENCONTRADO NA PROPRIA COTACAO DA SUPERMETAL: aba 'Container in a GDT' do arquivo 'Pathfinder GDT MRS - Cost Rollup - Rev0.xlsx', linha 13, coluna 'Regime Especial 1,5%'. Ou seja, a propria Supermetal ja cotou essa operacao assumindo um Regime Especial de ICMS de 1,5% (em vez do 18% padrao usado hoje no Cenario A). PRECISA CONFIRMACAO: (a) se esse regime existe e esta ativo hoje junto a SEFAZ-MG, (b) se e aplicavel a essa operacao especifica (indutrializacao por encomenda) e nao apenas a cotacao antiga.")
ipi_regime_row = line("IPI - Regime Especial Supermetal (materiais)", 0.0, None,
                       "MESMA FONTE: aba 'Container in a GDT', linha 14 ('IPI' = 0). A cotacao da Supermetal ja assumia IPI = 0% para essa operacao. CONFIRMAR COM TAX se isso e por a operacao ser indutrializacao por encomenda de bem de terceiro (fora do campo de incidencia do IPI) ou se e apenas uma omissao da cotacao antiga.")
iss_row = line("ISS sobre a parcela de servico (mao de obra)", 0.05, None,
               "PLACEHOLDER A CONFIRMAR: nao encontramos a aliquota real do ISS do municipio onde a Supermetal presta o servico. Usamos 5% (teto comum para servicos de instalacao/indutrializacao no item 14.01/14.06 da LC 116/03) apenas para dimensionar a ordem de grandeza. Precisa ser substituido pela aliquota real do municipio antes de qualquer decisao.")
r += 1

# ---------------------------------------------------------------------------
section("1) DECOMPOSICAO DO CUSTO SUPERMETAL EM MATERIAL x SERVICO")
header_row()
mat_container = line("Container + Laudo Contaminantes (SM)", 70000, f"=C{r}/${'C'}${fx_row}",
                      "Fonte: 'Container in a GDT', linha 24, coluna 'Extended'.")
mat_pneum = line("Pneumatica (tubos, mangueiras, conexoes) (SM)", 25000, f"=C{r}/$C${fx_row}",
                  "Fonte: linha 25 da mesma aba.")
mat_elec_rack = line("Conexao do rack eletrico", 25000, f"=C{r}/$C${fx_row}", "Fonte: linha 26.")
mat_elec = line("Eletrica (fiacao, luzes, strobe, steplights) (SM)", 5000, f"=C{r}/$C${fx_row}", "Fonte: linha 27.")
mat_mu = line("Conectores/jumpers MU (SM)", 60000, f"=C{r}/$C${fx_row}", "Fonte: linha 28.")
mat_ac = line("Ar condicionado", 100000, f"=C{r}/$C${fx_row}", "Fonte: linha 29.")
mat_isol = line("Isolamento termico", 100000, f"=C{r}/$C${fx_row}", "Fonte: linha 30.")
mat_steel = line("Aco (reparo do vagao, porta de acesso, escada)", 12500, f"=C{r}/$C${fx_row}", "Fonte: linha 31 (25 R$/kg).")
mat_paint = line("Pintura e rotulagem", 5000, f"=C{r}/$C${fx_row}", "Fonte: linha 32. Nao inclui pintura do vagao inteiro, so rotulagem/retoque.")
mat_misc = line("Miscelaneos (materiais)", 10000, f"=C{r}/$C${fx_row}", "Fonte: linha 33.")
mat_transport = line("Transporte do container (frete)", 20000, f"=C{r}/$C${fx_row}",
                      "Fonte: linha 36 ('Trazer Container'). PREMISSA: tratado aqui junto com materiais porque frete intermunicipal/interestadual e fato gerador de ICMS-transporte, nao de ISS. Confirmar com tax se cabe destacar como servico de transporte a parte.")
mat_subtotal_row = line("SUBTOTAL MATERIAL + LOGISTICA (sujeito a ICMS/IPI)", None, None,
                         "Soma das linhas de material e transporte acima.", bold=True, fill=SECTION_FILL)
ws[f"C{mat_subtotal_row}"] = f"=SUM(C{mat_container}:C{mat_misc})+C{mat_transport}"
ws[f"D{mat_subtotal_row}"] = f"=C{mat_subtotal_row}/$C${fx_row}"

r += 1
labor_row = line("Mao de obra direta (Labor Cost)", 122752, f"=C{r}/$C${fx_row}",
                  "Fonte: 'Container in a GDT', linha 35. Calculado como 1.096 horas x R$112/h (linhas 4 e 5 da mesma aba).")
nre_row = line("NRE - engenharia de projeto", 39424, f"=C{r}/$C${fx_row}",
               "Fonte: linha 40 ('1 eng + 1 proj durante 1 mes').")
warranty_row = line("Garantia (12 meses)", 39000, f"=C{r}/$C${fx_row}",
                     "Fonte: linha 43 (3% sobre o preco). E uma provisao de servico de garantia, nao um bem.")
training_row = line("Treinamento", 2000, f"=C{r}/$C${fx_row}", "Fonte: linha 45.")
commission_row = line("Comissionamento", 44352, f"=C{r}/$C${fx_row}", "Fonte: linha 46.")
sga_row = line("SG&A alocado (5%)", 27762.6, f"=C{r}/$C${fx_row}",
               "Fonte: linha 48 (5% de overhead alocado sobre o preco). PREMISSA: tratado como parte do custo do servico porque remunera estrutura administrativa da propria prestacao de servico, nao e um insumo fisico.")
service_subtotal_row = line("SUBTOTAL SERVICO (sujeito a ISS)", None, None,
                             "Soma de mao de obra + NRE + garantia + treinamento + comissionamento + SG&A alocado.",
                             bold=True, fill=SECTION_FILL)
ws[f"C{service_subtotal_row}"] = f"=SUM(C{labor_row}:C{sga_row})"
ws[f"D{service_subtotal_row}"] = f"=C{service_subtotal_row}/$C${fx_row}"

r += 1
total_check_row = line("TOTAL (deve bater com a linha 'Supermetal' do Cenario A atual)", None, None,
                        "Confere com Pathfinder_Cost_Summary_PT.xlsx, aba 'Executive Summary (P&L)', celula B16 = R$ 707.790,60.",
                        bold=True, fill=GOOD_FILL,
                        comment="Checagem: 412.500 (material) + 20.000 (transporte) + 275.290,60 (servico) = 707.790,60 - bate exatamente com o valor hoje lancado como uma unica linha no Cenario A.")
ws[f"C{total_check_row}"] = f"=C{mat_subtotal_row}+C{service_subtotal_row}"
ws[f"D{total_check_row}"] = f"=C{total_check_row}/$C${fx_row}"
r += 1

# ---------------------------------------------------------------------------
section("2) OUTROS CUSTOS DO DEAL (nao mudam entre Cenario A e B)")
header_row()
raw_imp_row = line("Custo bruto - itens importados", 1346980.92, f"=C{r}/$C${fx_row}",
                    "Copiado do Executive Summary (P&L), celula B10. Nao muda no Cenario B.")
raw_pac_row = line("Custo bruto - itens PAC (ja nacionalizados)", 166060.16, f"=C{r}/$C${fx_row}",
                    "Copiado da celula B11. PAC ja e custo local todo-incluso; nao entra na discussao de ISS x ICMS deste cenario.")
tp_row = line("TP - Transfer Price (margem propria da WBT USA)", 198163.69, f"=C{r}/$C${fx_row}",
              "Copiado da celula B12. Segue tratado como nesta sessao: soma ao custo, mas fica fora da base de margem da WBT Brazil.")
freight_row = line("Frete internacional", 18998.55, f"=C{r}/$C${fx_row}", "Copiado da celula B13.")
ii_row = line("II - Imposto de Importacao", 280915.82, f"=C{r}/$C${fx_row}", "Copiado da celula B14.")
dom_freight_row = line("Frete domestico", 21753.10, f"=C{r}/$C${fx_row}", "Copiado da celula B15.")
r += 1

# ---------------------------------------------------------------------------
section("3) CENARIO A (ATUAL) - TUDO TRIBUTADO COMO MERCADORIA")
header_row()
final_cost_a_row = line("Custo final (com TP), incluindo Supermetal inteiro", None, None,
                         "Soma de todos os custos acima + o total Supermetal (material+servico juntos).", bold=True)
ws[f"C{final_cost_a_row}"] = (f"=C{raw_imp_row}+C{raw_pac_row}+C{tp_row}+C{freight_row}+C{ii_row}"
                               f"+C{dom_freight_row}+C{total_check_row}")
ws[f"D{final_cost_a_row}"] = f"=C{final_cost_a_row}/$C${fx_row}"

base_margin_a_row = line("Base para margem (custo final - TP)", None, None, "Custo final acima menos TP.")
ws[f"C{base_margin_a_row}"] = f"=C{final_cost_a_row}-C{tp_row}"
ws[f"D{base_margin_a_row}"] = f"=C{base_margin_a_row}/$C${fx_row}"

net_price_a_row = line("Preco liquido s/ TP (base / (1-TDM))", None, None, "Mesma formula do Cenario A hoje.")
ws[f"C{net_price_a_row}"] = f"=C{base_margin_a_row}/(1-$C${tdm_row})"
ws[f"D{net_price_a_row}"] = f"=C{net_price_a_row}/$C${fx_row}"

net_price_a_tp_row = line("Preco liquido (+ TP de volta)", None, None, "Mesma formula do Cenario A hoje.")
ws[f"C{net_price_a_tp_row}"] = f"=C{net_price_a_row}+C{tp_row}"
ws[f"D{net_price_a_tp_row}"] = f"=C{net_price_a_tp_row}/$C${fx_row}"

pc_a_row = line("+ PIS/COFINS (gross-up)", None, None, "Mesma formula do Cenario A hoje.")
ws[f"C{pc_a_row}"] = f"=C{net_price_a_tp_row}/(1-$C${piscofins_row})"
ws[f"D{pc_a_row}"] = f"=C{pc_a_row}/$C${fx_row}"

icms_a_row = line("+ ICMS 18% padrao (gross-up, sobre TUDO)", None, None,
                   "Aliquota cheia de 18% aplicada sobre o valor inteiro, inclusive a parte de servico da Supermetal - e exatamente o ponto que o Cenario B corrige.")
ws[f"C{icms_a_row}"] = f"=C{pc_a_row}/(1-0.18)"
ws[f"D{icms_a_row}"] = f"=C{icms_a_row}/$C${fx_row}"

ipi_a_row = line("+ IPI 9,69% (sobre TUDO)", None, None,
                  "Aliquota blended de IPI aplicada sobre o valor inteiro, inclusive mao de obra - tambem corrigido no Cenario B.")
ws[f"C{ipi_a_row}"] = f"=C{icms_a_row}*0.0969"
ws[f"D{ipi_a_row}"] = f"=C{ipi_a_row}/$C${fx_row}"

final_client_a_row = line("PRECO FINAL AO CLIENTE - CENARIO A", None, None,
                           "ICMS acumulado mais IPI.", bold=True, fill=WARN_FILL)
ws[f"C{final_client_a_row}"] = f"=C{icms_a_row}+C{ipi_a_row}"
ws[f"D{final_client_a_row}"] = f"=C{final_client_a_row}/$C${fx_row}"
r += 1

# ---------------------------------------------------------------------------
section("4) CENARIO B - REGIME ESPECIAL (MATERIAL) + ISS (SERVICO)")
header_row()
final_cost_b_row = line("Custo final (com TP), SEM o Supermetal (tratado a parte abaixo)", None, None,
                         "Mesma soma da secao 2, mas sem a linha Supermetal - ela ganha tratamento tributario proprio abaixo.", bold=True)
ws[f"C{final_cost_b_row}"] = f"=C{raw_imp_row}+C{raw_pac_row}+C{tp_row}+C{freight_row}+C{ii_row}+C{dom_freight_row}"
ws[f"D{final_cost_b_row}"] = f"=C{final_cost_b_row}/$C${fx_row}"

base_margin_b_row = line("Base para margem (sem Supermetal, sem TP)", None, None, "Custo acima menos TP.")
ws[f"C{base_margin_b_row}"] = f"=C{final_cost_b_row}-C{tp_row}"
ws[f"D{base_margin_b_row}"] = f"=C{base_margin_b_row}/$C${fx_row}"

net_price_b_row = line("Preco liquido s/ TP", None, None, "Mesma margem-alvo (TDM) do Cenario A.")
ws[f"C{net_price_b_row}"] = f"=C{base_margin_b_row}/(1-$C${tdm_row})"
ws[f"D{net_price_b_row}"] = f"=C{net_price_b_row}/$C${fx_row}"

net_price_b_tp_row = line("Preco liquido (+ TP de volta)", None, None, "Idem Cenario A.")
ws[f"C{net_price_b_tp_row}"] = f"=C{net_price_b_row}+C{tp_row}"
ws[f"D{net_price_b_tp_row}"] = f"=C{net_price_b_tp_row}/$C${fx_row}"

pc_b_row = line("+ PIS/COFINS (gross-up, parte importada/PAC)", None, None, "Mesma aliquota, mesma logica do Cenario A.")
ws[f"C{pc_b_row}"] = f"=C{net_price_b_tp_row}/(1-$C${piscofins_row})"
ws[f"D{pc_b_row}"] = f"=C{pc_b_row}/$C${fx_row}"

icms_b_row = line("+ ICMS 18% (parte importada/PAC, sem Supermetal)", None, None, "Mesma aliquota padrao do Cenario A, aplicada so ao resto do BOM.")
ws[f"C{icms_b_row}"] = f"=C{pc_b_row}/(1-0.18)"
ws[f"D{icms_b_row}"] = f"=C{icms_b_row}/$C${fx_row}"

ipi_b_row = line("+ IPI 9,69% (parte importada/PAC, sem Supermetal)", None, None, "Idem.")
ws[f"C{ipi_b_row}"] = f"=C{icms_b_row}*0.0969"
ws[f"D{ipi_b_row}"] = f"=C{ipi_b_row}/$C${fx_row}"

final_non_sm_row = line("Preco final ao cliente - parte NAO-Supermetal", None, None, "ICMS mais IPI acima.", bold=True)
ws[f"C{final_non_sm_row}"] = f"=C{icms_b_row}+C{ipi_b_row}"
ws[f"D{final_non_sm_row}"] = f"=C{final_non_sm_row}/$C${fx_row}"

r += 1
ws.merge_cells(f"B{r}:E{r}")
ws[f"B{r}"] = "4a) Parcela MATERIAL da Supermetal - Regime Especial (ICMS 1,5% / IPI 0%)"
ws[f"B{r}"].font = Font(bold=True, italic=True, color=NAVY)
r += 1

mat_net_row = line("Material+logistica com margem (TDM aplicada)", None, None,
                    "Mesma margem-alvo (TDM) aplicada so sobre o subtotal de material.")
ws[f"C{mat_net_row}"] = f"=C{mat_subtotal_row}/(1-$C${tdm_row})"
ws[f"D{mat_net_row}"] = f"=C{mat_net_row}/$C${fx_row}"

mat_pc_row = line("+ PIS/COFINS (mesma aliquota do Cenario A)", None, None, "Premissa explicada na celula B do input PIS/COFINS acima.")
ws[f"C{mat_pc_row}"] = f"=C{mat_net_row}/(1-$C${piscofins_row})"
ws[f"D{mat_pc_row}"] = f"=C{mat_pc_row}/$C${fx_row}"

mat_icms_row = line("+ ICMS Regime Especial Supermetal (1,5%, nao 18%)", None, None,
                     "Fonte da aliquota: input do topo (celula C do 'ICMS - Regime Especial Supermetal').",
                     fill=GOOD_FILL)
ws[f"C{mat_icms_row}"] = f"=C{mat_pc_row}/(1-$C${icms_regime_row})"
ws[f"D{mat_icms_row}"] = f"=C{mat_icms_row}/$C${fx_row}"

mat_ipi_row = line("+ IPI Regime Especial Supermetal (0%, nao 9,69%)", None, None,
                    "Fonte: input do topo. Zera o IPI para este bloco.", fill=GOOD_FILL)
ws[f"C{mat_ipi_row}"] = f"=C{mat_icms_row}*C{ipi_regime_row}"
ws[f"D{mat_ipi_row}"] = f"=C{mat_ipi_row}/$C${fx_row}"

final_mat_row = line("Preco final - parcela MATERIAL Supermetal", None, None, "ICMS regime especial mais IPI regime especial.", bold=True)
ws[f"C{final_mat_row}"] = f"=C{mat_icms_row}+C{mat_ipi_row}"
ws[f"D{final_mat_row}"] = f"=C{final_mat_row}/$C${fx_row}"

r += 1
ws.merge_cells(f"B{r}:E{r}")
ws[f"B{r}"] = "4b) Parcela SERVICO da Supermetal - ISS (sem ICMS, sem IPI)"
ws[f"B{r}"].font = Font(bold=True, italic=True, color=NAVY)
r += 1

serv_net_row = line("Servico com margem (TDM aplicada)", None, None, "Mesma margem-alvo (TDM) aplicada so sobre o subtotal de servico.")
ws[f"C{serv_net_row}"] = f"=C{service_subtotal_row}/(1-$C${tdm_row})"
ws[f"D{serv_net_row}"] = f"=C{serv_net_row}/$C${fx_row}"

serv_pc_row = line("+ PIS/COFINS (premissa a confirmar, ver input acima)", None, None, "Mesma ressalva do input PIS/COFINS do topo.")
ws[f"C{serv_pc_row}"] = f"=C{serv_net_row}/(1-$C${piscofins_row})"
ws[f"D{serv_pc_row}"] = f"=C{serv_pc_row}/$C${fx_row}"

serv_iss_row = line("+ ISS (nao ICMS, nao IPI) - aliquota do municipio da Supermetal", None, None,
                     "Fonte: input 'ISS sobre a parcela de servico' no topo - PLACEHOLDER 5%, substituir pela aliquota real.",
                     fill=WARN_FILL)
ws[f"C{serv_iss_row}"] = f"=C{serv_pc_row}*C{iss_row}"
ws[f"D{serv_iss_row}"] = f"=C{serv_iss_row}/$C${fx_row}"

final_serv_row = line("Preco final - parcela SERVICO Supermetal", None, None, "PIS/COFINS acumulado mais ISS.", bold=True)
ws[f"C{final_serv_row}"] = f"=C{serv_pc_row}+C{serv_iss_row}"
ws[f"D{final_serv_row}"] = f"=C{final_serv_row}/$C${fx_row}"

r += 1
final_client_b_row = line("PRECO FINAL AO CLIENTE - CENARIO B (soma das 3 partes)", None, None,
                           "Parte nao-Supermetal mais parte material Supermetal (regime especial) mais parte servico Supermetal (ISS).",
                           bold=True, fill=GOOD_FILL)
ws[f"C{final_client_b_row}"] = f"=C{final_non_sm_row}+C{final_mat_row}+C{final_serv_row}"
ws[f"D{final_client_b_row}"] = f"=C{final_client_b_row}/$C${fx_row}"
r += 1

# ---------------------------------------------------------------------------
section("5) COMPARACAO E ECONOMIA ESTIMADA")
header_row()
diff_row = line("Diferenca (Cenario A - Cenario B)", None, None,
                 "Quanto o Cenario B reduziria o preco final ao cliente, so nesta linha de Supermetal. Nao inclui outras frentes (ex.: II sobre o resto do BOM), que continuam iguais nos dois cenarios.",
                 bold=True, fill=GOOD_FILL)
ws[f"C{diff_row}"] = f"=C{final_client_a_row}-C{final_client_b_row}"
ws[f"D{diff_row}"] = f"=C{diff_row}/$C${fx_row}"
pct_row = line("Diferenca em % do preco final atual", None, None, "Diferenca dividida pelo preco final do Cenario A.")
ws[f"C{pct_row}"] = f"=C{diff_row}/C{final_client_a_row}"
ws[f"C{pct_row}"].number_format = "0.0%"
ws[f"D{pct_row}"] = None

r += 1
ws.merge_cells(f"B{r}:E{r}")
note = ws[f"B{r}"]
note.value = ("IMPORTANTE: os valores de ICMS Regime Especial (1,5%), IPI (0%) e ISS (5% placeholder) precisam "
              "de confirmacao formal do time de tax/juridico antes de qualquer uso comercial ou contratual. "
              "Este arquivo existe para orientar essa conversa, nao para ser usado como preco final.")
note.font = Font(bold=True, italic=True, color="9C0006")
note.alignment = Alignment(wrap_text=True)
ws.row_dimensions[r].height = 40
ws.merge_cells(f"B{r}:E{r+1}")

wb.save(OUT)
print("Saved", OUT)

# ---------------------------------------------------------------------------
# Segunda aba: hipotese estrategica CPTM
ws2 = wb.create_sheet("Nota Estrategica - CPTM")
ws2.sheet_view.showGridLines = False
for col, w in {"A": 3, "B": 100}.items():
    ws2.column_dimensions[col].width = w

rr = 1
ws2.merge_cells(f"B{rr}:B{rr}")
c = ws2[f"B{rr}"]
c.value = "HIPOTESE ESTRATEGICA: VAGAO PATHFINDER COMO ATIVO COMPARTILHADO COM A OBRIGACAO CPTM"
c.font = Font(bold=True, size=14, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=HEADER_FILL)
ws2.row_dimensions[rr].height = 30
rr += 2

paragraphs = [
    ("Contexto da obrigacao", "A concessao da MRS, renovada em julho/2022 por 30 anos (ate 2056), tem no seu Plano de "
     "Investimentos um pilar chamado 'Investimentos estruturantes de interesse publico', que inclui explicitamente a "
     "segregacao de linhas com a CPTM (Companhia Paulista de Trens Metropolitanos) em Sao Paulo. Fonte: pagina oficial "
     "'Renovacao da Concessao' do site da MRS."),
    ("Hipotese levantada por Matheus", "O vagao Pathfinder poderia embarcar, no mesmo bordo (onboard), tanto o "
     "sistema de sinalizacao/deteccao do Pathfinder quanto equipamento de sinalizacao util para o projeto de "
     "segregacao de linha com a CPTM. Se tecnicamente viavel, o mesmo ativo fisico serviria a dois propositos: "
     "(1) o produto comercial Pathfinder GDT vendido a MRS, e (2) parte do cumprimento da obrigacao contratual de "
     "investimento estruturante junto a CPTM prevista na concessao."),
    ("Por que isso pode importar para o tema fiscal", "Se uma parcela do vagao puder ser formalmente reconhecida como "
     "cumprimento de obrigacao de investimento da concessao (nao uma compra comercial comum), isso muda a natureza "
     "juridica e possivelmente contabil/fiscal dessa parcela - por exemplo, podendo abrir espaco para tratamento "
     "de investimento em infraestrutura (potencialmente elegivel a regimes como o REPORTO, sujeito a confirmacao) "
     "em vez de uma venda mercantil integral tributada normalmente."),
    ("O que NAO estamos afirmando", "Nao estamos assumindo que esse enquadramento e valido, nem estimando um valor de "
     "beneficio. Esta e uma hipotese de negocio/juridica que depende de: (a) confirmacao tecnica de que o mesmo bordo "
     "de sinalizacao atende aos dois propositos sem conflito de especificacao; (b) validacao junto a area de relacoes "
     "institucionais/juridico da MRS se isso pode contar para a obrigacao de investimento da concessao; (c) validacao "
     "do time de tax se ha algum efeito tributario associado a esse enquadramento."),
    ("Proximo passo sugerido", "Levar esta hipotese, junto com o Cenario B (aba anterior), para uma reuniao conjunta "
     "com o time de tax e com a area responsavel pelas obrigacoes de concessao/CPTM da MRS, antes de assumir qualquer "
     "numero de economia fiscal ligado a esse argumento."),
]
for heading, body in paragraphs:
    ws2[f"B{rr}"] = heading
    ws2[f"B{rr}"].font = Font(bold=True, size=11, color=NAVY)
    rr += 1
    ws2.merge_cells(f"B{rr}:B{rr}")
    cell = ws2[f"B{rr}"]
    cell.value = body
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.font = Font(size=10.5)
    # approximate row height for wrapped text
    ws2.row_dimensions[rr].height = 18 * (len(body) // 95 + 2)
    rr += 2

wb.save(OUT)
print("Saved with CPTM note", OUT)
