"""Builds MRS_Pathfinder_Simulation_Local_Taxes_Bilingual.pptx (slide 1 EN, slide 2 PT).
All numbers come from workbook tab Simulation (Pathfinder_Cost_Summary_EN.xlsx); none computed here.
Run: PYTHONIOENCODING=utf-8 python build_simulation_slide.py"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from lxml import etree

OUT = Path(__file__).parent / "MRS_Pathfinder_Simulation_Local_Taxes_Bilingual.pptx"
RED, DARK_BLUE, BLUE, INK, MUTED = "D70010", "003158", "3685C0", "2B2A2F", "757588"
PANEL, LIGHT, WHITE = "F8F8F8", "E1E0D9", "FFFFFF"
FONT, FONT_H = "Aptos", "Aptos Display"
NET = [945, 845, 908, 813]; TAX = [461, 412, 442, 396]; FINAL = [1406, 1257, 1350, 1210]
YMAX = 1600; REF = 900
# inner plot area as fractions of chart frame
PX, PY, PW, PH = 0.03, 0.10, 0.76, 0.66
CX, CY, CW, CH = 0.45, 1.25, 7.7, 4.0


def k_en(v):
    return "US$" + format(v, ",") + "k"


def k_pt(v):
    return "US$" + format(v, ",").replace(",", ".") + "k"


T = {
    "EN": dict(
        title="Even with zero import duty and one battery set, local taxes keep the Pathfinder expensive",
        cats=["1  Base\n(2 battery sets,\nimport duty in)", "2  No import\nduty (II = 0)", "3  One battery\nset (II in)", "4  No II +\none battery set"],
        s_net="Our price (net, before local taxes)", s_tax="Local taxes (PIS/COFINS + ICMS + IPI)",
        ref="US domestic reference US$900k", fmt=k_en,
        tiles=[("US$149k", "Effect of removing ALL import duty (II) on the final client price"),
               ("US$461k", "Local taxes on top of our net price in the base case = 3.1x the II effect"),
               ("+34%", "Scenario 4 final price (US$1,210k) is still above the US$900k US reference")],
        msg="Removing all import duty and using one battery set still leaves a final price of US$1,210k. What makes the Pathfinder expensive for MRS is the local Brazilian taxes, not the import duty.",
        note="II matters more than its face value because it sits inside the cost: it is multiplied by the margin (1/(1-45%)) and by the local tax cascade (~2.7x), and local taxes still outweigh it.\n"
             "Assumptions: TDM 45%; FX 5.10 BRL/USD; II 17.96% blended (only 6 of 47 imported items have a confirmed NCM, the rest at a worst-case 18%); PIS/COFINS 10.06%, ICMS 18%, IPI 9.69% passed through to MRS; Supermetal refurbishment included, no ISS; 27 items still without a quote are excluded. Source: workbook tab Simulation."),
    "PT": dict(
        title="Sem importação, o Pathfinder segue caro: são os impostos locais",
        cats=["1  Base\n(2 jogos de baterias,\nII incluído)", "2  Sem imposto de\nimportação (II = 0)", "3  Um jogo de\nbaterias (II incluído)", "4  Sem II +\num jogo de baterias"],
        s_net="Nosso preço (líquido, antes dos impostos locais)", s_tax="Impostos locais (PIS/COFINS + ICMS + IPI)",
        ref="Referência doméstica EUA US$900k", fmt=k_pt,
        tiles=[("US$149k", "Efeito de zerar TODO o imposto de importação (II) no preço final ao cliente"),
               ("US$461k", "Impostos locais sobre o nosso preço líquido no caso base = 3,1x o efeito do II"),
               ("+34%", "O preço final do cenário 4 (US$1.210k) ainda fica acima da referência EUA de US$900k")],
        msg="Zerar todo o imposto de importação e usar um só jogo de baterias ainda deixa o preço final em US$1.210k. O que encarece o Pathfinder para a MRS são os impostos locais brasileiros, não o imposto de importação.",
        note="O II pesa mais que o seu valor nominal porque está dentro do custo: é multiplicado pela margem (1/(1-45%)) e pela cascata de impostos locais (~2,7x), e mesmo assim os impostos locais pesam mais.\n"
             "Premissas: TDM 45%; câmbio 5,10 BRL/USD; II 17,96% médio (só 6 de 47 itens importados têm NCM confirmada, o restante a 18% pior caso); PIS/COFINS 10,06%, ICMS 18%, IPI 9,69% repassados à MRS; reforma da Supermetal incluída, sem ISS; 27 itens ainda sem cotação excluídos. Fonte: aba Simulação da planilha."),
}


def rgb(h):
    return RGBColor.from_string(h)


def fill(tf, paras, anchor=MSO_ANCHOR.MIDDLE, inset=(0.1, 0.05, 0.1, 0.05)):
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(x) for x in inset]
    for i, (runs, align) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for text, size, bold, color, font in runs:
            r = p.add_run()
            r.text = text
            r.font.size, r.font.bold, r.font.name = Pt(size), bold, font
            r.font.color.rgb = rgb(color)


def box(sl, x, y, w, h, fillc, paras, line=None, anchor=MSO_ANCHOR.MIDDLE, inset=(0.1, 0.05, 0.1, 0.05)):
    s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if fillc:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fillc)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    fill(s.text_frame, paras, anchor, inset)
    return s


def bar(sl, x, y, w, h, col):
    s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(col)
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def build(prs, lang):
    t = T[lang]
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bar(sl, 0.45, 0.3, 0.09, 0.8, RED)
    box(sl, 0.6, 0.3, 12.3, 0.8, None, [([(t["title"], 26, True, INK, FONT_H)], PP_ALIGN.LEFT)])
    cd = CategoryChartData()
    cd.categories = t["cats"]
    cd.add_series(t["s_net"], NET)
    cd.add_series(t["s_tax"], TAX)
    gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(CX), Inches(CY), Inches(CW), Inches(CH), cd)
    ch = gf.chart
    ch.font.name = FONT
    ch.font.size = Pt(11)
    ch.font.color.rgb = rgb(INK)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(11)
    va = ch.value_axis
    va.minimum_scale = 0
    va.maximum_scale = YMAX
    va.visible = False
    va.has_major_gridlines = False
    ch.category_axis.tick_labels.font.size = Pt(11)
    ch.category_axis.format.line.color.rgb = rgb(MUTED)
    plot = ch.plots[0]
    plot.gap_width = 55
    plot.overlap = 100
    for ser, col in zip(plot.series, (RED, BLUE)):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(col)
        dl = ser.data_labels
        dl.show_value = True
        dl.number_format = '"US$"0"k"'
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.CENTER
        dl.font.size = Pt(14)
        dl.font.bold = True
        dl.font.color.rgb = rgb(WHITE)
    pa = ch._chartSpace.chart.plotArea
    ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    old = pa.find("{%s}layout" % ns)
    if old is not None:
        pa.remove(old)
    lay = etree.fromstring(
        '<c:layout xmlns:c="%s"><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>'
        '<c:x val="%s"/><c:y val="%s"/><c:w val="%s"/><c:h val="%s"/></c:manualLayout></c:layout>' % (ns, PX, PY, PW, PH))
    pa.insert(0, lay)
    px0 = CX + PX * CW
    pw = PW * CW
    py0 = CY + PY * CH
    ph = PH * CH

    def yv(v):
        return py0 + ph * (1 - v / YMAX)

    for i, f in enumerate(FINAL):
        cx = px0 + pw * (i + 0.5) / 4
        box(sl, cx - 0.8, yv(f) - 0.5, 1.6, 0.46, None,
            [([(t["fmt"](f), 16, True, INK, FONT_H)], PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.BOTTOM, inset=(0, 0, 0, 0.02))
    ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(px0), Inches(yv(REF)), Inches(px0 + pw + 0.15), Inches(yv(REF)))
    ln.line.color.rgb = rgb(INK)
    ln.line.width = Pt(1.75)
    ln.line.dash_style = MSO_LINE.DASH
    lx = px0 + pw + 0.18
    box(sl, lx, yv(REF) - 0.4, CX + CW - lx, 0.8, PANEL,
        [([(t["ref"], 11, True, INK, FONT)], PP_ALIGN.LEFT)], line=LIGHT, inset=(0.06, 0.03, 0.04, 0.03))
    tx, tw, th, gap = 8.45, 4.43, 1.18, 0.13
    for i, (big, small) in enumerate(t["tiles"]):
        y = 1.25 + i * (th + gap)
        bar(sl, tx, y, 0.08, th, RED if i == 1 else BLUE)
        box(sl, tx + 0.08, y, tw - 0.08, th, PANEL,
            [([(big, 30, True, INK, FONT_H)], PP_ALIGN.LEFT), ([(small, 12, False, INK, FONT)], PP_ALIGN.LEFT)],
            inset=(0.18, 0.04, 0.12, 0.04))
    box(sl, 0.45, 5.4, 12.43, 0.85, DARK_BLUE, [([(t["msg"], 16, True, WHITE, FONT)], PP_ALIGN.LEFT)], inset=(0.25, 0.05, 0.25, 0.05))
    a, b = t["note"].split("\n")
    box(sl, 0.45, 6.32, 12.43, 1.0, None,
        [([(a, 10, False, INK, FONT)], PP_ALIGN.LEFT), ([(b, 10, False, MUTED, FONT)], PP_ALIGN.LEFT)],
        anchor=MSO_ANCHOR.TOP, inset=(0, 0, 0, 0))
    return sl


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
build(prs, "EN")
build(prs, "PT")
prs.save(OUT)
print("saved", OUT)
