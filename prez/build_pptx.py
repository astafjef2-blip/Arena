# -*- coding: utf-8 -*-
"""
Презентация по докладу руководителя Управления Росреестра по Воронежской области
Е.П. Перегудовой (Форум по недвижимости в Черноземье, 2026).
Стиль: корпоративный шаблон Росреестра (Rosreestr corporate template).
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn
import copy

A = "assets/"

# ----------------------------------------------------------------------------
# Palette
# ----------------------------------------------------------------------------
PRIMARY   = RGBColor(0x00, 0x8C, 0xFF)   # primary blue
SECONDARY = RGBColor(0x18, 0x90, 0xFA)   # secondary blue
TEAL      = RGBColor(0x00, 0xCC, 0x99)   # accent teal
ORANGE    = RGBColor(0xFF, 0x79, 0x00)   # accent orange
LIGHTBLUE = RGBColor(0x99, 0xD1, 0xFF)   # light blue
MINT      = RGBColor(0xE0, 0xF9, 0xEA)   # mint fill
PALEBLUE  = RGBColor(0xE7, 0xF1, 0xFE)   # pale blue card
LIGHTGREY = RGBColor(0xF2, 0xF2, 0xF2)   # light grey
TXTGREY   = RGBColor(0x59, 0x59, 0x59)   # text grey
TXTGREY2  = RGBColor(0x4D, 0x4D, 0x4D)   # text grey dark
DARKSLATE = RGBColor(0x39, 0x50, 0x5E)   # dark slate
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)

GRAD_A = "008CFF"
GRAD_B = "00CC99"
GRAD_A2 = "1890FA"
GRAD_B2 = "47FFD1"

SW, SH = 13.333, 7.5  # slide size, inches

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def slide():
    return prs.slides.add_slide(BLANK)

def _noshadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass
    return shape

def solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    return shape

def noline(shape):
    shape.line.fill.background()
    return shape

def line(shape, color, w=1.0, dash=None):
    shape.line.color.rgb = color
    shape.line.width = Pt(w)
    if dash is not None:
        shape.line.dash_style = dash
    return shape

def gradient(shape, c1, c2, angle=135):
    shape.fill.gradient()
    stops = shape.fill.gradient_stops
    # ensure exactly 2 stops
    while len(stops) > 2:
        stops[-1]._stop.getparent().remove(stops[-1]._stop)
        stops = shape.fill.gradient_stops
    stops[0].color.rgb = c1
    stops[0].position = 0.0
    stops[1].color.rgb = c2
    stops[1].position = 1.0
    gf = shape.fill._xPr.find(qn('a:gradFill'))
    if gf is not None:
        gf.set('rot', str(int(angle * 60000)))
        gf.set('flip', 'none')
    shape.line.fill.background()
    return shape

def glow(shape, color, alpha=0.45, radius_pt=10):
    spPr = shape._element.spPr
    eff = spPr.find(qn('a:effectLst'))
    if eff is None:
        eff = spPr.makeelement(qn('a:effectLst'), {})
        # insert after ln if present, else append
        ln = spPr.find(qn('a:ln'))
        if ln is not None:
            ln.addnext(eff)
        else:
            spPr.append(eff)
    g = eff.makeelement(qn('a:glow'), {'rad': str(int(radius_pt * 12700))})
    clr = g.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % (color[0], color[1], color[2])})
    a = clr.makeelement(qn('a:alpha'), {'val': str(int(alpha * 100000))})
    clr.append(a)
    g.append(clr)
    eff.append(g)
    return shape

def rect(s, x, y, w, h, fill=None, lcolor=None, lw=1.0, shape_type=MSO_SHAPE.RECTANGLE,
         radius=None, dash=None, ang=None, g1=None, g2=None):
    sh = s.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    _noshadow(sh)
    if radius is not None:
        try:
            short = min(w, h)
            sh.adjustments[0] = max(0.0, min(0.5, radius / short))
        except Exception:
            pass
    if g1 is not None:
        gradient(sh, g1, g2, ang if ang is not None else 135)
    elif fill is not None:
        solid(sh, fill)
    else:
        sh.fill.background()
    if lcolor is not None:
        line(sh, lcolor, lw, dash)
    else:
        noline(sh)
    sh.text_frame.word_wrap = True
    return sh

def rrect(s, x, y, w, h, fill=None, lcolor=None, lw=1.25, radius=0.3, dash=None):
    return rect(s, x, y, w, h, fill=fill, lcolor=lcolor, lw=lw,
                shape_type=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius, dash=dash)

def pill(s, x, y, w, h, fill=None, lcolor=None, lw=1.0, g1=None, g2=None, ang=135, radius=None):
    return rect(s, x, y, w, h, fill=fill, lcolor=lcolor, lw=lw,
                shape_type=MSO_SHAPE.ROUNDED_RECTANGLE,
                radius=(radius if radius is not None else (h / 2.0)),
                g1=g1, g2=g2, ang=ang)

def _set_run(r, t, size, bold, color, italic=False):
    r.text = t
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = 'Arial'
    f.color.rgb = color

def text(s, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         wrap=True, ml=0.0, mr=0.0, mt=0.0, mb=0.0):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(ml)
    tf.margin_right = Inches(mr)
    tf.margin_top = Inches(mt)
    tf.margin_bottom = Inches(mb)
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = para.get('align', align)
        if 'sa' in para:
            p.space_after = Pt(para['sa'])
        if 'sb' in para:
            p.space_before = Pt(para['sb'])
        if 'ls' in para:
            p.line_spacing = para['ls']
        for rs in para['runs']:
            r = p.add_run()
            _set_run(r, rs['t'], rs.get('s', 14), rs.get('b', False),
                     rs.get('c', TXTGREY), rs.get('i', False))
    return tb

def eagle(s, x, y, size, on_blue=True):
    """Emblem: solid blue circle + white double-headed eagle."""
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(size), Inches(size))
    _noshadow(circ)
    solid(circ, PRIMARY)
    glow(circ, PRIMARY, alpha=0.35, radius_pt=8)
    img = A + "eagle_white.png"
    pad = size * 0.20
    s.shapes.add_picture(img, Inches(x + pad), Inches(y + pad),
                         height=Inches(size - 2 * pad))
    return circ

def check_bullet(s, x, y, d, fill, size=14, color=WHITE):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    _noshadow(c)
    solid(c, fill)
    tf = c.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    _set_run(r, "\u2713", size, True, color)
    return c

def page_no(s, n, total=11):
    text(s, SW - 1.5, SH - 0.5, 1.0, 0.4,
         [{'runs': [{'t': f"{n} / {total}", 's': 11, 'b': False, 'c': TXTGREY, 'i': True}],
           'align': PP_ALIGN.RIGHT}],
         anchor=MSO_ANCHOR.MIDDLE)

def bottom_strip(s):
    """Narrow tech-ornament strip along the bottom."""
    y = SH - 0.42
    rect(s, 0, y, SW, 0.42, fill=LIGHTGREY)
    # circuit line
    rect(s, 0.0, y + 0.16, SW, 0.02, fill=LIGHTBLUE)
    # round nodes
    for nx in [1.1, 3.4, 6.2, 9.0, 11.6]:
        rect(s, nx, y + 0.11, 0.10, 0.10, fill=PRIMARY, shape_type=MSO_SHAPE.OVAL)
    # diagonal hatch group
    for k in range(6):
        rect(s, SW - 2.7 + k * 0.30, y + 0.04, 0.02, 0.30, fill=SECONDARY)
    # binary digits
    text(s, 0.3, y + 0.02, 2.6, 0.38,
         [{'runs': [{'t': "01001011  11010010  00101101", 's': 9, 'b': True, 'c': LIGHTBLUE}],
           'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.MIDDLE)

def header(s, title, n):
    eagle(s, 0.55, 0.34, 0.58, on_blue=True)
    text(s, 1.32, 0.28, SW - 2.6, 0.74,
         [{'runs': [{'t': title, 's': 22, 'b': True, 'c': PRIMARY}],
           'align': PP_ALIGN.LEFT, 'ls': 0.95}],
         anchor=MSO_ANCHOR.MIDDLE)
    # thin grey rule
    rect(s, 0.55, 1.10, SW - 1.1, 0.02, fill=RGBColor(0xD9, 0xD9, 0xD9))
    bottom_strip(s)
    page_no(s, n)

def tab_label(s, x, y, w, label, fill=PRIMARY, color=WHITE, h=0.42, size=12):
    t = rrect(s, x, y, w, h, fill=fill, radius=(h / 2.0))
    tf = t.text_frame
    tf.margin_left = Inches(0.08); tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.0); tf.margin_bottom = Inches(0.0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _set_run(p.add_run(), label, size, True, color)
    return t

def kpi_card(s, x, y, w, h, value, label, sub=None, accent=TEAL):
    # soft blue outer glow behind the card
    g = rrect(s, x, y, w, h, fill=None, lcolor=None, radius=0.26)
    glow(g, PRIMARY, alpha=0.18, radius_pt=10)
    card = rrect(s, x, y, w, h, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.22); tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.18); tf.margin_bottom = Inches(0.14)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    _set_run(p.add_run(), label, 11, True, TXTGREY2)
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT; p2.space_before = Pt(4)
    _set_run(p2.add_run(), value, 30, True, PRIMARY)
    if sub:
        p3 = tf.add_paragraph(); p3.alignment = PP_ALIGN.LEFT; p3.space_before = Pt(3)
        _set_run(p3.add_run(), sub, 11.5, False, accent)
    return card

# ============================================================================
# SLIDE 1 — COVER
# ============================================================================
s = slide()
# three horizontal bands
band_top_h = 2.08   # ~28%
band_bot_h = 1.88   # ~25%
# top band (gradient blue)
tb = rect(s, 0, 0, SW, band_top_h, g1=PRIMARY, g2=SECONDARY, ang=90)
# bottom band (gradient blue)
bb = rect(s, 0, SH - band_bot_h, SW, band_bot_h, g1=SECONDARY, g2=PRIMARY, ang=90)
# large gradient circle crossing the bands (right)
circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(SW - 4.1), Inches(SH/2 - 2.3),
                          Inches(6.6), Inches(6.6))
_noshadow(circ)
gradient(circ, PRIMARY, TEAL, angle=135)
glow(circ, TEAL, alpha=0.35, radius_pt=16)
# inner lighter circle for depth
circ2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(SW - 3.4), Inches(SH/2 - 1.9),
                           Inches(4.9), Inches(4.9))
_noshadow(circ2)
gradient(circ2, SECONDARY, RGBColor(0x47, 0xFF, 0xD1), angle=135)

# emblem top-left on blue band
eagle(s, 0.62, 0.55, 1.05, on_blue=True)
text(s, 1.95, 0.62, 7.5, 0.9,
     [{'runs': [{'t': "УПРАВЛЕНИЕ РОСРЕЕСТРА", 's': 15, 'b': True, 'c': WHITE}],
       'align': PP_ALIGN.LEFT, 'sa': 2},
      {'runs': [{'t': "ВОРОНЕЖСКАЯ ОБЛАСТЬ", 's': 12, 'b': False, 'c': LIGHTBLUE}],
       'align': PP_ALIGN.LEFT}],
     anchor=MSO_ANCHOR.MIDDLE)

# title in white band
text(s, 0.9, 2.42, 7.9, 1.9,
     [{'runs': [{'t': "Доклад руководителя", 's': 34, 'b': True, 'c': PRIMARY}],
       'align': PP_ALIGN.LEFT, 'sa': 2, 'ls': 0.98},
      {'runs': [{'t': "Управления Росреестра", 's': 34, 'b': True, 'c': PRIMARY}],
       'align': PP_ALIGN.LEFT, 'sa': 2, 'ls': 0.98},
      {'runs': [{'t': "по Воронежской области", 's': 34, 'b': True, 'c': DARKSLATE}],
       'align': PP_ALIGN.LEFT, 'ls': 0.98}],
     anchor=MSO_ANCHOR.TOP)

# tagline
text(s, 0.9, 4.42, 7.9, 0.5,
     [{'runs': [{'t': "Недвижимость — основа благосостояния, стабильности и развития территорий",
                 's': 13.5, 'b': False, 'i': True, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# subtitle / event line
text(s, 0.9, 4.95, 7.9, 0.5,
     [{'runs': [{'t': "Е.П. Перегудова  ·  пленарное заседание Форума по недвижимости в Черноземье",
                 's': 14, 'b': False, 'c': SECONDARY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# bottom band content
text(s, 0.95, SH - band_bot_h + 0.45, 8.5, 0.8,
     [{'runs': [{'t': "ФОРУМ ПО НЕДВИЖИМОСТИ В ЧЕРНОЗЕМЬЕ", 's': 13, 'b': True, 'c': WHITE}],
       'align': PP_ALIGN.LEFT, 'sa': 2},
      {'runs': [{'t': "2026 год", 's': 12, 'b': False, 'c': LIGHTBLUE}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)
text(s, SW - 3.2, SH - band_bot_h + 0.45, 2.6, 0.9,
     [{'runs': [{'t': "2026", 's': 30, 'b': True, 'c': WHITE}],
       'align': PP_ALIGN.RIGHT}], anchor=MSO_ANCHOR.MIDDLE)

# ============================================================================
# SLIDE 2 — КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ
# ============================================================================
s = slide()
header(s, "Ключевые показатели работы за 9 месяцев 2026 года", 2)

# intro
text(s, 0.6, 1.30, 12.1, 0.4,
     [{'runs': [{'t': "Общий объём поступивших заявлений демонстрирует значительный рост по сравнению с 2025 годом",
                 's': 13, 'b': False, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# three KPI cards
cw, ch, gap = 3.94, 1.45, 0.20
cx0, cy = 0.6, 1.85
kpi_card(s, cx0, cy, cw, ch, "635 508", "ВСЕГО ЗАЯВЛЕНИЙ",
         "+53,2%  (+220 664)", accent=TEAL)
kpi_card(s, cx0 + (cw + gap), cy, cw, ch, "293 208", "УЧЕТНО-РЕГИСТРАЦИОННЫЕ",
         "+1,3%  (КУ, РП, ЕП)", accent=TEAL)
kpi_card(s, cx0 + 2 * (cw + gap), cy, cw, ch, "75,3%", "ДОЛЯ ЭЛЕКТРОННЫХ (РП)",
         "с 65,9%", accent=PRIMARY)

# table
rows_data = [
    ("Показатель", "9 мес. 2025", "9 мес. 2026", "Изменение", "Темп, %"),
    ("Кадастровый учет", "55 333", "42 825", "\u221212 508", "\u221222,6"),
    ("Регистрация прав", "217 773", "234 158", "+16 385", "+7,5"),
    ("Единая процедура", "16 473", "16 225", "\u2212248", "\u22121,5"),
    ("ИТОГО (КУ, РП, ЕП)", "289 579", "293 208", "+3 629", "+1,3"),
    ("Всего, включая иные", "414 844", "635 508", "+220 664", "+53,2"),
]
tbl_x, tbl_y, tbl_w = 0.6, 3.55, 12.13
tbl_h = 2.95
gt = s.shapes.add_table(len(rows_data), 5, Inches(tbl_x), Inches(tbl_y),
                        Inches(tbl_w), Inches(tbl_h)).table
# column widths
gt.columns[0].width = Inches(4.4)
gt.columns[1].width = Inches(2.1)
gt.columns[2].width = Inches(2.1)
gt.columns[3].width = Inches(1.9)
gt.columns[4].width = Inches(1.63)
for r, row in enumerate(rows_data):
    gt.rows[r].height = Inches(0.42)
    for c, val in enumerate(row):
        cell = gt.cell(r, c)
        cell.margin_left = Inches(0.12); cell.margin_right = Inches(0.10)
        cell.margin_top = Inches(0.04); cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
        run = p.add_run()
        if r == 0:
            _set_run(run, val, 12, True, WHITE)
            cell.fill.solid(); cell.fill.fore_color.rgb = PRIMARY
        elif r == len(rows_data) - 1 or r == len(rows_data) - 2:
            _set_run(run, val, 12, True, DARKSLATE if c == 0 else PRIMARY)
            cell.fill.solid()
            cell.fill.fore_color.rgb = MINT if r == len(rows_data) - 1 else PALEBLUE
        else:
            _set_run(run, val, 11.5, (c == 0), TXTGREY2)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r % 2 == 1 else LIGHTGREY
# remove table default banding style -> set first row header off
tblPr = gt._tbl.find(qn('a:tblPr'))
if tblPr is not None:
    tblPr.set('firstRow', '0'); tblPr.set('bandRow', '0')

# footnote
text(s, 0.6, 6.68, 12.1, 0.3,
     [{'runs': [{'t': "Структурные изменения: кадастровый учет \u221222,6%, регистрация прав +7,5%, единая процедура \u22121,5%.",
                 's': 10.5, 'b': False, 'i': True, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# ============================================================================
# SLIDE 3 — ИПОТЕКА
# ============================================================================
s = slide()
header(s, "Ипотечное кредитование: динамика роста", 3)

text(s, 0.6, 1.28, 12.1, 0.4,
     [{'runs': [{'t': "Рынок ипотечного кредитования региона демонстрирует количественный рост и беспрецедентный сдвиг в сторону цифровизации",
                 's': 13, 'b': False, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# 2x2 cards
cards = [
    ("34 424", "ЗАЯВЛЕНИЙ НА РЕГИСТРАЦИЮ ИПОТЕКИ", "+19,7%  (28 763)", TEAL),
    ("37 705", "ЗАРЕГИСТРИРОВАННЫХ ОБРЕМЕНЕНИЙ", "+27,8%  (29 495)", TEAL),
    ("13 956", "ПО ДОГОВОРУ  (2025: 7 522)", "+85,5%", TEAL),
    ("20 468", "ПО ЗАКОНУ  (2025: 21 241)", "\u22123,6%", ORANGE),
]
cw, ch, gap = 5.98, 1.5, 0.17
cx0, cy = 0.6, 1.85
for i, (val, lab, sub, col) in enumerate(cards):
    r, c = divmod(i, 2)
    x = cx0 + c * (cw + gap)
    y = cy + r * (ch + gap)
    rrect(s, x, y, cw, ch, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
    g = rrect(s, x, y, cw, ch, fill=None, radius=0.26); glow(g, PRIMARY, alpha=0.15, radius_pt=10)
    rrect(s, x, y, cw, ch, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
    # left accent bar
    rect(s, x, y + 0.22, 0.07, ch - 0.44, fill=col)
    text(s, x + 0.28, y + 0.16, cw - 0.5, ch - 0.3,
         [{'runs': [{'t': lab, 's': 11, 'b': True, 'c': TXTGREY2}], 'sa': 3},
          {'runs': [{'t': val, 's': 30, 'b': True, 'c': PRIMARY}], 'sa': 3},
          {'runs': [{'t': sub, 's': 12, 'b': True, 'c': col}]}],
         anchor=MSO_ANCHOR.MIDDLE)

# electronic highlight strip
hy = cy + 2 * (ch + gap) + 0.12
rrect(s, 0.6, hy, 12.13, 1.15, fill=MINT, lcolor=TEAL, lw=1.5, radius=0.26)
# gradient chevron
chv = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(0.85), Inches(hy + 0.30),
                         Inches(0.7), Inches(0.55))
_noshadow(chv); gradient(chv, PRIMARY, TEAL, 135)
text(s, 1.75, hy + 0.12, 10.7, 0.9,
     [{'runs': [{'t': "Цифровизация: ", 's': 13, 'b': True, 'c': DARKSLATE},
                {'t': "электронные заявления +51,2% (28 140 шт.)", 's': 13, 'b': False, 'c': TXTGREY2}],
       'sa': 3, 'ls': 1.0},
      {'runs': [{'t': "Доля электронных документов выросла с 64,7% до ", 's': 13, 'b': False, 'c': TXTGREY2},
                {'t': "81,7%", 's': 15, 'b': True, 'c': PRIMARY},
                {'t': "  \u2192 практически полный переход на дистанционное взаимодействие", 's': 13, 'b': False, 'c': TXTGREY2}],
       'ls': 1.0}], anchor=MSO_ANCHOR.MIDDLE)

# ============================================================================
# SLIDE 4 — СРОКИ ГОСУСЛУГ
# ============================================================================
s = slide()
header(s, "Сроки оказания государственных услуг", 4)

text(s, 0.6, 1.28, 12.1, 0.4,
     [{'runs': [{'t': "Показатели полностью укладываются в требования Федерального проекта «Национальная система пространственных данных»",
                 's': 13, 'b': False, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# service cards with pill progress bars
svcs = [
    ("Регистрация прав", "1,6", "2,0", TEAL),
    ("Кадастровый учет", "1,3", "2,0", TEAL),
    ("Единая процедура", "2,9", "3,0", TEAL),
    ("Электронные обращения", "1,0", "1,0", PRIMARY),
]
cw, ch, gap = 2.95, 2.0, 0.17
cx0, cy = 0.6, 1.95
for i, (name, fact, norm, col) in enumerate(svcs):
    x = cx0 + i * (cw + gap)
    rrect(s, x, cy, cw, ch, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
    g = rrect(s, x, cy, cw, ch, fill=None, radius=0.26); glow(g, PRIMARY, alpha=0.13, radius_pt=9)
    rrect(s, x, cy, cw, ch, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
    # tab label overlapping top edge
    tab_label(s, x + 0.55, cy - 0.20, cw - 1.10, name.upper(), fill=PRIMARY, h=0.42, size=11)
    # fact number
    text(s, x + 0.2, cy + 0.45, cw - 0.4, 0.8,
         [{'runs': [{'t': fact, 's': 34, 'b': True, 'c': PRIMARY}],
           'align': PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.2, cy + 1.18, cw - 0.4, 0.3,
         [{'runs': [{'t': "дней \u2014 факт", 's': 11, 'b': False, 'c': TXTGREY}],
           'align': PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.TOP)
    # progress pill
    py = cy + 1.52
    pill(s, x + 0.4, py, cw - 0.8, 0.26, fill=LIGHTGREY, radius=0.13)
    _f = float(fact.replace(",", "."))
    _n = float(norm.replace(",", "."))
    frac = min(1.0, _f / _n) if _n else 1.0
    pill(s, x + 0.4, py, (cw - 0.8) * frac, 0.26, g1=SECONDARY, g2=TEAL, ang=135, radius=0.13)
    text(s, x + 0.4, py + 0.30, cw - 0.8, 0.24,
         [{'runs': [{'t': f"норма \u2264 {norm} дней", 's': 10, 'b': False, 'i': True, 'c': TXTGREY}],
           'align': PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.TOP)

# improvement strip with teal down arrows
iy = 4.45
rrect(s, 0.6, iy, 12.13, 1.55, fill=PALEBLUE, lcolor=PRIMARY, lw=1.25, radius=0.26)
# teal downward arrows
for i, (txt) in enumerate(["Единая процедура: 3,7 \u2192 2,9 дня",
                           "Норма снижена: 4,0 \u2192 3,0 дня",
                           "Электронные \u2014 в срок \u2264 1 дня"]):
    ax = 1.0 + i * 3.95
    ar = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(ax), Inches(iy + 0.30),
                            Inches(0.42), Inches(0.95))
    _noshadow(ar); solid(ar, TEAL)
    text(s, ax + 0.55, iy + 0.42, 3.0, 0.7,
         [{'runs': [{'t': txt, 's': 12.5, 'b': True, 'c': DARKSLATE}], 'ls': 1.0}],
         anchor=MSO_ANCHOR.MIDDLE)

# ============================================================================
# SLIDE 5 — «ПОСТРОИЛ — ОФОРМИ» / 487-ФЗ
# ============================================================================
s = slide()
header(s, "Принцип «построил \u2014 оформил»: Закон № 487-ФЗ", 5)

text(s, 0.6, 1.26, 12.1, 0.4,
     [{'runs': [{'t': "С 1 марта 2025 года вступил в силу Закон № 487-ФЗ: эксплуатировать построенные объекты можно только после надлежащего оформления",
                 's': 13, 'b': False, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# left badge column
bx, by, bw, bh = 0.6, 1.85, 3.5, 4.95
rrect(s, bx, by, bw, bh, fill=PRIMARY, radius=0.28)
eagle(s, bx + (bw - 1.3) / 2, by + 0.5, 1.3, on_blue=False)
text(s, bx + 0.3, by + 2.0, bw - 0.6, 1.6,
     [{'runs': [{'t': "Закон", 's': 18, 'b': True, 'c': WHITE}],
       'align': PP_ALIGN.CENTER, 'sa': 2},
      {'runs': [{'t': "№ 487-ФЗ", 's': 30, 'b': True, 'c': WHITE}],
       'align': PP_ALIGN.CENTER, 'sa': 6},
      {'runs': [{'t': "вступил в силу 1 марта 2025 года", 's': 12, 'b': False, 'c': LIGHTBLUE}],
       'align': PP_ALIGN.CENTER, 'ls': 1.0}], anchor=MSO_ANCHOR.MIDDLE)

# right checklist
items = [
    ("Статья 52 ГК РФ", "момент завершения строительства \u2014 день получения разрешения на ввод в эксплуатацию (ИЖС, садовые дома)"),
    ("Без межевания \u2014 без распоряжения", "сделки в отношении участка приостановлены, если границы отсутствуют в ЕГРН (кроме сервитута)"),
    ("Электронное взаимодействие юрлиц", "учёт и регистрация прав юрлиц исключительно в электронном виде"),
    ("Долевое строительство", "заявление о регистрации права дольщика \u2014 в течение 30 рабочих дней, только электронный вид"),
    ("Земли под объектами", "собственники зданий и сооружений обязаны оформить права на земельные участки"),
]
ix, iy, iw = 4.45, 1.85, 8.28
iycur = iy
for i, (head, body) in enumerate(items):
    ih = 0.92
    rrect(s, ix, iycur, iw, ih, fill=WHITE, lcolor=PRIMARY, lw=1.0, radius=0.22)
    g = rrect(s, ix, iycur, iw, ih, fill=None, radius=0.22); glow(g, PRIMARY, alpha=0.10, radius_pt=8)
    rrect(s, ix, iycur, iw, ih, fill=WHITE, lcolor=PRIMARY, lw=1.0, radius=0.22)
    check_bullet(s, ix + 0.22, iycur + 0.24, 0.42, TEAL, size=16)
    text(s, ix + 0.82, iycur + 0.12, iw - 1.0, ih - 0.2,
         [{'runs': [{'t': head, 's': 13, 'b': True, 'c': PRIMARY}], 'sa': 2, 'ls': 0.95},
          {'runs': [{'t': body, 's': 11, 'b': False, 'c': TXTGREY2}], 'ls': 0.95}],
         anchor=MSO_ANCHOR.MIDDLE)
    iycur += ih + 0.10

# ============================================================================
# SLIDE 6 — ФГИС ЕЦП НСПД
# ============================================================================
s = slide()
header(s, "Внедрение ФГИС ЕЦП «Национальная система пространственных данных»", 6)

text(s, 0.6, 1.26, 12.1, 0.4,
     [{'runs': [{'t': "Воронежская область — признанный лидер внедрения Единой цифровой платформы НСПД, один из 31 пилотного субъекта (октябрь 2022)",
                 's': 13, 'b': False, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# KPI band
kpi = [("100%", "СЛОЁВ ПЕРЕДАНО В НСПД", TEAL),
       ("32 + 5", "СЛОЁВ (СОГД + ГЕОИС)", PRIMARY),
       ("2", "ИНФОРМАЦИОННЫЕ СИСТЕМЫ", PRIMARY),
       ("2-й", "ПОТОК ВНЕДРЕНИЕВ 2026", TEAL)]
cw, ch, gap = 2.95, 1.35, 0.13
cx0, cy = 0.6, 1.85
for i, (val, lab, col) in enumerate(kpi):
    x = cx0 + i * (cw + gap)
    rrect(s, x, cy, cw, ch, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
    g = rrect(s, x, cy, cw, ch, fill=None, radius=0.26); glow(g, PRIMARY, alpha=0.14, radius_pt=9)
    rrect(s, x, cy, cw, ch, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
    text(s, x + 0.15, cy + 0.14, cw - 0.3, ch - 0.28,
         [{'runs': [{'t': lab, 's': 10.5, 'b': True, 'c': TXTGREY2}], 'align': PP_ALIGN.CENTER, 'sa': 4, 'ls': 0.9},
          {'runs': [{'t': val, 's': 27, 'b': True, 'c': col}], 'align': PP_ALIGN.CENTER}],
         anchor=MSO_ANCHOR.MIDDLE)

# left: structure list
lx, ly, lw = 0.6, 3.45, 6.05
rrect(s, lx, ly, lw, 3.35, fill=PALEBLUE, lcolor=PRIMARY, lw=1.25, radius=0.26)
text(s, lx + 0.28, ly + 0.18, lw - 0.56, 0.4,
     [{'runs': [{'t': "ОРГАНИЗАЦИЯ ВНЕДРЕНИЕВ", 's': 13, 'b': True, 'c': PRIMARY}]}],
     anchor=MSO_ANCHOR.TOP)
org = [
    "Создан оперативный штаб под руководством Губернатора",
    "Представители области включены в межведомственную рабочую группу Росреестра",
    "Выполнены обязательства по загрузке пространственных данных (100% слоёв)",
    "Две региональные ИС зарегистрированы как поставщики информации",
]
oy = ly + 0.72
for o in org:
    check_bullet(s, lx + 0.30, oy + 0.02, 0.34, PRIMARY, size=13)
    text(s, lx + 0.78, oy, lw - 1.1, 0.6,
         [{'runs': [{'t': o, 's': 12, 'b': False, 'c': TXTGREY2}], 'ls': 0.98}],
         anchor=MSO_ANCHOR.TOP)
    oy += 0.62

# right: 4 adapted services
rx, ry, rw = 6.85, 3.45, 5.88
rrect(s, rx, ry, rw, 3.35, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
g = rrect(s, rx, ry, rw, 3.35, fill=None, radius=0.26); glow(g, PRIMARY, alpha=0.12, radius_pt=9)
rrect(s, rx, ry, rw, 3.35, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.26)
text(s, rx + 0.28, ry + 0.18, rw - 0.56, 0.4,
     [{'runs': [{'t': "4 УСЛУГИ В РЕГИОНАЛЬНОМ ПЛАНЕ (2026)", 's': 13, 'b': True, 'c': PRIMARY}]}],
     anchor=MSO_ANCHOR.TOP)
svcs6 = [
    "Уведомления о соответствии параметров и вводе в эксплуатацию (ИЖС, садовые дома)",
    "Выдача разрешений на использование земель без предоставления участков",
    "Установление публичных сервитутов",
    "Предварительное согласование предоставления земельных участков",
]
oy = ry + 0.72
for i, sv in enumerate(svcs6):
    check_bullet(s, rx + 0.30, oy + 0.02, 0.34, TEAL, size=13)
    text(s, rx + 0.78, oy, rw - 1.1, 0.6,
         [{'runs': [{'t': sv, 's': 12, 'b': False, 'c': TXTGREY2}], 'ls': 0.98}],
         anchor=MSO_ANCHOR.TOP)
    oy += 0.62

# ============================================================================
# SLIDE 7 — СЕРВИС «АНАЛИТИКА РЫНКА НЕДВИЖИМОСТИ» (IMAGE)
# ============================================================================
s = slide()
header(s, "Сервис «Аналитика рынка недвижимости» ЕЦП НСПД", 7)

text(s, 0.6, 1.24, 6.4, 0.56,
     [{'runs': [{'t': "Специализированный инструмент расчёта и визуализации индексных показателей цен, разработанный Росреестром и ППК «Роскадастр»",
                 's': 12.5, 'b': False, 'c': TXTGREY}], 'ls': 1.0}], anchor=MSO_ANCHOR.TOP)

# left content bullets
left_items = [
    ("Аудитория", "риелторы, оценщики, кредитные организации, органы власти"),
    ("Фундамент данных", "обезличенные сведения о реальных сделках из ЕГРН (с 2018), Фонд данных госкадастровой оценки, открытые рыночные данные"),
    ("Индикаторы", "ежемесячные, ежеквартальные, полугодовые и годовые индикаторы цен"),
    ("Представление", "таблицы, графики, тепловые карты; фильтры по площади и границам (до кадастрового квартала)"),
    ("Доступ", "авторизация через портал «Госуслуги»"),
]
iy = 1.80
for head, body in left_items:
    ih = 0.76
    rrect(s, 0.6, iy, 6.4, ih, fill=WHITE, lcolor=PRIMARY, lw=1.0, radius=0.22)
    g = rrect(s, 0.6, iy, 6.4, ih, fill=None, radius=0.22); glow(g, PRIMARY, alpha=0.10, radius_pt=8)
    rrect(s, 0.6, iy, 6.4, ih, fill=WHITE, lcolor=PRIMARY, lw=1.0, radius=0.22)
    check_bullet(s, 0.82, iy + 0.18, 0.40, TEAL, size=15)
    text(s, 1.36, iy + 0.08, 5.5, ih - 0.16,
         [{'runs': [{'t': head, 's': 12.5, 'b': True, 'c': PRIMARY}], 'sa': 2, 'ls': 0.95},
          {'runs': [{'t': body, 's': 10.5, 'b': False, 'c': TXTGREY2}], 'ls': 0.95}],
         anchor=MSO_ANCHOR.MIDDLE)
    iy += ih + 0.08

# highlight card bottom-left
rrect(s, 0.6, 6.06, 6.4, 0.84, fill=MINT, lcolor=TEAL, lw=1.5, radius=0.26)
text(s, 0.85, 6.11, 6.0, 0.74,
     [{'runs': [{'t': "Первая в России ", 's': 12.5, 'b': True, 'c': DARKSLATE},
                {'t': "система комплексной аналитики, основанная на реальных сделках", 's': 12, 'b': False, 'c': TXTGREY2}],
       'ls': 1.0}], anchor=MSO_ANCHOR.MIDDLE)

# image on the right with frame (preserve aspect 1459x680)
img_w = 5.50
img_h = img_w * 680 / 1459.0
img_x = 7.25
img_y = 2.55
fx, fy = img_x - 0.12, img_y - 0.12
fw, fh = img_w + 0.24, img_h + 0.24
# soft glow behind frame
g = rrect(s, fx, fy, fw, fh, fill=None, radius=0.22)
glow(g, PRIMARY, alpha=0.16, radius_pt=10)
rrect(s, fx, fy, fw, fh, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.22)
# tab label overlapping image top edge
tab_label(s, img_x + 0.55, img_y - 0.21, 3.7, "АНАЛИТИКА РЫНКА НЕДВИЖИМОСТИ",
          fill=PRIMARY, h=0.42, size=11)
# picture
s.shapes.add_picture(A + "Аналитика РН.png", Inches(img_x), Inches(img_y),
                     width=Inches(img_w), height=Inches(img_h))

# caption card below the image
cap_y = fy + fh + 0.28
rrect(s, fx, cap_y, fw, 1.32, fill=PALEBLUE, lcolor=PRIMARY, lw=1.0, radius=0.22)
check_bullet(s, fx + 0.22, cap_y + 0.34, 0.40, TEAL, size=15)
text(s, fx + 0.72, cap_y + 0.14, fw - 1.0, 1.04,
     [{'runs': [{'t': "Представление данных", 's': 12.5, 'b': True, 'c': PRIMARY}], 'sa': 3, 'ls': 0.95},
      {'runs': [{'t': "таблицы, графики и тепловые карты с гибкой настройкой фильтров (площадь, границы до кадастрового квартала)",
                 's': 10.5, 'b': False, 'c': TXTGREY2}], 'ls': 0.98}],
     anchor=MSO_ANCHOR.MIDDLE)

# ============================================================================
# SLIDE 8 — ЗАКОН № 295-ФЗ (ВРИ)
# ============================================================================
s = slide()
header(s, "Закон № 295-ФЗ: виды разрешенного использования (ВРИ)", 8)

text(s, 0.6, 1.26, 12.1, 0.4,
     [{'runs': [{'t': "С 1 марта 2026 года вступил в силу Закон № 295-ФЗ, закрепляющий новую главу «Виды разрешенного использования земельных участков»",
                 's': 13, 'b': False, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# left badge
bx, by, bw, bh = 0.6, 1.85, 3.2, 4.95
rrect(s, bx, by, bw, bh, fill=PRIMARY, radius=0.28)
rrect(s, bx, by, bw, bh, fill=PRIMARY, radius=0.28)
eagle(s, bx + (bw - 1.3) / 2, by + 0.5, 1.3, on_blue=False)
text(s, bx + 0.3, by + 2.0, bw - 0.6, 1.6,
     [{'runs': [{'t': "Закон", 's': 18, 'b': True, 'c': WHITE}], 'align': PP_ALIGN.CENTER, 'sa': 2},
      {'runs': [{'t': "№ 295-ФЗ", 's': 30, 'b': True, 'c': WHITE}], 'align': PP_ALIGN.CENTER, 'sa': 6},
      {'runs': [{'t': "ВРИ земельных участков", 's': 12, 'b': False, 'c': LIGHTBLUE}], 'align': PP_ALIGN.CENTER, 'ls': 1.0}],
     anchor=MSO_ANCHOR.MIDDLE)

items8 = [
    ("Устранение неопределенности", "ВРИ теперь устанавливаются не только градостроительным регламентом, но и лесохозяйственным регламентом, положением об особо охраняемой территории"),
    ("Соответствие ВРИ", "ВРИ здания или сооружения должен соответствовать ВРИ участка (исключения: линейные объекты, геодезические пункты, объекты обороны)"),
    ("Объединение участков", "разрешено объединять и перераспределять участки с разными ВРИ при соблюдении условий (госнужды, одна территориальная зона)"),
    ("Запрет на изменение ВРИ", "для участков в государственной или муниципальной собственности (аренда, безвозмездное пользование)"),
    ("Сельхозугодья", "переходные положения: субъекты РФ получают полномочия по изменению ВРИ, но без права установления ВРИ, не соответствующего целевому назначению"),
]
ix, iy, iw = 4.15, 1.85, 8.58
iycur = iy
for head, body in items8:
    ih = 0.90
    rrect(s, ix, iycur, iw, ih, fill=WHITE, lcolor=PRIMARY, lw=1.0, radius=0.22)
    g = rrect(s, ix, iycur, iw, ih, fill=None, radius=0.22); glow(g, PRIMARY, alpha=0.10, radius_pt=8)
    rrect(s, ix, iycur, iw, ih, fill=WHITE, lcolor=PRIMARY, lw=1.0, radius=0.22)
    check_bullet(s, ix + 0.22, iycur + 0.24, 0.42, TEAL, size=16)
    text(s, ix + 0.82, iycur + 0.10, iw - 1.0, ih - 0.18,
         [{'runs': [{'t': head, 's': 12.5, 'b': True, 'c': PRIMARY}], 'sa': 2, 'ls': 0.95},
          {'runs': [{'t': body, 's': 10.5, 'b': False, 'c': TXTGREY2}], 'ls': 0.95}],
         anchor=MSO_ANCHOR.MIDDLE)
    iycur += ih + 0.09

# ============================================================================
# SLIDE 9 — ЗАКОН № 301-ФЗ (ЗОНЫ ЗАТОПЛЕНИЯ)
# ============================================================================
s = slide()
header(s, "Закон № 301-ФЗ: зоны затопления и подтопления", 9)

text(s, 0.6, 1.26, 12.1, 0.4,
     [{'runs': [{'t': "Федеральный закон от 4 августа 2026 г. № 301-ФЗ вносит изменения в Водный кодекс РФ и закон «О безопасности гидротехнических сооружений»",
                 's': 13, 'b': False, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)

# left big prohibition card
bx, by, bw, bh = 0.6, 1.9, 4.2, 4.55
rrect(s, bx, by, bw, bh, fill=PRIMARY, radius=0.28)
rrect(s, bx, by, bw, bh, fill=PRIMARY, radius=0.28)
# prohibition icon: white circle + orange diagonal
ixc, icy, ics = bx + bw / 2 - 0.75, by + 0.65, 1.5
ring = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(ixc), Inches(icy), Inches(ics), Inches(ics))
_noshadow(ring); solid(ring, WHITE); line(ring, ORANGE, 3.0)
cx0, cy0 = ixc + ics / 2, icy + ics / 2
lw_, lh_ = ics * 1.35, 0.16
diag = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx0 - lw_ / 2), Inches(cy0 - lh_ / 2),
                          Inches(lw_), Inches(lh_))
_noshadow(diag); solid(diag, ORANGE); diag.rotation = 45
text(s, bx + 0.3, by + 2.25, bw - 0.6, 1.9,
     [{'runs': [{'t': "ЗАПРЕЩЕНО", 's': 26, 'b': True, 'c': WHITE}], 'align': PP_ALIGN.CENTER, 'sa': 6},
      {'runs': [{'t': "строительство объектов капитального строительства в границах зон затопления и подтопления",
                 's': 13, 'b': False, 'c': LIGHTBLUE}], 'align': PP_ALIGN.CENTER, 'ls': 1.05}],
     anchor=MSO_ANCHOR.TOP)

# right content
right_items = [
    ("Исключения", "разрешено только строительство сооружений инженерной защиты территорий и объектов (водоограждающие дамбы, берегоукрепительные сооружения и т.д.)"),
    ("Вступление в силу", "Закон вступил в силу с 4 августа 2026 года"),
    ("Контроль Росреестра", "строгое соблюдение норм контролируется при осуществлении учётно-регистрационных действий"),
]
rx, ry, rw = 5.15, 1.9, 7.58
iycur = ry
for head, body in right_items:
    ih = 1.35
    rrect(s, rx, iycur, rw, ih, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.24)
    g = rrect(s, rx, iycur, rw, ih, fill=None, radius=0.24); glow(g, PRIMARY, alpha=0.12, radius_pt=9)
    rrect(s, rx, iycur, rw, ih, fill=WHITE, lcolor=PRIMARY, lw=1.25, radius=0.24)
    check_bullet(s, rx + 0.28, iycur + 0.32, 0.5, TEAL, size=18)
    text(s, rx + 1.0, iycur + 0.22, rw - 1.3, ih - 0.4,
         [{'runs': [{'t': head, 's': 15, 'b': True, 'c': PRIMARY}], 'sa': 4, 'ls': 0.95},
          {'runs': [{'t': body, 's': 12.5, 'b': False, 'c': TXTGREY2}], 'ls': 1.0}],
         anchor=MSO_ANCHOR.MIDDLE)
    iycur += ih + 0.25

# ============================================================================
# SLIDE 10 — ИНФОРМАЦИОНАЯ ПОЛИТИКА
# ============================================================================
s = slide()
header(s, "Информационная политика и открытость", 10)

text(s, 0.6, 1.26, 12.1, 0.4,
     [{'runs': [{'t': "Росреестр много лет подряд — лидер рейтинга по открытости федеральных органов исполнительной власти; Управление держится в топе лидеров среди территориальных управлений",
                 's': 13, 'b': False, 'c': TXTGREY}],
       'align': PP_ALIGN.LEFT, 'ls': 1.0}], anchor=MSO_ANCHOR.TOP)

# left highlight card
lx, ly, lw, lh = 0.6, 1.95, 4.4, 4.5
rrect(s, lx, ly, lw, lh, fill=PRIMARY, radius=0.28)
rrect(s, lx, ly, lw, lh, fill=PRIMARY, radius=0.28)
eagle(s, lx + (lw - 1.5) / 2, ly + 0.45, 1.5, on_blue=False)
text(s, lx + 0.4, ly + 2.15, lw - 0.8, 2.1,
     [{'runs': [{'t': "ЛИДЕР", 's': 16, 'b': True, 'c': WHITE}], 'align': PP_ALIGN.CENTER, 'sa': 4},
      {'runs': [{'t': "по открытости", 's': 16, 'b': True, 'c': WHITE}], 'align': PP_ALIGN.CENTER, 'sa': 8},
      {'runs': [{'t': "информационной политики и доверия граждан", 's': 12, 'b': False, 'c': LIGHTBLUE}],
       'align': PP_ALIGN.CENTER, 'ls': 1.05}], anchor=MSO_ANCHOR.TOP)

# right channels list
channels = [
    ("Средства массовой информации", "регулярные выступления в СМИ и публикации в печатных изданиях"),
    ("«Горячие линии»", "телефонные консультации дважды в месяц по кадастровому учёту и регистрации прав"),
    ("Онлайн-консультации", "видеоконференцсвязь на площадке МФЦ; еженедельные консультации в филиалах «Мои Документы»"),
    ("Социальные сети и мессенджеры", "оперативные новости, ответы на вопросы, отчёты с мероприятий"),
]
rx, ry, rw = 5.25, 1.95, 7.48
iycur = ry
for head, body in channels:
    ih = 0.98
    rrect(s, rx, iycur, rw, ih, fill=WHITE, lcolor=PRIMARY, lw=1.0, radius=0.22)
    g = rrect(s, rx, iycur, rw, ih, fill=None, radius=0.22); glow(g, PRIMARY, alpha=0.10, radius_pt=8)
    rrect(s, rx, iycur, rw, ih, fill=WHITE, lcolor=PRIMARY, lw=1.0, radius=0.22)
    check_bullet(s, rx + 0.24, iycur + 0.26, 0.44, TEAL, size=16)
    text(s, rx + 0.82, iycur + 0.12, rw - 1.05, ih - 0.2,
         [{'runs': [{'t': head, 's': 13, 'b': True, 'c': PRIMARY}], 'sa': 2, 'ls': 0.95},
          {'runs': [{'t': body, 's': 11, 'b': False, 'c': TXTGREY2}], 'ls': 0.95}],
         anchor=MSO_ANCHOR.MIDDLE)
    iycur += ih + 0.10

# social pills
socials = ["ВКонтакте", "Одноклассники", "Яндекс.Дзен", "МАХ", "Rutube"]
sx, sy = rx + 0.1, iycur + 0.02
sw_ = 1.44
for i, so in enumerate(socials):
    x = sx + i * (sw_ + 0.08)
    pill(s, x, sy, sw_, 0.42, fill=PALEBLUE, lcolor=SECONDARY, lw=1.0, radius=0.21)
    text(s, x, sy, sw_, 0.42,
         [{'runs': [{'t': so, 's': 10.5, 'b': True, 'c': DARKSLATE}], 'align': PP_ALIGN.CENTER}],
         anchor=MSO_ANCHOR.MIDDLE)

# ============================================================================
# SLIDE 11 — СПАСИБО
# ============================================================================
s = slide()
band_top_h = 2.08
band_bot_h = 1.88
rect(s, 0, 0, SW, band_top_h, g1=PRIMARY, g2=SECONDARY, ang=90)
rect(s, 0, SH - band_bot_h, SW, band_bot_h, g1=SECONDARY, g2=PRIMARY, ang=90)
# decorative gradient circle bottom-left (crosses the bands, avoids centered text)
circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-1.9), Inches(SH - 3.7), Inches(5.4), Inches(5.4))
_noshadow(circ); gradient(circ, PRIMARY, TEAL, 135); glow(circ, TEAL, alpha=0.3, radius_pt=16)

eagle(s, 0.62, 0.55, 1.05, on_blue=True)
text(s, 1.95, 0.62, 7.5, 0.9,
     [{'runs': [{'t': "УПРАВЛЕНИЕ РОСРЕЕСТРА", 's': 15, 'b': True, 'c': WHITE}], 'sa': 2},
      {'runs': [{'t': "ВОРОНЕЖСКАЯ ОБЛАСТЬ", 's': 12, 'b': False, 'c': LIGHTBLUE}]}],
     anchor=MSO_ANCHOR.MIDDLE)

text(s, 1.2, 2.70, 10.93, 1.4,
     [{'runs': [{'t': "Благодарю за внимание!", 's': 40, 'b': True, 'c': PRIMARY}],
       'align': PP_ALIGN.CENTER, 'sa': 8, 'ls': 1.0},
      {'runs': [{'t': "Готова ответить на ваши вопросы", 's': 16, 'b': False, 'c': DARKSLATE}],
       'align': PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.TOP)

text(s, 1.2, 4.45, 10.93, 0.6,
     [{'runs': [{'t': "Приглашаю подписаться на наши официальные страницы для получения достоверной и проверенной информации",
                 's': 13, 'b': False, 'i': True, 'c': TXTGREY}], 'align': PP_ALIGN.CENTER, 'ls': 1.0}],
     anchor=MSO_ANCHOR.TOP)

text(s, 0.95, SH - band_bot_h + 0.45, 11.5, 0.8,
     [{'runs': [{'t': "ФОРУМ ПО НЕДВИЖИМОСТИ В ЧЕРНОЗЕМЬЕ  ·  2026", 's': 12, 'b': True, 'c': WHITE}],
       'align': PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.MIDDLE)

# ---- save ----
out = "Росреестр_Воронежская_область_Доклад.pptx"
prs.save(out)
print("SAVED", out, "slides:", len(prs.slides._sldIdLst))
