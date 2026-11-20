# -*- coding: utf-8 -*-
"""Render a python-pptx deck to a multi-page PDF (custom renderer, no LibreOffice)."""
import io, math
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from pptx import Presentation
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

SRC = "Росреестр_Воронежская_область_Доклад.pptx"
OUT = "Росреестр_Воронежская_область_Доклад.pdf"
DPI = 150.0

FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
_fc = {}
def get_font(size_pt, bold):
    key = (round(size_pt, 1), bool(bold))
    if key not in _fc:
        px = max(6, int(round(size_pt * DPI / 72.0)))
        _fc[key] = ImageFont.truetype(FONT_BOLD if bold else FONT_REG, px)
    return _fc[key]

def px(emu):
    return int(round(emu / 914400.0 * DPI))
def in_(emu):
    return emu / 914400.0

def rgb_of(color_obj, default=(0x59, 0x59, 0x59)):
    try:
        if color_obj and color_obj.type is not None and color_obj.rgb is not None:
            h = str(color_obj.rgb)
            return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        pass
    return default

def srgb_from_hex(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

# ---------- geometry helpers ----------
def get_prst(shape):
    try:
        sp = shape._element.spPr
        g = sp.find(qn('a:prstGeom'))
        return g.get('prst') if g is not None else 'rect'
    except Exception:
        return 'rect'

def get_radius_px(shape, w_px, h_px):
    try:
        adj = shape.adjustments[0]
        return adj * min(w_px, h_px)
    except Exception:
        return 0

def make_mask(w, h, prst, radius):
    m = Image.new('L', (w, h), 0)
    d = ImageDraw.Draw(m)
    if prst == 'ellipse':
        d.ellipse([0, 0, w - 1, h - 1], fill=255)
    elif prst == 'roundRect':
        d.rounded_rectangle([0, 0, w - 1, h - 1], radius=radius, fill=255)
    elif prst == 'chevron':
        n = w * 0.30
        pts = [(0, 0), (w - n, 0), (w, h // 2), (w - n, h), (0, h)]
        d.polygon(pts, fill=255)
    elif prst == 'downArrow':
        marg = w * 0.26
        ah = h * 0.42
        pts = [(marg, 0), (w - marg, 0), (w - marg, h - ah), (w, h),
               (0, h), (0, h - ah), (marg, h - ah)]
        d.polygon(pts, fill=255)
    else:  # rect
        d.rectangle([0, 0, w - 1, h - 1], fill=255)
    return m

# ---------- fill / line / glow ----------
def get_fill(shape):
    try:
        ft = shape.fill.type
        if ft == MSO_FILL_TYPE.SOLID:
            return ('solid', rgb_of(shape.fill.fore_color))
        if ft == MSO_FILL_TYPE.GRADIENT:
            stops = shape.fill.gradient_stops
            c1 = rgb_of(stops[0].color)
            c2 = rgb_of(stops[1].color)
            ang = 135.0
            gf = shape.fill._xPr.find(qn('a:gradFill'))
            if gf is not None and gf.get('rot'):
                ang = int(gf.get('rot')) / 60000.0
            return ('grad', c1, c2, ang)
    except Exception:
        pass
    return ('none',)

def get_line(shape):
    try:
        lc = shape.line.color
        if lc is not None and lc.type is not None and lc.rgb is not None:
            w = shape.line.width.pt if shape.line.width is not None else 1.0
            return (rgb_of(lc), max(1, px(int(w * 12700))))
    except Exception:
        pass
    return None

def get_glow(shape):
    try:
        spPr = shape._element.spPr
        eff = spPr.find(qn('a:effectLst'))
        if eff is None:
            return None
        g = eff.find(qn('a:glow'))
        if g is None:
            return None
        clr = g.find(qn('a:srgbClr'))
        color = srgb_from_hex(clr.get('val')) if clr is not None else (0, 140, 255)
        al = g.find(qn('a:alpha'))
        alpha = int(al.get('val')) / 100000.0 if al is not None else 0.4
        rad = px(int(g.get('rad'))) if g.get('rad') else 12
        return (color, alpha, rad)
    except Exception:
        return None

def make_solid_layer(w, h, fill, mask):
    if fill[0] == 'solid':
        img = Image.new('RGB', (w, h), fill[1])
    elif fill[0] == 'grad':
        _, c1, c2, ang = fill
        ang = math.radians(ang)
        dx, dy = math.cos(ang), math.sin(ang)
        yy, xx = np.mgrid[0:h, 0:w]
        t = xx * dx + yy * dy
        t = t - t.min()
        t = t / (t.max() + 1e-9)
        a = np.asarray(c1, dtype=float)[None, None, :]
        b = np.asarray(c2, dtype=float)[None, None, :]
        rgb = (a * (1 - t[..., None]) + b * t[..., None]).astype('uint8')
        img = Image.fromarray(rgb, 'RGB')
    else:
        return None
    img.putalpha(mask)
    return img

def build_layer(shape):
    w, h = px(shape.width), px(shape.height)
    prst = get_prst(shape)
    radius = get_radius_px(shape, w, h)
    mask = make_mask(w, h, prst, radius)
    fill = get_fill(shape)
    layer = make_solid_layer(w, h, fill, mask)
    if layer is None:
        layer = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    else:
        layer = layer.convert('RGBA')
    line = get_line(shape)
    if line:
        d = ImageDraw.Draw(layer)
        col, wd = line
        if prst == 'ellipse':
            d.ellipse([0, 0, w - 1, h - 1], outline=col, width=wd)
        elif prst == 'roundRect':
            d.rounded_rectangle([0, 0, w - 1, h - 1], radius=radius, outline=col, width=wd)
        elif prst == 'rect':
            d.rectangle([0, 0, w - 1, h - 1], outline=col, width=wd)
    return layer, prst, radius

def composite(canvas, shape, layer):
    x, y, w, h = px(shape.left), px(shape.top), px(shape.width), px(shape.height)
    cx, cy = x + w / 2.0, y + h / 2.0
    # glow behind
    glow = get_glow(shape)
    if glow:
        gcol, galpha, grad = glow
        pad = int(grad * 2)
        mw, mh = w + 2 * pad, h + 2 * pad
        m = make_mask(w, h, get_prst(shape), get_radius_px(shape, w, h))
        big = Image.new('L', (mw, mh), 0)
        big.paste(m, (pad, pad))
        big = big.filter(ImageFilter.GaussianBlur(grad))
        halo = Image.new('RGB', (mw, mh), gcol)
        mask = big.point(lambda v: int(v * galpha))
        canvas.paste(halo, (int(cx - mw / 2), int(cy - mh / 2)), mask)
    # layer
    rot = getattr(shape, 'rotation', 0) or 0
    if rot:
        layer = layer.rotate(rot, expand=True, resample=Image.BICUBIC)
        canvas.paste(layer, (int(cx - layer.width / 2), int(cy - layer.height / 2)), layer)
    else:
        canvas.paste(layer, (x, y), layer)

def paste_picture(canvas, shape):
    blob = shape.image.blob
    img = Image.open(io.BytesIO(blob)).convert('RGBA')
    w, h = px(shape.width), px(shape.height)
    img = img.resize((w, h), Image.LANCZOS)
    canvas.paste(img, (px(shape.left), px(shape.top)), img)

# ---------- text ----------
def run_color(run, default=(0x59, 0x59, 0x59)):
    return rgb_of(run.font.color, default)

def layout_lines(tf, avail_w):
    lines = []
    for p in tf.paragraphs:
        segs = []
        for run in p.runs:
            if not run.text:
                continue
            size = run.font.size.pt if run.font.size else 12
            bold = bool(run.font.bold)
            col = run_color(run)
            f = get_font(size, bold)
            for w in run.text.split(' '):
                if w != '':
                    segs.append((w, f, col, size))
        align = p.alignment
        if not segs:
            lines.append(([], 12, align))
            continue
        cur, cur_w, cur_size = [], 0.0, segs[0][3]
        for (w, f, col, size) in segs:
            ww = f.getlength(w)
            if cur and cur_w + ww > avail_w:
                lines.append((cur, cur_size, align))
                cur, cur_w, cur_size = [], 0.0, size
            cur.append((w, f, col, size))
            cur_w += ww
            cur_size = max(cur_size, size)
        if cur:
            lines.append((cur, cur_size, align))
    return lines

def draw_text(canvas, shape):
    tf = shape.text_frame
    x, y, w, h = px(shape.left), px(shape.top), px(shape.width), px(shape.height)
    ml = px(tf.margin_left or 0); mr = px(tf.margin_right or 0)
    mt = px(tf.margin_top or 0); mb = px(tf.margin_bottom or 0)
    avail_w = max(1, w - ml - mr)
    lines = layout_lines(tf, avail_w)
    line_hs = [max(2, int(round(sz * 1.26 * DPI / 72.0))) for (_, sz, _) in lines]
    total = sum(line_hs)
    anchor = tf.vertical_anchor
    if anchor == 1:  # MIDDLE
        y0 = y + (h - total) / 2.0
    elif anchor == 2:  # BOTTOM
        y0 = y + (h - total)
    else:
        y0 = y
    d = ImageDraw.Draw(canvas)
    cy = y0
    for (segs, sz, align), lh in zip(lines, line_hs):
        if not segs:
            cy += lh
            continue
        line_w = sum(f.getlength(w) for (w, f, _, _) in segs) + \
                 sum(f.getlength(' ') for (_, f, _, _) in segs)
        if align == 1:  # CENTER
            lx = x + (w - line_w) / 2.0
        elif align == 2:  # RIGHT
            lx = x + (w - line_w)
        else:
            lx = x
        for (wd, f, col, _) in segs:
            d.text((lx, cy), wd, font=f, fill=col)
            lx += f.getlength(wd) + f.getlength(' ')
        cy += lh

# ---------- table ----------
def render_table(canvas, shape):
    tbl = shape.table
    gf_left, gf_top = px(shape.left), px(shape.top)
    col_ws = [px(c.width) for c in tbl.columns]
    row_hs = [px(r.height) for r in tbl.rows]
    for r_i, row in enumerate(tbl.rows):
        y = gf_top + sum(row_hs[:r_i])
        for c_i, cell in enumerate(row.cells):
            x = gf_left + sum(col_ws[:c_i])
            w, h = col_ws[c_i], row_hs[r_i]
            try:
                ft = cell.fill.type
                if ft == MSO_FILL_TYPE.SOLID:
                    col = rgb_of(cell.fill.fore_color)
                    canvas.paste(Image.new('RGB', (w, h), col), (x, y))
            except Exception:
                pass
            d = ImageDraw.Draw(canvas)
            d.rectangle([x, y, x + w - 1, y + h - 1], outline=(0xD9, 0xD9, 0xD9), width=1)
            tf = cell.text_frame
            p = tf.paragraphs[0]
            if not p.runs:
                continue
            size = p.runs[0].font.size.pt if p.runs[0].font.size else 11
            bold = bool(p.runs[0].font.bold)
            col = run_color(p.runs[0])
            f = get_font(size, bold)
            txt = tf.text
            tw = f.getlength(txt)
            align = p.alignment
            if align == 1:
                tx = x + (w - tw) / 2.0
            else:
                tx = x + 6
            ty = y + (h - size * DPI / 72.0) / 2.0
            d.text((tx, ty), txt, font=f, fill=col)

# ---------- main ----------
def main():
    prs = Presentation(SRC)
    pages = []
    for sl in prs.slides:
        W = px(prs.slide_width); H = px(prs.slide_height)
        canvas = Image.new('RGB', (W, H), (0xFF, 0xFF, 0xFF))
        for shape in sl.shapes:
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    paste_picture(canvas, shape)
                    continue
                if getattr(shape, 'has_table', False) and shape.has_table:
                    render_table(canvas, shape)
                    continue
                layer, prst, radius = build_layer(shape)
                if layer.getbbox() is not None or get_fill(shape)[0] != 'none':
                    composite(canvas, shape, layer)
                if shape.has_text_frame and shape.text_frame.text.strip():
                    draw_text(canvas, shape)
            except Exception as e:
                print("  ! shape error:", e)
        pages.append(canvas)
    pages[0].save(OUT, 'PDF', save_all=True, append_images=pages[1:],
                  resolution=DPI)
    print("SAVED", OUT, "pages:", len(pages), "size:", pages[0].size)

if __name__ == "__main__":
    main()
