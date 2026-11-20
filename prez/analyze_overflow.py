# -*- coding: utf-8 -*-
"""Heuristic overflow checker: estimate whether text fits inside each shape box."""
from pptx import Presentation
from pptx.util import Emu
from PIL import ImageFont

DPI = 150.0
FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

_cache = {}
def font(size_pt, bold):
    key = (round(size_pt, 1), bold)
    if key not in _cache:
        px = max(6, int(round(size_pt * DPI / 72.0)))
        _cache[key] = ImageFont.truetype(FONT_BOLD if bold else FONT_REG, px)
    return _cache[key]

def emu_in(v):
    return v / 914400.0

def color_of(run):
    try:
        c = run.font.color
        if c and c.type is not None and c.rgb is not None:
            return c.rgb
    except Exception:
        pass
    return None

def analyze(path):
    prs = Presentation(path)
    issues = []
    for si, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            if not tf.text.strip():
                continue
            box_w = emu_in(sh.width)
            box_h = emu_in(sh.height)
            ml = emu_in(tf.margin_left or 0)
            mr = emu_in(tf.margin_right or 0)
            mt = emu_in(tf.margin_top or 0)
            mb = emu_in(tf.margin_bottom or 0)
            avail_w = box_w - ml - mr
            avail_h = box_h - mt - mb
            total_h = 0.0
            max_line_w = 0.0
            for p in tf.paragraphs:
                runs = p.runs
                if not runs:
                    total_h += 0.2
                    continue
                # size from first run with a size
                psize = None
                for r in runs:
                    if r.font.size:
                        psize = r.font.size.pt
                        break
                psize = psize or 12
                bold = bool(runs[0].font.bold)
                f = font(psize, bold)
                # build concatenated text and per-run widths
                line_words = []
                for r in runs:
                    txt = r.text
                    if not txt:
                        continue
                    for w in txt.split(" "):
                        if w:
                            line_words.append(w)
                # wrap
                lines = []
                cur = ""
                for w in line_words:
                    trial = (cur + " " + w).strip() if cur else w
                    if f.getlength(trial) <= avail_w * DPI:
                        cur = trial
                    else:
                        if cur:
                            lines.append(cur)
                        # word itself may be too long
                        cur = w
                if cur:
                    lines.append(cur)
                if not lines:
                    lines = [""]
                for ln in lines:
                    max_line_w = max(max_line_w, f.getlength(ln) / DPI)
                nlines = len(lines)
                lh = psize * 1.22 / 72.0  # line height inches
                sa = (p.space_after.pt / 72.0) if p.space_after else 0.0
                sb = (p.space_before.pt / 72.0) if p.space_before else 0.0
                total_h += nlines * lh + sa + sb
            if total_h > avail_h * 1.12 and avail_h > 0:
                issues.append((si, sh.shape_type, round(box_w,2), round(box_h,2),
                               round(total_h,2), tf.text.strip().replace("\n"," | ")[:70]))
    return issues

if __name__ == "__main__":
    iss = analyze("Росреестр_Воронежская_область_Доклад.pptx")
    if not iss:
        print("No overflow risks detected.")
    else:
        print(f"Potential overflow in {len(iss)} shape(s):")
        for si, st, w, h, th, txt in iss:
            print(f"  slide {si}: box {w}x{h}in, est text {th}in  -> {txt}")
