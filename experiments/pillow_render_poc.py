# -*- coding: utf-8 -*-
"""
Pillow 纯 Python PPTX 渲染 PoC

受限模板规范（本项目海报模板家族）：
- 全屏背景图（p:pic）
- 文本框：字体/字号/粗体/字距(spc)/行距(spcPct)/对齐(algn)/anchor
- 文字填充：纯色(srgbClr/schemeClr) 或 线性渐变(gradFill + lin ang)
- 文字外阴影(outerShdw: blurRad/dist/dir/颜色/alpha)
不支持：旋转、组合、表格、图表、艺术字变形、图片裁剪、描边、glow

用法：
    .venv/Scripts/python.exe experiments/pillow_render_poc.py --pptx "10月递交贺报(3).pptx" --quality 2
"""
import argparse
import io
import math
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

# ---------------- 校准常数（对照 PowerPoint 导出图调整） ----------------
LINE_HEIGHT_RATIO = 1.25   # 微软雅黑 单倍行距 / 字号（实测 PowerPoint 108px/28pt@120% 校准）
BASELINE0_FRAC = 0.72      # 首行基线在行高盒中的位置（二轮校准，残余 |Δtop|≤9px@q2）
SHADOW_SIGMA_FRAC = 0.5    # PowerPoint blurRad -> Pillow GaussianBlur sigma 的系数
# -----------------------------------------------------------------------

FONT_FILES = {
    ('微软雅黑', False): r'C:\Windows\Fonts\msyh.ttc',
    ('微软雅黑', True): r'C:\Windows\Fonts\msyhbd.ttc',
}
DEFAULT_FONT = r'C:\Windows\Fonts\msyh.ttc'

EMU_PER_PT = 12700
PT_PER_IN = 72


def hex2rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def lerp(a, b, t):
    return a + (b - a) * t


def lerp_color(c1, c2, t):
    return tuple(int(round(lerp(a, b, t))) for a, b in zip(c1, c2))


class Theme:
    """从 pptx zip 解析 theme1.xml 的配色，供 schemeClr 映射"""

    def __init__(self, pptx_path):
        self.colors = {'dk1': '#000000', 'lt1': '#FFFFFF', 'dk2': '#000000', 'lt2': '#FFFFFF'}
        try:
            with zipfile.ZipFile(pptx_path) as z:
                xml = z.read('ppt/theme/theme1.xml').decode('utf-8')
            for name in ('dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
                         'accent4', 'accent5', 'accent6'):
                m = re.search(rf'<a:{name}>(.*?)</a:{name}>', xml, re.S)
                if not m:
                    continue
                inner = m.group(1)
                m2 = re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', inner)
                if m2:
                    self.colors[name] = '#' + m2.group(1).upper()
                    continue
                m2 = re.search(r'sysClr val="\w+" lastClr="([0-9A-Fa-f]{6})"', inner)
                if m2:
                    self.colors[name] = '#' + m2.group(1).upper()
        except Exception:
            pass


def parse_color_el(clr_el, theme):
    """解析颜色元素（srgbClr/schemeClr/sysClr），返回 (rgb_tuple, alpha 0~1)"""
    from lxml import etree
    tag = etree.QName(clr_el).localname
    alpha = 1.0
    if tag == 'srgbClr':
        rgb = hex2rgb(clr_el.get('val'))
    elif tag == 'schemeClr':
        rgb = hex2rgb(theme.colors.get(clr_el.get('val'), '#000000'))
    elif tag == 'sysClr':
        rgb = hex2rgb(clr_el.get('lastClr') or '000000')
    else:
        rgb = (0, 0, 0)
    a = clr_el.find(qn('a:alpha'))
    if a is not None:
        alpha = int(a.get('val')) / 100000.0
    return rgb, alpha


def parse_run_fill(rPr, theme):
    """返回 None（继承=黑） | ('solid', rgb, alpha) | ('grad', [stops], ang_deg)"""
    if rPr is None:
        return None
    grad = rPr.find(qn('a:gradFill'))
    if grad is not None:
        stops = []
        for gs in grad.findall(qn('a:gsLst') + '/' + qn('a:gs')):
            pos = int(gs.get('pos')) / 100000.0
            clr = list(gs)[0]
            rgb, a = parse_color_el(clr, theme)
            stops.append((pos, rgb, a))
        lin = grad.find(qn('a:lin'))
        ang = int(lin.get('ang')) / 60000.0 if lin is not None else 0.0
        return ('grad', stops, ang)
    solid = rPr.find(qn('a:solidFill'))
    if solid is not None:
        clr = list(solid)[0]
        rgb, a = parse_color_el(clr, theme)
        return ('solid', rgb, a)
    return None


def parse_shadow(rPr):
    """返回 None | dict(blur_pt, dist_pt, dir_deg, rgb, alpha)"""
    if rPr is None:
        return None
    shdw = rPr.find(qn('a:effectLst') + '/' + qn('a:outerShdw'))
    if shdw is None:
        return None
    clr = list(shdw)[0]
    # 颜色与 alpha 需要一起返回，这里偷懒用全局 theme 后补
    return {
        'blur_pt': int(shdw.get('blurRad', '0')) / EMU_PER_PT,
        'dist_pt': int(shdw.get('dist', '0')) / EMU_PER_PT,
        'dir_deg': int(shdw.get('dir', '0')) / 60000.0,
        'clr_el': clr,
    }


def tokenize(text):
    """CJK 逐字、英文按词、空格独立成 token（换行断点）"""
    return [m.group(0) for m in re.finditer(r'[A-Za-z0-9]+(?:[.\u00b7][A-Za-z0-9]+)*|\s+|.', text)]


class Style:
    __slots__ = ('size_pt', 'bold', 'spc_pt', 'fill', 'shadow', 'font_path')

    def __init__(self, size_pt, bold, spc_pt, fill, shadow, font_path):
        self.size_pt, self.bold, self.spc_pt = size_pt, bold, spc_pt
        self.fill, self.shadow, self.font_path = fill, shadow, font_path

    def key(self):
        return (self.size_pt, self.bold, self.spc_pt, repr(self.fill), repr(self.shadow))


def read_style(run, theme):
    rPr = run._r.find(qn('a:rPr'))
    sz = 18.0
    b = False
    spc = 0.0
    font_path = DEFAULT_FONT
    if rPr is not None:
        if rPr.get('sz'):
            sz = int(rPr.get('sz')) / 100.0
        b = rPr.get('b') == '1'
        if rPr.get('spc'):
            spc = int(rPr.get('spc')) / 100.0
        for tag in ('a:ea', 'a:latin'):
            tf = rPr.find(qn(tag))
            if tf is not None and tf.get('typeface'):
                path = FONT_FILES.get((tf.get('typeface'), b))
                if path:
                    font_path = path
                    break
    fill = parse_run_fill(rPr, theme)
    shadow = parse_shadow(rPr)
    return Style(sz, b, spc, fill, shadow, font_path)


def make_linear_gradient(w, h, stops, ang_deg):
    """按 stops [(pos, rgb, alpha)] 生成 w×h RGBA 线性渐变图。ang 0°=左→右，90°=上→下"""
    theta = math.radians(ang_deg)
    vx, vy = math.cos(theta), math.sin(theta)
    corners = [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]
    projs = [x * vx + y * vy for x, y in corners]
    p0, p1 = min(projs), max(projs)
    span = (p1 - p0) or 1.0
    grad = Image.new('RGBA', (w, h))
    px = grad.load()
    for y in range(h):
        for x in range(w):
            t = ((x * vx + y * vy) - p0) / span
            t = min(max(t, 0.0), 1.0)
            for i in range(len(stops) - 1):
                pa, ca, aa = stops[i]
                pb, cb, ab = stops[i + 1]
                if pa <= t <= pb:
                    tt = (t - pa) / (pb - pa) if pb > pa else 0.0
                    px[x, y] = (*lerp_color(ca, cb, tt),
                                int(round(lerp(aa, ab, tt) * 255)))
                    break
            else:
                pos, c, a = stops[-1] if t >= stops[-1][0] else stops[0]
                px[x, y] = (*c, int(round(a * 255)))
    return grad


class PillowRenderer:
    def __init__(self, pptx_path, quality=1.0):
        self.prs = Presentation(pptx_path)
        self.quality = quality
        self.dpi = 96.0 * quality
        self.theme = Theme(pptx_path)
        self.w_emu, self.h_emu = self.prs.slide_width, self.prs.slide_height
        self.scale = self.dpi / 914400.0  # EMU -> px
        self.font_cache = {}

    def emu2px(self, v):
        return v * self.scale

    def pt2px(self, v):
        return v * self.dpi / PT_PER_IN

    def font(self, style):
        key = (style.font_path, style.size_pt)
        if key not in self.font_cache:
            self.font_cache[key] = ImageFont.truetype(style.font_path, self.pt2px(style.size_pt))
        return self.font_cache[key]

    def render(self):
        W = int(round(self.w_emu * self.scale))
        H = int(round(self.h_emu * self.scale))
        canvas = Image.new('RGBA', (W, H), (255, 255, 255, 255))
        slide = self.prs.slides[0]
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                self._draw_picture(canvas, shape)
            elif shape.has_text_frame:
                self._draw_text_frame(canvas, shape)
            else:
                print(f'[跳过] 不支持的形状: {shape.shape_type} name={shape.name}')
        return canvas

    def _draw_picture(self, canvas, shape):
        img = Image.open(io.BytesIO(shape.image.blob)).convert('RGBA')
        w = max(1, int(round(self.emu2px(shape.width))))
        h = max(1, int(round(self.emu2px(shape.height))))
        img = img.resize((w, h), Image.LANCZOS)
        canvas.alpha_composite(img, (int(round(self.emu2px(shape.left))),
                                     int(round(self.emu2px(shape.top)))))

    # ---------------- 文本 ----------------

    def _body_insets(self, tf):
        bodyPr = tf._txBody.find(qn('a:bodyPr'))
        d = {'lIns': 91440, 'rIns': 91440, 'tIns': 45720, 'bIns': 45720}
        if bodyPr is not None:
            for k in d:
                if bodyPr.get(k) is not None:
                    d[k] = int(bodyPr.get(k))
        return {k: self.emu2px(v) for k, v in d.items()}

    def _draw_text_frame(self, canvas, shape):
        tf = shape.text_frame
        ins = self._body_insets(tf)
        box_l = self.emu2px(shape.left)
        box_t = self.emu2px(shape.top)
        avail = self.emu2px(shape.width) - ins['lIns'] - ins['rIns']

        # 1) 展开段落 -> 带样式的 token 序列
        paragraphs = []  # (algn, lnSpc_pct, [(token, style)])
        for para in tf.paragraphs:
            algn = para.alignment
            algn = {1: 'l', 2: 'ctr', 3: 'r', 4: 'just'}.get(algn if algn is None else int(algn), 'l')
            pPr = para._p.find(qn('a:pPr'))
            lnSpc = 1.0
            if pPr is not None:
                sp = pPr.find(qn('a:lnSpc') + '/' + qn('a:spcPct'))
                if sp is not None:
                    lnSpc = int(sp.get('val')) / 100000.0
            toks = []
            for run in para.runs:
                st = read_style(run, self.theme)
                for tk in tokenize(run.text):
                    toks.append((tk, st))
            paragraphs.append((str(algn), lnSpc, toks))

        # 2) 逐段布局行
        layers = {}   # style.key() -> [(char, font, x_px, baseline_px)]
        y_baseline = box_t + ins['tIns']
        first_line = True
        for algn, lnSpc, toks in paragraphs:
            if not toks:
                continue
            max_sz = max(st.size_pt for _, st in toks)
            line_h = self.pt2px(max_sz) * LINE_HEIGHT_RATIO * lnSpc
            if first_line:
                y_baseline += BASELINE0_FRAC * line_h
                first_line = False
            else:
                y_baseline += line_h  # 上一行基线 + 行高

            # 贪心断行
            lines, cur, cur_w = [], [], 0.0
            for tk, st in toks:
                f = self.font(st)
                adv = sum(f.getlength(ch) + self.pt2px(st.spc_pt) for ch in tk)
                if cur and not tk.isspace() and cur_w + adv > avail:
                    lines.append(cur)
                    cur, cur_w = [], 0.0
                    if tk.isspace():
                        continue
                cur.append((tk, st, adv))
                cur_w += adv
            if cur:
                lines.append(cur)

            for line in lines:
                lw = sum(adv for _, _, adv in line)
                last_st = line[-1][1]
                lw -= self.pt2px(last_st.spc_pt)  # 行尾字距不计入居中宽度
                if algn == 'ctr':
                    x = box_l + ins['lIns'] + (avail - lw) / 2.0
                elif algn == 'r':
                    x = box_l + ins['lIns'] + (avail - lw)
                else:
                    x = box_l + ins['lIns']
                for tk, st, _ in line:
                    f = self.font(st)
                    for ch in tk:
                        if not ch.isspace():
                            layers.setdefault(st.key(), []).append(
                                (ch, st, x, y_baseline))
                        x += f.getlength(ch) + self.pt2px(st.spc_pt)
                y_baseline += line_h

        # 3) 按样式分组绘制（阴影 -> 填充）
        if not layers:
            return
        W, H = canvas.size
        text_layer = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        for _, glyphs in layers.items():
            st = glyphs[0][1]
            mask = Image.new('L', (W, H), 0)
            md = ImageDraw.Draw(mask)
            f = self.font(st)
            for ch, _, x, yb in glyphs:
                md.text((x, yb), ch, font=f, fill=255, anchor='ls')
            bbox = mask.getbbox()
            if not bbox:
                continue
            # 阴影
            if st.shadow:
                s = st.shadow
                rgb, alpha = parse_color_el(s['clr_el'], self.theme)
                sigma = self.pt2px(s['blur_pt']) * SHADOW_SIGMA_FRAC
                sh = mask.filter(ImageFilter.GaussianBlur(sigma))
                dist = self.pt2px(s['dist_pt'])
                th = math.radians(s['dir_deg'])
                dx, dy = dist * math.cos(th), dist * math.sin(th)
                sh_img = Image.new('RGBA', (W, H), (*rgb, 0))
                sh_img.putalpha(sh.point(lambda v, a=int(alpha * 255): v * a // 255))
                text_layer.alpha_composite(sh_img, (int(round(dx)), int(round(dy))))
            # 填充
            if st.fill and st.fill[0] == 'grad':
                _, stops, ang = st.fill
                gx0, gy0, gx1, gy1 = bbox
                grad = make_linear_gradient(gx1 - gx0, gy1 - gy0, stops, ang)
                tile = Image.new('RGBA', (W, H), (0, 0, 0, 0))
                tile.paste(grad, (gx0, gy0))
                tile.putalpha(mask)
                text_layer.alpha_composite(tile)
            else:
                rgb, alpha = (0, 0, 0), 1.0
                if st.fill and st.fill[0] == 'solid':
                    rgb, alpha = st.fill[1], st.fill[2]
                fill_img = Image.new('RGBA', (W, H), (*rgb, 0))
                fill_img.putalpha(mask.point(lambda v, a=int(alpha * 255): v * a // 255))
                text_layer.alpha_composite(fill_img)
        canvas.alpha_composite(text_layer)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pptx', required=True)
    ap.add_argument('--quality', type=float, default=2.0)
    ap.add_argument('--out', default=None)
    args = ap.parse_args()

    r = PillowRenderer(args.pptx, args.quality)
    img = r.render()
    out = args.out or os.path.join(os.path.dirname(__file__), 'out',
                                   f'poc_slide1_q{args.quality:g}.png')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    img.convert('RGB').save(out)
    print(f'渲染完成: {out} ({img.size[0]}x{img.size[1]})')


if __name__ == '__main__':
    main()
