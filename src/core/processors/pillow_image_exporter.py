"""
纯 Python 图片导出器（Pillow 渲染引擎）
不依赖 Microsoft Office / WPS Office，直接将 PPTX 渲染为图片

支持范围（受限模板规范，海报模板家族）：
- 背景图/图片（p:pic，含 srcRect 裁剪以外的常规 stretch 填充）
- 文本框：字体/字号/粗体/斜体/字距(spc)/行距(spcPct)/对齐(l/ctr/r)/顶端锚定
- 文字填充：纯色(srgbClr/schemeClr/sysClr) 或 线性渐变(gradFill + lin ang)
- 文字外阴影(outerShdw: blurRad/dist/dir/颜色/alpha)
- 多页演示文稿、PNG/JPG 输出、quality 缩放倍数（对应 96dpi 的倍数）

已知边界（渲染时打印警告并跳过该元素，不中断导出）：
旋转/翻转、组合形状、非矩形形状几何、表格、图表、SmartArt、
艺术字变形、文字描边、glow/3D、图片裁剪(srcRect)、格式继承
（run 未显式书写格式且来自母版/版式继承时，将退化为默认黑字）

排版度量经 PowerPoint COM 导出对照校准（见 experiments/calibrate.py）：
LINE_HEIGHT_RATIO=1.25 / BASELINE0_FRAC=0.72，残余偏差 ±4.5 逻辑像素。
"""

import logging
import math
import os
import re
import zipfile
from typing import List

from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

logger = logging.getLogger(__name__)

# ---------------- 校准常数（对照 PowerPoint 导出图调整，勿随意改动） ----------------
LINE_HEIGHT_RATIO = 1.25   # 微软雅黑 单倍行距/字号（PowerPoint 按字体 metrics）
BASELINE0_FRAC = 0.72      # 首行基线在行高盒中的位置
SHADOW_SIGMA_FRAC = 0.5    # PowerPoint blurRad -> Pillow GaussianBlur sigma 的系数
# -----------------------------------------------------------------------

EMU_PER_PT = 12700

# 字体注册表：模板 typeface -> Windows 字体文件（按需扩充）
_FONT_FILES = {
    ('微软雅黑', False): r'C:\Windows\Fonts\msyh.ttc',
    ('微软雅黑', True): r'C:\Windows\Fonts\msyhbd.ttc',
    ('Microsoft YaHei', False): r'C:\Windows\Fonts\msyh.ttc',
    ('Microsoft YaHei', True): r'C:\Windows\Fonts\msyhbd.ttc',
    ('宋体', False): r'C:\Windows\Fonts\simsun.ttc',
    ('宋体', True): r'C:\Windows\Fonts\simsun.ttc',
    ('黑体', False): r'C:\Windows\Fonts\simhei.ttf',
    ('黑体', True): r'C:\Windows\Fonts\simhei.ttf',
    ('楷体', False): r'C:\Windows\Fonts\simkai.ttf',
    ('楷体', True): r'C:\Windows\Fonts\simkai.ttf',
}
_DEFAULT_FONT = r'C:\Windows\Fonts\msyh.ttc'


def _hex2rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _lerp(a, b, t):
    return a + (b - a) * t


def _lerp_color(c1, c2, t):
    return tuple(int(round(_lerp(a, b, t))) for a, b in zip(c1, c2))


class _Theme:
    """从 pptx zip 解析 theme 配色，供 schemeClr 映射"""

    def __init__(self, pptx_path: str):
        self.colors = {'dk1': '#000000', 'lt1': '#FFFFFF', 'dk2': '#000000', 'lt2': '#FFFFFF'}
        try:
            with zipfile.ZipFile(pptx_path) as z:
                names = [n for n in z.namelist()
                         if re.match(r'ppt/theme/theme\d+\.xml$', n)]
                if names:
                    xml = z.read(sorted(names)[0]).decode('utf-8', errors='ignore')
                    for name in ('dk1', 'lt1', 'dk2', 'lt2',
                                 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'):
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
        except Exception as e:
            logger.warning(f"解析主题配色失败（使用默认黑白）: {e}")


def _parse_color_el(clr_el, theme):
    """解析颜色元素（srgbClr/schemeClr/sysClr），返回 (rgb_tuple, alpha 0~1)"""
    from lxml import etree
    tag = etree.QName(clr_el).localname
    if tag == 'srgbClr':
        rgb = _hex2rgb(clr_el.get('val'))
    elif tag == 'schemeClr':
        rgb = _hex2rgb(theme.colors.get(clr_el.get('val'), '#000000'))
    elif tag == 'sysClr':
        rgb = _hex2rgb(clr_el.get('lastClr') or '000000')
    else:
        rgb = (0, 0, 0)
    alpha = 1.0
    a = clr_el.find(qn('a:alpha'))
    if a is not None:
        alpha = int(a.get('val')) / 100000.0
    return rgb, alpha


def _parse_run_fill(rPr, theme):
    """返回 None（继承=黑） | ('solid', rgb, alpha) | ('grad', [stops], ang_deg)"""
    if rPr is None:
        return None
    grad = rPr.find(qn('a:gradFill'))
    if grad is not None:
        stops = []
        for gs in grad.findall(qn('a:gsLst') + '/' + qn('a:gs')):
            pos = int(gs.get('pos')) / 100000.0
            clr = list(gs)[0]
            rgb, a = _parse_color_el(clr, theme)
            stops.append((pos, rgb, a))
        lin = grad.find(qn('a:lin'))
        ang = int(lin.get('ang')) / 60000.0 if lin is not None else 0.0
        return ('grad', stops, ang)
    solid = rPr.find(qn('a:solidFill'))
    if solid is not None:
        rgb, a = _parse_color_el(list(solid)[0], theme)
        return ('solid', rgb, a)
    return None


def _parse_shadow(rPr):
    """返回 None | dict(blur_pt, dist_pt, dir_deg, clr_el)"""
    if rPr is None:
        return None
    shdw = rPr.find(qn('a:effectLst') + '/' + qn('a:outerShdw'))
    if shdw is None:
        return None
    return {
        'blur_pt': int(shdw.get('blurRad', '0')) / EMU_PER_PT,
        'dist_pt': int(shdw.get('dist', '0')) / EMU_PER_PT,
        'dir_deg': int(shdw.get('dir', '0')) / 60000.0,
        'clr_el': list(shdw)[0],
    }


def _tokenize(text):
    """CJK 逐字、英文按词、空格独立成 token（换行断点）"""
    return [m.group(0) for m in re.finditer(r'[A-Za-z0-9]+(?:[.\u00b7][A-Za-z0-9]+)*|\s+|.', text)]


class _Style:
    __slots__ = ('size_pt', 'bold', 'italic', 'spc_pt', 'fill', 'shadow', 'font_path')

    def __init__(self, size_pt, bold, italic, spc_pt, fill, shadow, font_path):
        self.size_pt, self.bold, self.italic = size_pt, bold, italic
        self.spc_pt, self.fill, self.shadow, self.font_path = spc_pt, fill, shadow, font_path

    def key(self):
        return (self.size_pt, self.bold, self.italic, self.spc_pt,
                repr(self.fill), repr(self.shadow))


def _read_style(run, theme):
    rPr = run._r.find(qn('a:rPr'))
    size_pt, bold, italic, spc_pt = 18.0, False, False, 0.0
    font_path = None
    if rPr is not None:
        if rPr.get('sz'):
            size_pt = int(rPr.get('sz')) / 100.0
        bold = rPr.get('b') == '1'
        italic = rPr.get('i') == '1'
        if rPr.get('spc'):
            spc_pt = int(rPr.get('spc')) / 100.0
        for tag in ('a:ea', 'a:latin'):
            tf_el = rPr.find(qn(tag))
            if tf_el is not None and tf_el.get('typeface'):
                path = _FONT_FILES.get((tf_el.get('typeface'), bold))
                if path:
                    font_path = path
                    break
                logger.warning(f"字体未注册，回退微软雅黑: {tf_el.get('typeface')!r}")
    if font_path is None:
        font_path = _FONT_FILES.get(('微软雅黑', bold), _DEFAULT_FONT)
    fill = _parse_run_fill(rPr, theme)
    shadow = _parse_shadow(rPr)
    return _Style(size_pt, bold, italic, spc_pt, fill, shadow, font_path)


def _make_linear_gradient(w, h, stops, ang_deg):
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
            t = min(max(((x * vx + y * vy) - p0) / span, 0.0), 1.0)
            for i in range(len(stops) - 1):
                pa, ca, aa = stops[i]
                pb, cb, ab = stops[i + 1]
                if pa <= t <= pb:
                    tt = (t - pa) / (pb - pa) if pb > pa else 0.0
                    px[x, y] = (*_lerp_color(ca, cb, tt), int(round(_lerp(aa, ab, tt) * 255)))
                    break
            else:
                pos, c, a = stops[-1] if t >= stops[-1][0] else stops[0]
                px[x, y] = (*c, int(round(a * 255)))
    return grad


class PillowImageExporter:
    """纯 Python PPTX -> 图片导出器（替代 COMImageExporter，无需 Office）"""

    def __init__(self):
        self.font_cache = {}
        self._dpi = 96.0

    @staticmethod
    def is_available() -> bool:
        """纯 Python 引擎始终可用"""
        return True

    @staticmethod
    def get_engine_info() -> str:
        return "纯 Python 渲染引擎（Pillow，无需 Office/WPS）"

    def export_to_images(self, pptx_path: str, output_dir: str,
                         image_format: str = "PNG", quality: float = 1.0) -> List[str]:
        """
        将 PPTX 渲染为图片

        Args:
            pptx_path: PPTX 文件路径
            output_dir: 输出目录
            image_format: 图片格式 (PNG, JPG)
            quality: 缩放倍数 (1.0=96dpi 基准, 2.0/3.0=增强)

        Returns:
            List[str]: 生成的图片路径列表（slide_1.ext, slide_2.ext, ...）
        """
        if not os.path.exists(pptx_path):
            raise FileNotFoundError(f"文件不存在：{pptx_path}")
        os.makedirs(output_dir, exist_ok=True)

        prs = Presentation(pptx_path)
        theme = _Theme(pptx_path)
        ext = 'jpg' if image_format.upper() in ('JPG', 'JPEG') else image_format.lower()
        self._dpi = 96.0 * quality

        generated = []
        for idx, slide in enumerate(prs.slides, 1):
            img = self._render_slide(prs, slide, theme, quality)
            out_path = os.path.join(output_dir, f"slide_{idx}.{ext}")
            if ext == 'jpg':
                img.convert('RGB').save(out_path, 'JPEG', quality=95)
            else:
                img.convert('RGB').save(out_path)
            logger.info(f"渲染图片: {out_path} ({img.size[0]}x{img.size[1]})")
            generated.append(out_path)
        return generated

    # ---------------- 渲染核心 ----------------

    def _render_slide(self, prs, slide, theme, quality):
        scale = 96.0 * quality / 914400.0
        W = int(round(prs.slide_width * scale))
        H = int(round(prs.slide_height * scale))
        canvas = Image.new('RGBA', (W, H), (255, 255, 255, 255))
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                self._draw_picture(canvas, shape, scale)
            elif shape.has_text_frame:
                self._draw_text_frame(canvas, shape, theme, quality, scale)
            else:
                logger.warning(
                    f"[纯Python渲染] 跳过不支持的形状: type={shape.shape_type} "
                    f"name={shape.name!r}（该元素不会出现在输出图片中）")
        return canvas

    def _draw_picture(self, canvas, shape, scale):
        import io
        img = Image.open(io.BytesIO(shape.image.blob)).convert('RGBA')
        w = max(1, int(round(shape.width * scale)))
        h = max(1, int(round(shape.height * scale)))
        img = img.resize((w, h), Image.LANCZOS)
        canvas.alpha_composite(img, (int(round(shape.left * scale)),
                                     int(round(shape.top * scale))))

    def _draw_text_frame(self, canvas, shape, theme, quality, scale):
        dpi = 96.0 * quality

        def pt2px(v):
            return v * dpi / 72.0

        tf = shape.text_frame
        bodyPr = tf._txBody.find(qn('a:bodyPr'))
        ins = {'lIns': 91440, 'rIns': 91440, 'tIns': 45720, 'bIns': 45720}
        for k in ins:
            if bodyPr is not None and bodyPr.get(k) is not None:
                ins[k] = int(bodyPr.get(k))
        ins = {k: v * scale for k, v in ins.items()}

        box_l = shape.left * scale
        box_t = shape.top * scale
        avail = shape.width * scale - ins['lIns'] - ins['rIns']

        # 1) 展开段落 -> 带样式的 token 序列
        paragraphs = []  # (algn, lnSpc_pct, [(token, style)])
        for para in tf.paragraphs:
            algn = para.alignment
            algn = {1: 'l', 2: 'ctr', 3: 'r'}.get(algn if algn is None else int(algn), 'l')
            pPr = para._p.find(qn('a:pPr'))
            lnSpc = 1.0
            if pPr is not None:
                sp = pPr.find(qn('a:lnSpc') + '/' + qn('a:spcPct'))
                if sp is not None:
                    lnSpc = int(sp.get('val')) / 100000.0
            toks = []
            for run in para.runs:
                st = _read_style(run, theme)
                if st.fill is None:
                    logger.warning(
                        f"[纯Python渲染] 形状 {shape.name!r} 的文字未显式指定颜色/渐变，"
                        f"若模板依赖母版/版式继承格式，输出可能与 PowerPoint 不一致")
                for tk in _tokenize(run.text):
                    toks.append((tk, st))
            paragraphs.append((str(algn), lnSpc, toks))

        # 2) 逐段布局行（与 PoC 校准版一致）
        layers = {}   # style.key() -> [(char, style, x_px, baseline_px)]
        y_baseline = box_t + ins['tIns']
        first_line = True
        for algn, lnSpc, toks in paragraphs:
            if not toks:
                continue
            max_sz = max(st.size_pt for _, st in toks)
            line_h = pt2px(max_sz) * LINE_HEIGHT_RATIO * lnSpc
            if first_line:
                y_baseline += BASELINE0_FRAC * line_h
                first_line = False
            else:
                y_baseline += line_h

            lines, cur, cur_w = [], [], 0.0
            for tk, st in toks:
                f = self._font(st)
                adv = sum(f.getlength(ch) + pt2px(st.spc_pt) for ch in tk)
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
                lw -= pt2px(line[-1][1].spc_pt)  # 行尾字距不计入居中宽度
                if algn == 'ctr':
                    x = box_l + ins['lIns'] + (avail - lw) / 2.0
                elif algn == 'r':
                    x = box_l + ins['lIns'] + (avail - lw)
                else:
                    x = box_l + ins['lIns']
                for tk, st, _ in line:
                    f = self._font(st)
                    for ch in tk:
                        if not ch.isspace():
                            layers.setdefault(st.key(), []).append((ch, st, x, y_baseline))
                        x += f.getlength(ch) + pt2px(st.spc_pt)
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
            f = self._font(st)
            for ch, _, x, yb in glyphs:
                md.text((x, yb), ch, font=f, fill=255, anchor='ls')
            bbox = mask.getbbox()
            if not bbox:
                continue
            if st.shadow:
                s = st.shadow
                rgb, alpha = _parse_color_el(s['clr_el'], theme)
                sigma = pt2px(s['blur_pt']) * SHADOW_SIGMA_FRAC
                sh = mask.filter(ImageFilter.GaussianBlur(sigma))
                dist = pt2px(s['dist_pt'])
                th = math.radians(s['dir_deg'])
                sh_img = Image.new('RGBA', (W, H), (*rgb, 0))
                sh_img.putalpha(sh.point(lambda v, a=int(alpha * 255): v * a // 255))
                text_layer.alpha_composite(
                    sh_img,
                    (int(round(dist * math.cos(th))), int(round(dist * math.sin(th))))
                )
            if st.fill and st.fill[0] == 'grad':
                _, stops, ang = st.fill
                gx0, gy0, gx1, gy1 = bbox
                grad = _make_linear_gradient(gx1 - gx0, gy1 - gy0, stops, ang)
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

    def _font(self, style):
        key = (style.font_path, style.size_pt, style.bold, style.italic)
        if key not in self.font_cache:
            size_px = style.size_pt * self._dpi / 72.0
            self.font_cache[key] = ImageFont.truetype(style.font_path, size_px)
        return self.font_cache[key]
