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
import io
import math
import os
import re
import tempfile
from typing import List

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

try:
    from src.render_core.format_inheritance import (
        FormatResolver, load_theme, parse_color_el, parse_fill, resolve_bg_ref,
        grad_stops,
    )
    from src.render_core.shape_geometry import preset_geometry, draw_path
    from src.render_core import font_registry
except ImportError:
    from render_core.format_inheritance import (
        FormatResolver, load_theme, parse_color_el, parse_fill, resolve_bg_ref,
        grad_stops,
    )
    from render_core.shape_geometry import preset_geometry, draw_path
    from render_core import font_registry

logger = logging.getLogger(__name__)

# ---------------- 校准常数（对照 PowerPoint 导出图调整，勿随意改动） ----------------
# 行距/基线/删除线三常数已按物理字体拆分到 font_registry（metrics_for 按字体文件查询），
# 此处保留全局兜底值（= 雅黑 COM 实测值）；阴影/斜体等其余仍为全局值
LINE_HEIGHT_RATIO = 1.25   # 微软雅黑 单倍行距/字号（PowerPoint 按字体 metrics）
BASELINE0_FRAC = 0.72      # 首行基线在行高盒中的位置
SHADOW_SIGMA_FRAC = 0.5    # PowerPoint blurRad -> Pillow GaussianBlur sigma 的系数
# glow 三常数（红/橙/黄/绿/蓝五色 COM 探针统一校准，与色相无关；
# 强度比全色 ±4%，光晕面积比 ≤+19%——换任意发光颜色无需重调）
GLOW_SIGMA_KS = (0.08, 0.18, 0.3)   # glow 各层 sigma = rad × k
GLOW_GAIN = 2.45             # glow alpha 增益（乘发光色 alpha 后封顶 255）
GLOW_GAMMA = 0.7             # glow alpha 伽马（<1 提亮：光晕边缘更浓）


def _glow_lut(alpha_255):
    """glow 蒙版值 -> 最终 alpha 的 256 项 LUT：先伽马再乘增益封顶"""
    top = min(255, int(alpha_255 * GLOW_GAIN))
    return [int((v / 255.0) ** GLOW_GAMMA * top) for v in range(256)]
ITALIC_SHEAR = 0.2126      # 合成斜体斜切率（tan 12°，无斜体字库时使用）
STRIKE_POS_EM = 0.245      # 删除线位于基线上方 0.245em（COM 对照微软雅黑实测校准）
STRIKE_THICK_EM = 0.05     # 删除线/下划线厚度 0.05em
UNDERLINE_POS_EM = 0.09    # 下划线位于基线下方 0.09em
# -----------------------------------------------------------------------

EMU_PER_PT = 12700

# 字体解析统一走 font_registry（捆绑思源黑体/宋体 + 系统 Windows 字体表，跨平台）
# 排版校准常数随物理字体绑定，见 font_registry.METRICS_*


def _load_rgba(blob, target_wh, crop_rect=None):
    """图片字节 -> 目标尺寸 RGBA（移动端 OOM 防护的内存路径优化）：
    - 大 blob（>8MB）写临时文件解码，避免 io.BytesIO 的整份字节拷贝
    - JPEG 且无需裁剪：draft 解码期降采样（DCT 缩放，跳过全分辨率解码）
    - 不透明图先 resize 再补 alpha，避免全分辨率 RGBA 中间副本
      （LANCZOS 按通道独立重采样，与"先 convert 再 resize"逐像素等价）
    - crop_rect (l,t,r,b 源图百分比)：srcRect 裁剪在原生模式上做，再 resize
    """
    tw = (max(1, int(target_wh[0])), max(1, int(target_wh[1])))
    tmp_path = None
    try:
        if len(blob) > 8 * 1024 * 1024:
            fd, tmp_path = tempfile.mkstemp(suffix='.imgblob')
            with os.fdopen(fd, 'wb') as f:
                f.write(blob)
            img = Image.open(tmp_path)
        else:
            img = Image.open(io.BytesIO(blob))
        try:
            has_alpha = (img.mode in ('RGBA', 'LA')
                         or (img.mode == 'P' and 'transparency' in img.info))
            if img.format == 'JPEG' and not crop_rect:
                img.draft('RGB', tw)
            if crop_rect:
                w0, h0 = img.size
                l, t, r, b = crop_rect
                img = img.crop((int(w0 * l), int(h0 * t),
                                int(w0 * (1 - r)), int(h0 * (1 - b))))
            if has_alpha:
                return img.convert('RGBA').resize(tw, Image.LANCZOS)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            return img.resize(tw, Image.LANCZOS).convert('RGBA')
        finally:
            img.close()   # 释放文件句柄/解码缓冲（派生图已复制，不受影响）
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except OSError:
                pass


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


def _parse_glow(rPr):
    """文字发光 a:glow：返回 None | dict(rad_pt, clr_el)。颜色带 alpha（如 40%）"""
    if rPr is None:
        return None
    glow = rPr.find(qn('a:effectLst') + '/' + qn('a:glow'))
    if glow is None:
        return None
    clr = list(glow)
    if not clr:
        return None
    return {'rad_pt': int(glow.get('rad', '0')) / EMU_PER_PT, 'clr_el': clr[0]}


def _parse_stroke(rPr, theme_colors):
    """解析文字描边 rPr/a:ln，返回 None | (width_pt, rgb, alpha)"""
    if rPr is None:
        return None
    ln = rPr.find(qn('a:ln'))
    if ln is None:
        return None
    solid = ln.find(qn('a:solidFill'))
    if solid is None:
        return None  # noFill 或渐变描边：不支持渐变描边
    rgb, alpha = parse_color_el(list(solid)[0], theme_colors)
    w_pt = int(ln.get('w', '9525')) / EMU_PER_PT
    return (w_pt, rgb, alpha)


def _tokenize(text):
    """CJK 逐字、英文按词、空格独立成 token（换行断点）"""
    return [m.group(0) for m in re.finditer(r'[A-Za-z0-9]+(?:[.\u00b7][A-Za-z0-9]+)*|\s+|.', text)]


class _Style:
    __slots__ = ('size_pt', 'bold', 'italic', 'spc_pt', 'fill', 'shadow',
                 'glow', 'stroke', 'strike', 'underline', 'font_path')

    def __init__(self, size_pt, bold, italic, spc_pt, fill, shadow, glow,
                 stroke, strike, underline, font_path):
        self.size_pt, self.bold, self.italic = size_pt, bold, italic
        self.spc_pt, self.fill, self.shadow = spc_pt, fill, shadow
        self.glow, self.stroke = glow, stroke
        self.strike, self.underline = strike, underline
        self.font_path = font_path

    def key(self):
        return (self.size_pt, self.bold, self.italic, self.spc_pt,
                repr(self.fill), repr(self.shadow), repr(self.glow),
                repr(self.stroke), self.strike, self.underline)


def _font_path_for(typeface, bold):
    """字体名 -> 字体文件；走字体注册表（策略见 font_registry），无命中回退默认并告警"""
    return font_registry.resolve(typeface, bold).path


def _style_from_props(run, props, theme_colors):
    """把继承解析出的有效属性转换为渲染样式（阴影/描边只取显式 rPr）"""
    rPr = run._r.find(qn('a:rPr'))
    shadow = _parse_shadow(rPr)
    glow = _parse_glow(rPr)
    stroke = _parse_stroke(rPr, theme_colors)
    return _Style(props['sz_pt'], props['bold'], props['italic'], props['spc_pt'],
                  props['fill'], shadow, glow, stroke,
                  props.get('strike', False), props.get('underline'),
                  _font_path_for(props['typeface'], props['bold']))


def _make_linear_gradient(w, h, stops, ang_deg, ref_w=None, ref_h=None,
                          off_x=0.0, off_y=0.0):
    """按 stops [(pos, rgb, alpha)] 生成 w×h RGBA 线性渐变图。ang 0°=左→右，90°=上→下。
    有 numpy 走向量化（快）；无 numpy（如安卓 Chaquopy 环境）自动退化纯 Pillow 路径。
    ref_w/ref_h：t 归一化基准尺寸（默认 = w,h）；off_x/off_y：投影坐标偏移。
    文字渐变按文本框 box 归一化（PowerPoint 语义），生成范围可大于框以覆盖溢出
    文字——框外 t 钳制到端点色，与 PowerPoint 溢出行为一致（不是黑色）。"""
    theta = math.radians(ang_deg)
    vx, vy = math.cos(theta), math.sin(theta)
    rw, rh = (ref_w or w), (ref_h or h)
    corners = [(0, 0), (rw - 1, 0), (0, rh - 1), (rw - 1, rh - 1)]
    projs = [x * vx + y * vy for x, y in corners]
    p0, p1 = min(projs), max(projs)
    span = (p1 - p0) or 1.0
    try:
        import numpy as np
    except ImportError:
        return _gradient_pure(w, h, stops, vx, vy, p0, span, off_x, off_y)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    t = np.clip((((xx + off_x) * vx + (yy + off_y) * vy) - p0) / span, 0.0, 1.0)
    pos = np.array([s[0] for s in stops], dtype=np.float32)
    channels = [
        np.interp(t, pos, np.array([s[1][c] for s in stops], dtype=np.float32))
        for c in range(3)
    ]
    channels.append(np.interp(
        t, pos, np.array([s[2] * 255 for s in stops], dtype=np.float32)))
    arr = np.dstack(channels).round().astype(np.uint8)
    return Image.fromarray(arr, 'RGBA')


def _gradient_pure(w, h, stops, vx, vy, p0, span, off_x=0.0, off_y=0.0):
    """无 numpy 的渐变路径（安卓零依赖）：按行线性投影量化 t，再用 bytes.translate
    做 C 级逐通道 LUT 映射（与 numpy 路径同数学，t 多一层 8bit 量化，差异 ≤2/255）。
    注意：整幅画布级大渐变（如 3584 宽）每张约 1~2s；海报场景渐变多为文字级小区域。"""
    def build_lut(vals):
        out = bytearray(256)
        pos = [s[0] for s in stops]
        n = len(pos)
        for i in range(256):
            t = min(1.0, max(0.0, i / 255.0))
            j = 0
            while j < n - 2 and t > pos[j + 1]:
                j += 1
            t0, t1 = pos[j], pos[j + 1]
            f = (t - t0) / (t1 - t0) if t1 > t0 else 0.0
            out[i] = int(round(vals[j] + (vals[j + 1] - vals[j]) * f)) & 0xFF
        return bytes(out)

    lut_r = build_lut([s[1][0] for s in stops])
    lut_g = build_lut([s[1][1] for s in stops])
    lut_b = build_lut([s[1][2] for s in stops])
    lut_a = build_lut([int(round(s[2] * 255)) for s in stops])

    base0 = (off_x * vx + off_y * vy - p0) / span
    dx, dy = vx / span, vy / span
    rows = [[] for _ in range(4)]
    for y in range(h):
        base = (base0 + dy * y) * 255.0
        dxx = dx * 255.0
        tq = bytes(min(255, max(0, int(base + dxx * x + 0.5))) for x in range(w))
        rows[0].append(tq.translate(lut_r))
        rows[1].append(tq.translate(lut_g))
        rows[2].append(tq.translate(lut_b))
        rows[3].append(tq.translate(lut_a))
    chans = [Image.frombytes('L', (w, h), b''.join(r)) for r in rows]
    return Image.merge('RGBA', chans)


class _Xform:
    """像素坐标变换（组合形状嵌套）：canvas = off + p * scale"""

    __slots__ = ('ox', 'oy', 'sx', 'sy')

    def __init__(self, ox=0.0, oy=0.0, sx=1.0, sy=1.0):
        self.ox, self.oy, self.sx, self.sy = ox, oy, sx, sy

    def mx(self, v):
        return self.ox + v * self.sx

    def my(self, v):
        return self.oy + v * self.sy

    @property
    def s_font(self):
        """文字统一缩放因子（非等比组合下取均值近似）"""
        return (self.sx + self.sy) / 2.0

    def child(self, off_px, ext_px, ch_off_px, ch_ext_px):
        """组合子坐标 -> 画布坐标的复合变换"""
        kx = (ext_px[0] / ch_ext_px[0]) if ch_ext_px[0] else 1.0
        ky = (ext_px[1] / ch_ext_px[1]) if ch_ext_px[1] else 1.0
        nsx, nsy = self.sx * kx, self.sy * ky
        nox = self.mx(off_px[0]) - nsx * ch_off_px[0]
        noy = self.my(off_px[1]) - nsy * ch_off_px[1]
        return _Xform(nox, noy, nsx, nsy)


_IDENTITY = _Xform()


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

    def precheck(self, pptx_path: str):
        """渲染预检（白名单制），返回问题清单；字体检查自动接入 FontChecker"""
        try:
            from src.render_core.render_precheck import RenderPrecheck
        except ImportError:
            from render_core.render_precheck import RenderPrecheck
        installed = font_registry.available_typefaces()
        try:
            try:
                from src.core.utils.font_checker import FontChecker
            except ImportError:
                from core.utils.font_checker import FontChecker
            installed |= FontChecker().get_installed_fonts()
        except Exception as e:
            logger.debug(f"系统字体检查跳过（仅用注册表字体集）: {e}")
        return RenderPrecheck(installed_fonts=installed).check_pptx(pptx_path)

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
        font_registry.load_embedded_fonts(pptx_path)

        # 渲染预检（白名单制）：本分支无 COM 兜底，问题仅告警后尽力渲染
        issues = self.precheck(pptx_path)
        for issue in issues:
            logger.warning(f"[预检] 第{issue.slide}页 {issue.shape!r} [{issue.code}] {issue.message}")
        if issues:
            logger.warning(f"[预检] 共 {len(issues)} 项不支持/风险提示，输出可能与 PowerPoint 存在差异")

        prs = Presentation(pptx_path)
        return self.export_presentation(prs, output_dir, image_format, quality,
                                        theme=load_theme(pptx_path))

    def export_presentation(self, prs, output_dir: str, image_format: str = "PNG",
                            quality: float = 1.0, theme=None) -> List[str]:
        """
        从已打开的 Presentation 渲染（批量个性化通道：模板/主题只加载一次，
        配合 personalize.Personalizer 的"替换->渲染->恢复"逐行复用）。

        Args:
            prs: 已打开的 python-pptx Presentation
            theme: 预加载的主题（None 时无法回源加载，仅当模板无主题依赖时可用）
        """
        os.makedirs(output_dir, exist_ok=True)
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

        layout = slide.slide_layout
        master = layout.slide_master

        # 背景继承链：slide -> layout -> master（取第一个 bgPr）
        for container in (slide, layout, master):
            if self._draw_slide_background(canvas, container, theme, W, H):
                break

        # z 序：母版静态形状 -> 版式静态形状 -> 幻灯片形状
        # （版式/母版的占位符是继承载体不渲染，PowerPoint 亦如此）
        for container in (master, layout):
            resolver = FormatResolver(container, theme.colors, theme.fonts)
            for shape in container.shapes:
                if self._is_placeholder(shape):
                    continue
                self._draw_shape(canvas, shape, resolver, theme.colors, quality, scale)

        resolver = FormatResolver(slide, theme.colors, theme.fonts)
        for shape in slide.shapes:
            self._draw_shape(canvas, shape, resolver, theme.colors, quality, scale)
        return canvas

    @staticmethod
    def _is_placeholder(shape):
        from pptx.oxml.ns import qn as _qn
        nvpr = shape._element.find(
            _qn('p:nvSpPr') + '/' + _qn('p:nvPr'))
        if nvpr is None:
            nvpr = shape._element.find(_qn('p:nvPicPr') + '/' + _qn('p:nvPr'))
        return nvpr is not None and nvpr.find(_qn('p:ph')) is not None

    def _draw_shape(self, canvas, shape, resolver, theme_colors, quality, scale,
                    xform=None, rot_deg=0.0, flip_h=False, flip_v=False):
        """渲染单个形状。xform：组合坐标变换；rot/flip：本形状的旋转翻转。

        旋转实现：先在独立图层上正常渲染形状（平移到画布原点），
        再绕形状中心旋转/翻转后贴回原位。嵌套组合的旋转只作用于组合整体
        （与 PowerPoint 行为一致：子形状坐标在组合的未旋转坐标系中）。
        """
        xform = xform or _IDENTITY
        tag = shape._element.tag.split('}')[-1]

        # 本层旋转/翻转信息
        my_rot, my_fh, my_fv = rot_deg, flip_h, flip_v
        spPr = shape._element.find(qn('p:spPr'))
        if tag != 'graphicFrame' and spPr is not None:
            xfrm = spPr.find(qn('a:xfrm'))
            if xfrm is not None:
                if xfrm.get('rot'):
                    my_rot = my_rot + int(xfrm.get('rot')) / 60000.0
                my_fh = my_fh or xfrm.get('flipH') == '1'
                my_fv = my_fv or xfrm.get('flipV') == '1'

        need_transform = bool(my_rot or my_fh or my_fv)
        target = canvas
        transform_ctx = None
        if need_transform:
            # 形状在独立图层上按原位渲染，稍后绕自身中心旋转/翻转后贴回
            target = Image.new('RGBA', canvas.size, (0, 0, 0, 0))
            transform_ctx = (
                my_rot, my_fh, my_fv,
                xform.mx(shape.left * scale), xform.my(shape.top * scale),
                xform.mx((shape.left + shape.width) * scale),
                xform.my((shape.top + shape.height) * scale),
            )

        if tag == 'grpSp':
            try:
                self._draw_group(target, shape, resolver, theme_colors,
                                 quality, scale, xform)
            except Exception as e:
                logger.warning(f"[纯Python渲染] 组合渲染失败 name={shape.name!r}: {e}")
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            self._draw_picture(target, shape, scale, xform)
        elif tag == 'sp':
            try:
                self._draw_autoshape(target, shape, theme_colors, quality,
                                     scale, xform)
            except Exception as e:
                logger.warning(
                    f"[纯Python渲染] 形状外观渲染失败 name={shape.name!r}: {e}")
            if shape.has_text_frame:
                self._draw_text_frame(target, shape, resolver, theme_colors,
                                      quality, scale, xform)
        elif shape.has_text_frame:
            self._draw_text_frame(target, shape, resolver, theme_colors,
                                  quality, scale, xform)
        else:
            logger.warning(
                f"[纯Python渲染] 跳过不支持的形状: type={shape.shape_type} "
                f"name={shape.name!r}（该元素不会出现在输出图片中）")
            return

        if transform_ctx:
            r, fh, fv, x0, y0, x1, y1 = transform_ctx
            layer = target
            W, H = layer.size
            # 先翻转（PowerPoint 顺序：flip 后 rot）：裁剪盒必须关于形状中心
            # 对称，镜像后才不会移位（画布边缘裁剪属正常行为）
            if fh or fv:
                m = 80  # 覆盖阴影/描边溢出的边距
                ccx, ccy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
                hw = max((x1 - x0) / 2.0, m)
                hh = max((y1 - y0) / 2.0, m)
                box = (max(0, int(ccx - hw)), max(0, int(ccy - hh)),
                       min(W, int(ccx + hw)), min(H, int(ccy + hh)))
                if box[2] > box[0] and box[3] > box[1]:
                    crop = layer.crop(box).transpose(
                        Image.FLIP_LEFT_RIGHT if fh else Image.FLIP_TOP_BOTTOM)
                    layer.paste(crop, (box[0], box[1]))
            if r:
                cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
                layer = layer.rotate(-r, resample=Image.BICUBIC,
                                     center=(cx, cy), fillcolor=(0, 0, 0, 0))
            canvas.alpha_composite(layer)

    def _draw_group(self, target, group, resolver, theme_colors,
                    quality, scale, xform):
        """递归渲染组合。子形状坐标系：off + (child - chOff) * (ext/chExt)"""
        spPr = group._element.find(qn('p:grpSpPr'))
        xfrm = spPr.find(qn('a:xfrm')) if spPr is not None else None
        if xfrm is None:
            return
        off = xfrm.find(qn('a:off'))
        ext = xfrm.find(qn('a:ext'))
        chOff = xfrm.find(qn('a:chOff'))
        chExt = xfrm.find(qn('a:chExt'))
        if off is None or ext is None:
            return
        off_px = (int(off.get('x')) * scale, int(off.get('y')) * scale)
        ext_px = (int(ext.get('cx')) * scale, int(ext.get('cy')) * scale)
        ch_off = ((int(chOff.get('x')) * scale, int(chOff.get('y')) * scale)
                  if chOff is not None else off_px)
        ch_ext = ((int(chExt.get('cx')) * scale, int(chExt.get('cy')) * scale)
                  if chExt is not None else ext_px)
        child_xf = xform.child(off_px, ext_px, ch_off, ch_ext)

        text_scale = child_xf.s_font
        if abs(text_scale - 1.0) > 0.02:
            logger.info(
                f"[纯Python渲染] 组合 {group.name!r} 带缩放({child_xf.sx:.2f}x"
                f"{child_xf.sy:.2f})，其中文字按 {text_scale:.2f} 缩放渲染")

        for sub in group.shapes:
            self._draw_shape(target, sub, resolver, theme_colors, quality,
                             scale, xform=child_xf)

    def _draw_autoshape(self, canvas, shape, theme_colors, quality, scale,
                        xform=None):
        """渲染 sp 形状外观：预设几何 + 填充(纯色/线性渐变/图片) + 外阴影 + 纯色描边"""
        import io
        xform = xform or _IDENTITY
        spPr = shape._element.spPr
        prstGeom = spPr.find(qn('a:prstGeom'))
        if spPr.find(qn('a:custGeom')) is not None:
            return  # 不支持自定义几何（预检已告警）
        prst = prstGeom.get('prst') if prstGeom is not None else None
        avLst = prstGeom.find(qn('a:avLst')) if prstGeom is not None else None
        x = xform.mx(shape.left * scale)
        y = xform.my(shape.top * scale)
        w = shape.width * scale * xform.sx
        h = shape.height * scale * xform.sy
        path = preset_geometry(prst, x, y, w, h, avLst)
        if path is None:
            return

        dpi = 96.0 * quality

        def pt2px(v):
            return v * dpi / 72.0

        fill = parse_fill(spPr, theme_colors)
        ln = spPr.find(qn('a:ln'))
        line_spec = None
        if ln is not None:
            solid = ln.find(qn('a:solidFill'))
            dash_el = ln.find(qn('a:prstDash'))
            dash = dash_el.get('val', 'solid') if dash_el is not None else 'solid'
            if solid is not None and dash == 'solid':
                w_pt = int(ln.get('w', '9525')) / EMU_PER_PT
                rgb, alpha = parse_color_el(list(solid)[0], theme_colors)
                line_spec = (max(1, int(round(pt2px(w_pt)))), rgb, alpha)

        shadow = _parse_shadow(spPr)
        if fill is None and line_spec is None and shadow is None:
            return  # 纯文本框：无外观可画

        W, H = canvas.size
        layer = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        mask = Image.new('L', (W, H), 0)
        draw_path(ImageDraw.Draw(mask), path, fill=255)

        # 1) 外阴影（垫底）
        if shadow:
            rgb, alpha = parse_color_el(shadow['clr_el'], theme_colors)
            sigma = pt2px(shadow['blur_pt']) * SHADOW_SIGMA_FRAC
            sh = mask.filter(ImageFilter.GaussianBlur(sigma))
            dist = pt2px(shadow['dist_pt'])
            th = math.radians(shadow['dir_deg'])
            sh_img = Image.new('RGBA', (W, H), (*rgb, 0))
            sh_img.putalpha(sh.point(lambda v, a=int(alpha * 255): v * a // 255))
            layer.alpha_composite(
                sh_img, (int(round(dist * math.cos(th))),
                         int(round(dist * math.sin(th)))))

        # 2) 填充
        if fill:
            if fill[0] == 'solid':
                _, rgb, alpha = fill
                fill_img = Image.new('RGBA', (W, H), (*rgb, 0))
                fill_img.putalpha(mask.point(
                    lambda v, a=int(alpha * 255): v * a // 255))
                layer.alpha_composite(fill_img)
            elif fill[0] == 'grad':
                _, stops, ang = fill
                grad = _make_linear_gradient(max(1, int(round(w))),
                                             max(1, int(round(h))), stops, ang)
                tile = Image.new('RGBA', (W, H), (0, 0, 0, 0))
                tile.paste(grad, (int(round(x)), int(round(y))))
                tile.putalpha(mask)
                layer.alpha_composite(tile)
            elif fill[0] == 'blip':
                try:
                    _, rId, src = fill
                    part = shape.part.related_part(rId)
                    rect = ({k: int(src.get(k, '0')) / 100000.0
                             for k in ('l', 't', 'r', 'b')}
                            if src is not None else None)
                    img = _load_rgba(
                        part.blob, (int(round(w)), int(round(h))),
                        crop_rect=(rect.get('l', 0), rect.get('t', 0),
                                   rect.get('r', 0), rect.get('b', 0))
                        if rect else None)
                    tile = Image.new('RGBA', (W, H), (0, 0, 0, 0))
                    tile.paste(img, (int(round(x)), int(round(y))))
                    tile.putalpha(mask)
                    layer.alpha_composite(tile)
                except Exception as e:
                    logger.warning(f"[纯Python渲染] 形状图片填充失败: {e}")

        # 3) 描边（填充之上）
        if line_spec:
            lw, rgb, alpha = line_spec
            line_mask = Image.new('L', (W, H), 0)
            draw_path(ImageDraw.Draw(line_mask), path, outline=255, width=lw)
            line_img = Image.new('RGBA', (W, H), (*rgb, 0))
            line_img.putalpha(line_mask.point(
                lambda v, a=int(alpha * 255): v * a // 255))
            layer.alpha_composite(line_img)

        canvas.alpha_composite(layer)

    def _draw_slide_background(self, canvas, container, theme, W, H) -> bool:
        """渲染 p:bg 背景（bgPr 显式填充或 bgRef 主题样式引用）。返回是否已渲染"""
        bg = container._element.find(qn('p:cSld') + '/' + qn('p:bg'))
        if bg is None:
            return False
        theme_colors = theme.colors
        bgRef = bg.find(qn('p:bgRef'))
        if bgRef is not None:
            try:
                fill = resolve_bg_ref(bgRef, theme)
                if fill is None:
                    return False
                if fill[0] == 'solid':
                    _, rgb, alpha = fill
                    canvas.alpha_composite(Image.new(
                        'RGBA', (W, H), (*rgb, int(alpha * 255))))
                elif fill[0] == 'grad':
                    _, stops, ang = fill
                    canvas.alpha_composite(_make_linear_gradient(W, H, stops, ang))
                return True
            except Exception as e:
                logger.warning(f"[纯Python渲染] bgRef 背景渲染失败: {e}")
                return False
        bgPr = bg.find(qn('p:bgPr'))
        if bgPr is None:
            return False
        try:
            solid = bgPr.find(qn('a:solidFill'))
            if solid is not None:
                rgb, alpha = parse_color_el(list(solid)[0], theme_colors)
                color_img = Image.new('RGBA', (W, H), (*rgb, int(alpha * 255)))
                canvas.alpha_composite(color_img)
                return True
            grad = bgPr.find(qn('a:gradFill'))
            if grad is not None:
                lin = grad.find(qn('a:lin'))
                if lin is not None:
                    stops = grad_stops(
                        grad, lambda clr: parse_color_el(clr, theme_colors))
                    canvas.alpha_composite(
                        _make_linear_gradient(W, H, stops, int(lin.get('ang')) / 60000.0))
                    return True
                return False
            blipFill = bgPr.find(qn('a:blipFill'))
            if blipFill is not None:
                blip = blipFill.find(qn('a:blip'))
                if blip is not None:
                    rId = blip.get(qn('r:embed'))
                    part = container.part.related_part(rId)
                    canvas.alpha_composite(_load_rgba(part.blob, (W, H)))
                    return True
            return False
        except Exception as e:
            logger.warning(f"[纯Python渲染] 幻灯片背景渲染失败（保持白底）: {e}")
            return False

    def _draw_picture(self, canvas, shape, scale, xform=None):
        import io
        xform = xform or _IDENTITY
        w = max(1, int(round(shape.width * scale * xform.sx)))
        h = max(1, int(round(shape.height * scale * xform.sy)))
        # srcRect 裁剪：l/t/r/b 为源图百分比（1/100000 单位）
        crop = None
        blipFill = shape._element.find(qn('p:blipFill'))
        if blipFill is not None:
            src = blipFill.find(qn('a:srcRect'))
            if src is not None:
                crop = tuple(int(src.get(k, '0')) / 100000.0
                             for k in ('l', 't', 'r', 'b'))
        img = _load_rgba(shape.image.blob, (w, h), crop_rect=crop)
        canvas.alpha_composite(img, (int(round(xform.mx(shape.left * scale))),
                                     int(round(xform.my(shape.top * scale)))))

    def _draw_text_frame(self, canvas, shape, resolver, theme_colors, quality, scale,
                         xform=None):
        import io
        xform = xform or _IDENTITY
        dpi = 96.0 * quality

        def pt2px(v):
            return v * dpi / 72.0

        tf = shape.text_frame
        bp = resolver.body_props(shape)
        ins = {k: bp[k] * scale * (xform.sx if k in ('lIns', 'rIns') else xform.sy)
               for k in ('lIns', 'rIns', 'tIns', 'bIns')}

        box_l = xform.mx(shape.left * scale)
        box_t = xform.my(shape.top * scale)
        avail = shape.width * scale * xform.sx - ins['lIns'] - ins['rIns']
        box_h_px = shape.height * scale * xform.sy

        # 1) 展开段落 -> 带样式的 token 序列（格式经继承链解析，应用 normAutofit 缩放）
        # 空段落无行可画但占一行高（PowerPoint 行为），行高取整框最大字号
        font_scale = bp.get('font_scale', 1.0) * xform.s_font
        paras_data = []  # [(algn, lnSpc_pct, toks)]
        frame_max_sz = 18.0
        for para in tf.paragraphs:
            pp = resolver.para_props(shape, para)
            algn = pp['algn'] if pp['algn'] in ('l', 'ctr', 'r') else 'l'
            toks = []
            for run in para.runs:
                props = resolver.run_props(shape, para, run)
                props['sz_pt'] = props['sz_pt'] * font_scale
                st = _style_from_props(run, props, theme_colors)
                frame_max_sz = max(frame_max_sz, st.size_pt)
                for tk in _tokenize(run.text):
                    toks.append((tk, st))
            paras_data.append((algn, pp['lnSpc_pct'], toks))

        # 行排版数据（typographic 模型，经 COM 对照实测）：
        # 行距推进 = descent(上一行) + ascent(下一行)，ascent=(a+g×lnSpc)×字号px。
        # 同字号时 = 行高（与历史校准自洽）；字号骤增的相邻行 PowerPoint 按此
        # 保证大字不压上一行（旧"按上行行高推进"模型会导致大字上叠，如承保贺报）。
        all_lines = []  # [(algn, line[(tk,style,adv)], ascent_px, descent_px)]
        for algn, lnSpc, toks in paras_data:
            para_dom = max(toks, key=lambda t: t[1].size_pt)[1] if toks else None
            para_max_sz = para_dom.size_pt if para_dom is not None else frame_max_sz

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

            def line_metrics(dom, sz):
                """按行主导字体返回 (ascent_px, descent_px)"""
                em_px = pt2px(sz)
                m = font_registry.metrics_for(dom.font_path) if dom is not None else None
                if m is not None:
                    asc = (m.baseline_a + m.baseline_g * lnSpc) * em_px
                    lh = (em_px * m.line_height_ratio * lnSpc
                          * (1.0 - bp.get('lnspc_reduction', 0.0)))
                else:
                    asc = BASELINE0_FRAC * LINE_HEIGHT_RATIO * lnSpc * em_px
                    lh = (em_px * LINE_HEIGHT_RATIO * lnSpc
                          * (1.0 - bp.get('lnspc_reduction', 0.0)))
                return asc, lh - asc

            if lines:
                for line in lines:
                    dom = max(line, key=lambda t: t[1].size_pt)[1]
                    asc, desc = line_metrics(dom, dom.size_pt)
                    all_lines.append((algn, line, asc, desc))
            else:
                # 空段落占一个整行盒（ascent=0 / descent=完整行高，块高不变）
                _asc, _desc = line_metrics(para_dom, para_max_sz)
                all_lines.append((algn, [], 0.0, _asc + _desc))

        if not all_lines:
            return

        # 2) anchor 垂直定位（文本块总高 = Σ(ascent+descent)，同字号时 ≡ Σ行高）
        total_h = sum(asc + desc for _a, _l, asc, desc in all_lines)
        content_h = box_h_px - ins['tIns'] - ins['bIns']
        if bp['anchor'] == 'ctr':
            y_off = max(0.0, (content_h - total_h) / 2.0)
        elif bp['anchor'] in ('b', 'just'):
            y_off = max(0.0, content_h - total_h)
        else:
            y_off = 0.0

        layers = {}   # style.key() -> [(char, style, x_px, baseline_px)]
        y_baseline = box_t + ins['tIns'] + y_off
        prev_desc = 0.0
        first = True
        for algn, line, asc, desc in all_lines:
            if first:
                y_baseline += asc
                first = False
            else:
                # 行距推进：上一行下延 + 本行上伸（PowerPoint typographic 模型）
                y_baseline += prev_desc + asc
            prev_desc = desc
            if not line:
                continue  # 空段落占位行：无字形可画
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
                if st.italic:
                    self._draw_italic_glyph(mask, md, ch, f, x, yb)
                else:
                    md.text((x, yb), ch, font=f, fill=255, anchor='ls')
            # 删除线/下划线画入蒙版 -> 自动继承纯色/渐变/图片填充与阴影
            if st.strike or st.underline:
                self._draw_decorations(md, glyphs, f)
            bbox = mask.getbbox()
            if not bbox:
                continue
            if st.shadow:
                s = st.shadow
                rgb, alpha = parse_color_el(s['clr_el'], theme_colors)
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
            if st.glow:
                g = st.glow
                rgb, alpha = parse_color_el(g['clr_el'], theme_colors)
                rad_px = pt2px(g['rad_pt'])
                # 发光 = 文字蒙版高斯外扩：多层递增 σ（贴边亮环 -> 外围扩散），
                # lighter 合成后乘发光色 alpha；与文字层做 screen（加色）混合——
                # PowerPoint 的 glow 是光晕，会照亮背景而非压暗（普通 alpha 混合偏暗）。
                # σ 组与增益经红/橙/黄/绿/蓝五色 COM 探针统一校准，与色相无关
                glow_a = Image.new('L', (W, H), 0)
                for k in GLOW_SIGMA_KS:
                    glow_a = ImageChops.lighter(
                        glow_a, mask.filter(ImageFilter.GaussianBlur(rad_px * k)))
                g_img = Image.new('RGBA', (W, H), (*rgb, 0))
                g_img.putalpha(glow_a.point(_glow_lut(alpha * 255)))
                text_layer = ImageChops.screen(text_layer, g_img)
            if st.stroke:
                w_pt, rgb, alpha = st.stroke
                stroke_mask = Image.new('L', (W, H), 0)
                sd = ImageDraw.Draw(stroke_mask)
                # OOXML 描边宽为全宽，Pillow stroke_width 每侧外扩 -> 取半
                sw = max(1, int(round(pt2px(w_pt) / 2)))
                for ch, _, x, yb in glyphs:
                    sd.text((x, yb), ch, font=f, fill=0, anchor='ls',
                            stroke_width=sw, stroke_fill=255)
                stroke_img = Image.new('RGBA', (W, H), (*rgb, 0))
                stroke_img.putalpha(
                    stroke_mask.point(lambda v, a=int(alpha * 255): v * a // 255))
                text_layer.alpha_composite(stroke_img)
            # 文字渐变/图片填充在【文本框 box】上投影（PowerPoint 语义）。
            # 生成范围取框与文字实际范围的并集：溢出文字的渐变按端点色延伸
            # （PowerPoint 行为），而不是落进框外透明黑变成黑字
            box_w_px = shape.width * scale * xform.sx
            mb = mask.getbbox() or (0, 0, 1, 1)
            gx0 = int(min(box_l, mb[0])) - 1
            gy0 = int(min(box_t, mb[1])) - 1
            gx1 = int(max(box_l + box_w_px, mb[2])) + 1
            gy1 = int(max(box_t + box_h_px, mb[3])) + 1
            gw, gh = max(1, gx1 - gx0), max(1, gy1 - gy0)
            if st.fill and st.fill[0] == 'grad':
                _, stops, ang = st.fill
                grad = _make_linear_gradient(
                    gw, gh, stops, ang,
                    ref_w=box_w_px, ref_h=box_h_px,
                    off_x=gx0 - int(box_l), off_y=gy0 - int(box_t))
                tile = Image.new('RGBA', (W, H), (0, 0, 0, 0))
                tile.paste(grad, (gx0, gy0))
                tile.putalpha(mask)
                text_layer.alpha_composite(tile)
            elif st.fill and st.fill[0] == 'blip':
                # 图片文字填充：图片拉伸铺满文本框 box；溢出区按图片边缘延伸
                # （PowerPoint 为 tileRect 平铺语义，海报场景边缘延伸已足够近似）
                _, blob, rect = st.fill
                img = _load_rgba(
                    blob, (box_w_px, box_h_px),
                    crop_rect=(rect.get('l', 0), rect.get('t', 0),
                               rect.get('r', 0), rect.get('b', 0))
                    if rect else None)
                pl = max(0, int(box_l) - gx0)
                pt_ = max(0, int(box_t) - gy0)
                pr = max(0, gx1 - int(box_l + box_w_px))
                pb = max(0, gy1 - int(box_t + box_h_px))
                bw, bh = max(1, int(box_w_px)), max(1, int(box_h_px))
                cover = Image.new('RGBA', (gw, gh), (0, 0, 0, 0))
                if pl or pt_ or pr or pb:   # 边缘延伸到溢出区
                    top = img.crop((0, 0, bw, 1)).resize((bw, pt_)) if pt_ else None
                    bot = img.crop((0, bh - 1, bw, bh)).resize((bw, pb)) if pb else None
                    ih = bh + (pt_ or 0) + (pb or 0)
                    left = img.crop((0, 0, 1, bh)).resize((pl, ih)) if pl else None
                    right = img.crop((bw - 1, 0, bw, bh)).resize((pr, ih)) if pr else None
                    cover.paste(top, (pl, 0))
                    cover.paste(bot, (pl, (pt_ or 0) + bh))
                    cover.paste(img, (pl, pt_ or 0))
                    cover.paste(left, (0, 0))
                    cover.paste(right, (pl + bw, 0))
                else:
                    cover.paste(img, (0, 0))
                tile = Image.new('RGBA', (W, H), (0, 0, 0, 0))
                tile.paste(cover, (gx0, gy0))
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

    @staticmethod
    def _draw_italic_glyph(mask, md, ch, f, x, yb):
        """合成斜体：字形画到临时画布后绕基线斜切 12°，加法式贴入蒙版
        （paste 以临时画布自身为掩码，避免覆盖相邻字形）"""
        ascent, descent = f.getmetrics()
        k = ITALIC_SHEAR
        pad = int(k * (ascent + descent)) + 4
        gw = int(math.ceil(f.getlength(ch))) + pad + 4
        gh = ascent + descent + 4
        tmp = Image.new('L', (max(2, gw), max(2, gh)), 0)
        ImageDraw.Draw(tmp).text((2, ascent + 2), ch, font=f, fill=255,
                                 anchor='ls')
        B = ascent + 2  # 临时画布内的基线行
        tmp = tmp.transform(tmp.size, Image.AFFINE,
                            (1, k, -k * B, 0, 1, 0), resample=Image.BILINEAR)
        mask.paste(tmp, (int(round(x)) - 2, int(round(yb)) - ascent - 2), tmp)

    @staticmethod
    def _draw_decorations(md, glyphs, f):
        """删除线/下划线：按行画横条到文字蒙版。
        行 = 相同基线 yb 的字形集合；横条跨度取该行字形 x 范围。"""
        em = f.size
        thick = max(1, int(round(STRIKE_THICK_EM * em)))
        lines = {}
        for ch, _, x, yb in glyphs:
            info = lines.setdefault(round(yb, 1), [x, x + f.getlength(ch)])
            info[0] = min(info[0], x)
            info[1] = max(info[1], x + f.getlength(ch))
        for yb, (x0, x1) in lines.items():
            if glyphs[0][1].strike:
                y = yb - font_registry.metrics_for(
                    glyphs[0][1].font_path).strike_pos_em * em
                md.rectangle((x0, y, x1, y + thick), fill=255)
            u = glyphs[0][1].underline
            if u:
                y = yb + UNDERLINE_POS_EM * em
                md.rectangle((x0, y, x1, y + thick), fill=255)
                if u == 'dbl':  # 双下划线：第二条再低一档
                    y2 = y + thick + max(1, thick)
                    md.rectangle((x0, y2, x1, y2 + thick), fill=255)

    def _font(self, style):
        key = (style.font_path, style.size_pt, style.bold, style.italic)
        if key not in self.font_cache:
            size_px = style.size_pt * self._dpi / 72.0
            try:
                font = ImageFont.truetype(style.font_path, size_px)
            except OSError as e:
                # 最终兜底：常见 CJK 字体链逐个尝试，避免单个字体文件
                # 缺失导致整个导出失败
                for fb in (r'C:\Windows\Fonts\msyh.ttc',
                           r'C:\Windows\Fonts\msyhbd.ttc',
                           r'C:\Windows\Fonts\simsun.ttc',
                           r'C:\Windows\Fonts\simhei.ttf'):
                    if os.path.exists(fb):
                        try:
                            font = ImageFont.truetype(fb, size_px)
                            logger.warning(
                                f"字体加载失败 {style.font_path!r}: {e}，回退 {fb}")
                            break
                        except OSError:
                            continue
                else:
                    raise
            self.font_cache[key] = font
        return self.font_cache[key]
