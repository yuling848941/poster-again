"""
纯 Python 图片导出器（Pillow 渲染引擎）
不依赖 Microsoft Office / WPS Office，直接将 PPTX 渲染为图片

本模块负责「PPTX 语义 → 像素」：解析格式继承、排字、算出每个字形的位置，
再把「这个字长什么样」交给 fontfx（渐变/图片填充、发光、阴影、描边、装饰线）。

支持范围（受限模板规范，海报模板家族）：
- 背景图/图片（p:pic，含 srcRect 裁剪）、图片/纯色/线性渐变填充的形状
- 文本框：字体/字号/粗体/斜体/字距(spc)/行距/对齐/锚定、normAutofit 缩放
- 文字填充（纯色/线性渐变/图片）、文字外阴影、发光、描边、删除线、下划线
- 预设形状几何、旋转/翻转、组合递归
- 多页演示文稿、PNG/JPG 输出、quality 缩放倍数（对应 96dpi 的倍数）

不支持的元素（render_precheck 白名单外）：表格、图表、SmartArt、艺术字变形、
3D 效果、路径渐变等——渲染时打印警告并跳过，不中断导出。

排版度量经 PowerPoint COM 导出对照校准（见 experiments/calibrate*.py），
校准常数按物理字体绑定在 font_registry.METRICS_*。
"""

import logging
import io
import math
import os
import re
import tempfile
from typing import List, Optional

import fontfx
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from .format_inheritance import (
    FormatResolver, load_theme, parse_color_el, parse_fill, resolve_bg_ref,
    grad_stops,
)
from .shape_geometry import preset_geometry, draw_path
from . import font_registry

logger = logging.getLogger(__name__)

# ---------------- 校准常数（对照 PowerPoint 导出图调整，勿随意改动） ----------------
# 行距/基线常数已按物理字体拆分到 font_registry（metrics_for 按字体文件查询），
# 此处保留全局兜底值（= 雅黑 COM 实测值）
LINE_HEIGHT_RATIO = 1.25   # 微软雅黑 单倍行距/字号（PowerPoint 按字体 metrics）
BASELINE0_FRAC = 0.72      # 首行基线在行高盒中的位置
SHADOW_SIGMA_FRAC = 0.5    # PowerPoint blurRad -> 高斯模糊 sigma 的系数（形状与文字通用）
STRIKE_THICK_EM = 0.05     # 删除线/下划线厚度 0.05em
UNDERLINE_POS_EM = 0.09    # 下划线位于基线下方 0.09em
# 发光观感（多层 sigma/增益/伽马）与合成斜体的斜切率在 fontfx 里，同样已 COM 对照校准
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


def _pic_cache_budget():
    """图片缓存上限（字节）。默认 256MB；环境变量 POSTERAGAIN_PIC_CACHE_MB 可改，0 = 关闭。

    缓存的是「缩放后」的位图，单张大小 = 宽 × 高 × 4：96dpi 的 1792×3584 海报
    背景约 26MB，3 倍画质约 231MB。移动端建议在导入前把上限调到 64 左右。
    """
    try:
        mb = float(os.environ.get('POSTERAGAIN_PIC_CACHE_MB', '256'))
    except ValueError:
        mb = 256.0
    return max(0.0, mb) * 1024 * 1024


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
                 'glow', 'stroke', 'strike', 'underline', 'font_path',
                 'font_index', 'font_variation')

    def __init__(self, size_pt, bold, italic, spc_pt, fill, shadow, glow,
                 stroke, strike, underline, font_path,
                 font_index=0, font_variation=''):
        self.size_pt, self.bold, self.italic = size_pt, bold, italic
        self.spc_pt, self.fill, self.shadow = spc_pt, fill, shadow
        self.glow, self.stroke = glow, stroke
        self.strike, self.underline = strike, underline
        # 字体的完整身份 = 文件 + TTC 字面序号 + 可变字体实例
        self.font_path = font_path
        self.font_index = font_index
        self.font_variation = font_variation

    def key(self):
        # 必须含字体身份：同一段文字按 key 合并后只用首字的字体画，
        # 不同字体的两段若被并到一起，后面那段会画成前面的字体
        return (self.size_pt, self.bold, self.italic, self.spc_pt,
                repr(self.fill), repr(self.shadow), repr(self.glow),
                repr(self.stroke), self.strike, self.underline,
                self.font_path, self.font_index, self.font_variation)


def _font_path_for(typeface, bold):
    """字体名 -> 字体文件；走字体注册表（策略见 font_registry），无命中回退默认并告警"""
    return font_registry.resolve(typeface, bold).path


def _style_from_props(run, props, theme_colors):
    """把继承解析出的有效属性转换为渲染样式（阴影/描边只取显式 rPr）"""
    rPr = run._r.find(qn('a:rPr'))
    shadow = _parse_shadow(rPr)
    glow = _parse_glow(rPr)
    stroke = _parse_stroke(rPr, theme_colors)
    spec = font_registry.resolve(props['typeface'], props['bold'])
    return _Style(props['sz_pt'], props['bold'], props['italic'], props['spc_pt'],
                  props['fill'], shadow, glow, stroke,
                  props.get('strike', False), props.get('underline'),
                  spec.path, spec.index, spec.variation)


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
        # 图片缓存：批量渲染时同一张背景图逐行复用（见 _picture）
        self._pic_cache = {}
        self._pic_cache_pkg = None
        self._pic_cache_bytes = 0
        self._pic_cache_budget = _pic_cache_budget()

    @staticmethod
    def is_available() -> bool:
        """纯 Python 引擎始终可用"""
        return True

    @staticmethod
    def get_engine_info() -> str:
        return "纯 Python 渲染引擎（Pillow，无需 Office/WPS）"

    def precheck(self, pptx_path: str):
        """渲染预检（白名单制），返回问题清单；字体检查同时看捆绑字体和系统已装字体"""
        from .render_precheck import RenderPrecheck
        installed = font_registry.available_typefaces()
        try:
            installed |= font_registry.installed_system_fonts()
        except Exception as e:
            logger.debug(f"系统字体检查跳过（仅用注册表字体集）: {e}")
        return RenderPrecheck(installed_fonts=installed,
                              font_resolver=font_registry.is_resolvable).check_pptx(pptx_path)

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
                grad = fontfx.linear_gradient(
                    (max(1, int(round(w))), max(1, int(round(h)))), stops, ang)
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
                    img = self._cached_part_rgba(
                        part, (int(round(w)), int(round(h))),
                        crop=(rect.get('l', 0), rect.get('t', 0),
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
                    canvas.alpha_composite(fontfx.linear_gradient((W, H), stops, ang))
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
                    canvas.alpha_composite(fontfx.linear_gradient(
                        (W, H), stops, int(lin.get('ang')) / 60000.0))
                    return True
                return False
            blipFill = bgPr.find(qn('a:blipFill'))
            if blipFill is not None:
                blip = blipFill.find(qn('a:blip'))
                if blip is not None:
                    rId = blip.get(qn('r:embed'))
                    part = container.part.related_part(rId)
                    canvas.alpha_composite(self._cached_part_rgba(part, (W, H)))
                    return True
            return False
        except Exception as e:
            logger.warning(f"[纯Python渲染] 幻灯片背景渲染失败（保持白底）: {e}")
            return False

    def _draw_picture(self, canvas, shape, scale, xform=None):
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
        img = self._picture(shape, (w, h), crop)
        canvas.alpha_composite(img, (int(round(xform.mx(shape.left * scale))),
                                     int(round(xform.my(shape.top * scale)))))

    def _picture(self, shape, size, crop):
        """图片形状的 RGBA 位图（已缩放到 size），走批量图片缓存。"""
        try:
            rId = shape._element.blipFill.blip.rEmbed
            part = shape.part.related_part(rId)
        except Exception:
            return _load_rgba(shape.image.blob, size, crop_rect=crop)
        return self._cached_part_rgba(part, size, crop)

    def _cached_part_rgba(self, part, size, crop=None):
        """图片部件 -> 缩放到 size 的 RGBA 位图，按「包 + 部件 + 尺寸 + 裁剪」缓存。

        批量渲染时背景图 / 图片形状这类静态图片逐行复用，省掉每行一次的整图解码
        与 LANCZOS 缩放（50MB 海报模板的大背景图上约占每行 1.7 秒）。
        缓存的是位图对象本身，命中与否像素完全一致。
        """
        if self._pic_cache_budget <= 0:
            return _load_rgba(part.blob, size, crop_rect=crop)
        pkg = part.package
        if pkg is not self._pic_cache_pkg:
            # 换了文档：旧缓存全部作废（持有包的强引用，也避免 id 被别的对象复用）
            self._pic_cache_pkg = pkg
            self._pic_cache = {}
            self._pic_cache_bytes = 0
        key = (str(part.partname), size[0], size[1], crop)
        img = self._pic_cache.get(key)
        if img is not None:
            return img
        img = _load_rgba(part.blob, size, crop_rect=crop)
        cost = size[0] * size[1] * 4
        if self._pic_cache_bytes + cost <= self._pic_cache_budget:
            self._pic_cache[key] = img
            self._pic_cache_bytes += cost
        return img

    def _draw_text_frame(self, canvas, shape, resolver, theme_colors, quality, scale,
                         xform=None):
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
        box_w_px = shape.width * scale * xform.sx      # 渐变/图片填充的参照框宽
        avail = box_w_px - ins['lIns'] - ins['rIns']
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

        # 3) 按样式分组交给 fontfx 合成（阴影 -> 发光 -> 描边 -> 填充）
        if not layers:
            return
        W, H = canvas.size
        text_layer = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        for _, glyphs in layers.items():
            st = glyphs[0][1]
            style = self._text_style(st, theme_colors, pt2px,
                                     (box_l, box_t, box_w_px, box_h_px))
            fontfx.draw_glyphs(text_layer,
                               ((ch, x, yb) for ch, _, x, yb in glyphs),
                               self._font(st), style)
        canvas.alpha_composite(text_layer)

    def _text_style(self, st, theme_colors, pt2px, box):
        """PPTX 样式 -> fontfx.Style。单位换算（磅 -> 像素）和 PowerPoint 语义都收在这一层。

        box=(x, y, 宽, 高) 是文本框在画布上的位置：渐变和图片填充以它为参照框，
        文字溢出框时渐变取端点色、图片取边缘延伸（与 PowerPoint 一致）。
        """
        decs = []
        if st.strike or st.underline:
            em = self._font(st).size
            thick = max(1, int(round(STRIKE_THICK_EM * em)))
            if st.strike:
                decs.append(fontfx.Decoration(
                    'strike',
                    -font_registry.metrics_for(st.font_path).strike_pos_em * em, thick))
            if st.underline:
                decs.append(fontfx.Decoration(
                    'underline', UNDERLINE_POS_EM * em, thick,
                    double=(st.underline == 'dbl')))

        shadow = glow = stroke = fill = None
        if st.shadow:
            s = st.shadow
            rgb, alpha = parse_color_el(s['clr_el'], theme_colors)
            dist = pt2px(s['dist_pt'])
            th = math.radians(s['dir_deg'])
            shadow = fontfx.Shadow(
                rgb, alpha=alpha, blur=pt2px(s['blur_pt']) * SHADOW_SIGMA_FRAC,
                offset=(int(round(dist * math.cos(th))),
                        int(round(dist * math.sin(th)))))
        if st.glow:
            g = st.glow
            rgb, alpha = parse_color_el(g['clr_el'], theme_colors)
            glow = fontfx.Glow(rgb, radius=pt2px(g['rad_pt']), alpha=alpha)
        if st.stroke:
            w_pt, rgb, alpha = st.stroke
            # OOXML 描边宽为全宽，Pillow stroke_width 每侧外扩 -> 取半
            stroke = fontfx.Stroke(
                rgb, width=max(1, int(round(pt2px(w_pt) / 2))), alpha=alpha)
        if st.fill:
            kind = st.fill[0]
            if kind == 'solid':
                _, rgb, alpha = st.fill
                fill = fontfx.SolidFill(rgb, alpha)
            elif kind == 'grad':
                _, stops, ang = st.fill
                fill = fontfx.LinearGradient(stops, ang, frame=box)
            elif kind == 'blip':
                _, blob, rect = st.fill
                crop = ((rect.get('l', 0), rect.get('t', 0),
                         rect.get('r', 0), rect.get('b', 0)) if rect else None)
                bw, bh = box[2], box[3]
                fill = fontfx.ImageFill(
                    lambda: _load_rgba(blob, (bw, bh), crop_rect=crop), frame=box)
        return fontfx.Style(fill=fill, shadow=shadow, glow=glow, stroke=stroke,
                            synthetic_italic=st.italic, decorations=tuple(decs))

    def _font(self, style):
        key = (style.font_path, style.size_pt, style.bold, style.italic,
               style.font_index, style.font_variation)
        if key not in self.font_cache:
            size_px = style.size_pt * self._dpi / 72.0
            try:
                font = ImageFont.truetype(style.font_path, size_px,
                                          index=style.font_index)
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
            if style.font_variation:
                # 可变字体：切到指定的具名实例（粗细/宽窄）；切不了就用默认实例
                try:
                    font.set_variation_by_name(style.font_variation)
                except Exception as e:
                    logger.debug(f"可变字体实例 {style.font_variation!r} 不可用，"
                                 f"用默认实例: {e}")
            self.font_cache[key] = font
        return self.font_cache[key]
