# -*- coding: utf-8 -*-
"""合成核心。

两类入口：
- ``linear_gradient``：生成线性渐变图。
- ``draw_glyphs``：把「一组已经排好位置的字形 + 一种样式」合成到一张透明 RGBA 图层上。

这里只做绘制，不做排版——字放哪里由调用方决定。PPTX 渲染引擎有自己的一套排版规则，
``fontfx.text`` 里是给普通项目用的简单排版。
"""

import functools
import math
from typing import Iterable, Optional, Sequence, Tuple

from PIL import Image, ImageChops, ImageDraw, ImageFilter

from .styles import Decoration, ImageFill, LinearGradient, SolidFill, Style, normalize_stops

# ---------------- 发光模型（对照 PowerPoint 文字发光校准，红/橙/黄/绿/蓝五色探针统一）----------------
# 发光 = 字形蒙版做多层递增 σ 的高斯外扩取最亮值，再按发光色着色并与已有内容做滤色。
# 三常数与色相无关：强度比全色 ±4%，光晕面积比 ≤+19%——换任意发光颜色无需重调。
GLOW_SIGMA_KS = (0.08, 0.18, 0.3)   # 各层 sigma = 发光半径 × k
GLOW_GAIN = 2.45                    # alpha 增益（乘发光色 alpha 后封顶 255）
GLOW_GAMMA = 0.7                    # alpha 伽马（<1 提亮：光晕边缘更浓）

ITALIC_SHEAR = 0.2126               # 合成斜体的斜切率（tan 12°）


# ---------------------------------------------------------------- 渐变

def linear_gradient(size, stops, angle: float = 0.0,
                    extent: Optional[Tuple[float, float]] = None,
                    origin: Tuple[float, float] = (0.0, 0.0)) -> Image.Image:
    """线性渐变图（RGBA）。

    Args:
        size: (宽, 高)，像素。
        stops: ``[(位置 0~1, 颜色[, 透明度])]``；位置需要升序，同一位置取后者。
        angle: 方向角（度）。0° = 左→右，90° = 上→下，45° = 左上→右下。
        extent: 参照框的 (宽, 高)。渐变的 0 和 1 分别落在参照框沿角度方向投影的两个
            端点上，框外取端点色。None = 以 size 为参照。
        origin: 参照框左上角在渐变图里的坐标 (x, y)，可以在图外（负值）。

    装了 numpy 走向量化实现；没装则自动退化为纯 Pillow 实现（同一套数学，
    位置多一层 8bit 量化，差异 ≤ 2/255）。
    """
    w, h = int(size[0]), int(size[1])
    if w <= 0 or h <= 0:
        raise ValueError(f"size 需要是正的 (宽, 高): {size!r}")
    stops = normalize_stops(stops)
    if len(stops) == 1:
        stops = stops * 2

    theta = math.radians(angle)
    vx, vy = math.cos(theta), math.sin(theta)
    rw, rh = (extent[0] or w, extent[1] or h) if extent else (w, h)
    corners = [(0, 0), (rw - 1, 0), (0, rh - 1), (rw - 1, rh - 1)]
    projs = [x * vx + y * vy for x, y in corners]
    p0, p1 = min(projs), max(projs)
    span = (p1 - p0) or 1.0
    off_x, off_y = -origin[0], -origin[1]

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
    return Image.fromarray(arr, "RGBA")


def _gradient_pure(w, h, stops, vx, vy, p0, span, off_x=0.0, off_y=0.0):
    """无 numpy 的渐变实现：按行线性投影量化 t，再用 bytes.translate 做 C 级逐通道
    LUT 映射。整幅画布级的大渐变会明显慢于 numpy 路径，文字级的小区域无感。"""
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
    lut_a = build_lut([s[2] * 255 for s in stops])

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
    chans = [Image.frombytes("L", (w, h), b"".join(r)) for r in rows]
    return Image.merge("RGBA", chans)


# ---------------------------------------------------------------- 字形绘制

def draw_italic_glyph(mask, md, ch, font, x, baseline):
    """合成斜体：字形画到临时画布后绕基线斜切 12°，加法式贴入蒙版
    （以临时画布自身为掩码，不覆盖相邻字形）"""
    ascent, descent = font.getmetrics()
    k = ITALIC_SHEAR
    pad = int(k * (ascent + descent)) + 4
    gw = int(math.ceil(font.getlength(ch))) + pad + 4
    gh = ascent + descent + 4
    tmp = Image.new("L", (max(2, gw), max(2, gh)), 0)
    ImageDraw.Draw(tmp).text((2, ascent + 2), ch, font=font, fill=255, anchor="ls")
    B = ascent + 2  # 临时画布内的基线行
    tmp = tmp.transform(tmp.size, Image.AFFINE,
                        (1, k, -k * B, 0, 1, 0), resample=Image.BILINEAR)
    mask.paste(tmp, (int(round(x)) - 2, int(round(baseline)) - ascent - 2), tmp)


def _draw_decoration(md, glyphs, font, dec):
    """一条装饰线：基线相同的字形归为一行，横条跨度取该行字形的 x 范围。"""
    lines = {}
    for ch, x, yb in glyphs:
        info = lines.setdefault(round(yb, 1), [x, x + font.getlength(ch)])
        info[0] = min(info[0], x)
        info[1] = max(info[1], x + font.getlength(ch))
    thick = max(1, int(round(dec.thickness)))
    for yb, (x0, x1) in lines.items():
        y = yb + dec.offset
        md.rectangle((x0, y, x1, y + thick), fill=255)
        if dec.kind == "underline" and dec.double:
            y2 = y + thick + max(1, thick)
            md.rectangle((x0, y2, x1, y2 + thick), fill=255)


@functools.lru_cache(maxsize=256)
def _glow_lut(alpha_255):
    """发光蒙版值 -> 最终 alpha 的 256 项 LUT：先伽马再乘增益封顶"""
    top = min(255, int(alpha_255 * GLOW_GAIN))
    return [int((v / 255.0) ** GLOW_GAMMA * top) for v in range(256)]


def _dim_lut(alpha):
    """蒙版值按透明度缩放的 256 项 LUT"""
    a = int(alpha * 255)
    return [v * a // 255 for v in range(256)]


# ---------------------------------------------------------------- 合成

def draw_glyphs(layer: Image.Image, glyphs: Iterable, font, style: Optional[Style] = None) -> Image.Image:
    """把一组同样式的字形画到 ``layer`` 上。**就地修改，返回入参本身。**

    一般先建一张全透明的 RGBA 图层，把各组样式依次画进去，最后整体贴到底图上。
    叠放顺序：阴影 → 发光 → 描边 → 填充。

    Args:
        layer: RGBA 图层，会被就地修改。
        glyphs: ``Glyph`` 或 ``(字符, x, 基线 y)`` 三元组；坐标相对 ``layer``，浮点像素。
        font: ``ImageFont`` 字体。
        style: 样式；None = 黑色实心。

    Returns:
        入参 ``layer``。
    """
    if layer.mode != "RGBA":
        raise ValueError("layer 需要是 RGBA 模式")
    style = style or Style()
    glyphs = [(g[0], g[1], g[2]) for g in glyphs]
    W, H = layer.size
    if not glyphs:
        return layer

    mask = Image.new("L", (W, H), 0)
    md = ImageDraw.Draw(mask)
    for ch, x, yb in glyphs:
        if style.synthetic_italic:
            draw_italic_glyph(mask, md, ch, font, x, yb)
        else:
            md.text((x, yb), ch, font=font, fill=255, anchor="ls")
    for dec in style.decorations:
        _draw_decoration(md, glyphs, font, dec)
    bbox = mask.getbbox()
    if not bbox:
        return layer

    if style.shadow is not None:
        s = style.shadow
        sh = mask.filter(ImageFilter.GaussianBlur(s.blur))
        sh_img = Image.new("RGBA", (W, H), (*s.color, 0))
        sh_img.putalpha(sh.point(_dim_lut(s.alpha)))
        layer.alpha_composite(
            sh_img, (int(round(s.offset[0])), int(round(s.offset[1]))))

    if style.glow is not None:
        g = style.glow
        glow_a = Image.new("L", (W, H), 0)
        for k in GLOW_SIGMA_KS:
            glow_a = ImageChops.lighter(
                glow_a, mask.filter(ImageFilter.GaussianBlur(g.radius * k)))
        g_img = Image.new("RGBA", (W, H), (*g.color, 0))
        g_img.putalpha(glow_a.point(_glow_lut(g.alpha * 255)))
        # 滤色会产生新图；拷回入参，保证这个函数永远就地修改
        layer.paste(ImageChops.screen(layer, g_img))

    if style.stroke is not None:
        s = style.stroke
        stroke_mask = Image.new("L", (W, H), 0)
        sd = ImageDraw.Draw(stroke_mask)
        for ch, x, yb in glyphs:
            sd.text((x, yb), ch, font=font, fill=0, anchor="ls",
                    stroke_width=s.width, stroke_fill=255)
        stroke_img = Image.new("RGBA", (W, H), (*s.color, 0))
        stroke_img.putalpha(stroke_mask.point(_dim_lut(s.alpha)))
        layer.alpha_composite(stroke_img)

    fill = style.fill
    if isinstance(fill, LinearGradient):
        # 只在「参照框 ∪ 字形范围」外加 1px 余量的区域里生成渐变，框外取端点色
        gx0, gy0, gx1, gy1, fox, foy, fw, fh = _fill_region(bbox, fill.frame)
        grad = linear_gradient((max(1, gx1 - gx0), max(1, gy1 - gy0)), fill.stops,
                               fill.angle, extent=(fw, fh),
                               origin=(int(fox) - gx0, int(foy) - gy0))
        tile = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        tile.paste(grad, (gx0, gy0))
        tile.putalpha(mask)
        layer.alpha_composite(tile)
    elif isinstance(fill, ImageFill):
        gx0, gy0, gx1, gy1, fox, foy, fw, fh = _fill_region(bbox, fill.frame)
        img = fill.image() if callable(fill.image) else fill.image
        if img.mode != "RGBA":
            img = img.convert("RGBA")
        bw, bh = max(1, int(fw)), max(1, int(fh))
        if img.size != (bw, bh):
            img = img.resize((bw, bh), Image.LANCZOS)
        # 文字超出参照框的部分取图片边缘像素延伸
        pl = max(0, int(fox) - gx0)
        pt_ = max(0, int(foy) - gy0)
        pr = max(0, gx1 - int(fox + fw))
        pb = max(0, gy1 - int(foy + fh))
        cover = Image.new("RGBA", (max(1, gx1 - gx0), max(1, gy1 - gy0)), (0, 0, 0, 0))
        if pl or pt_ or pr or pb:
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
        tile = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        tile.paste(cover, (gx0, gy0))
        tile.putalpha(mask)
        layer.alpha_composite(tile)
    else:
        rgb, alpha = (0, 0, 0), 1.0
        if isinstance(fill, SolidFill):
            rgb, alpha = fill.color, fill.alpha
        fill_img = Image.new("RGBA", (W, H), (*rgb, 0))
        fill_img.putalpha(mask.point(_dim_lut(alpha)))
        layer.alpha_composite(fill_img)
    return layer


def _fill_region(bbox, frame):
    """填充相关的计算区域与参照框。返回 (gx0, gy0, gx1, gy1, 框 x, 框 y, 框宽, 框高)。

    没给参照框时用字形包围盒自己（外扩 1px 余量）。"""
    if frame is None:
        fox, foy, fw, fh = float(bbox[0]), float(bbox[1]), float(bbox[2] - bbox[0]), float(bbox[3] - bbox[1])
    else:
        fox, foy, fw, fh = frame
    gx0 = int(min(fox, bbox[0])) - 1
    gy0 = int(min(foy, bbox[1])) - 1
    gx1 = int(max(fox + fw, bbox[2])) + 1
    gy1 = int(max(foy + fh, bbox[3])) + 1
    return gx0, gy0, gx1, gy1, fox, foy, fw, fh
