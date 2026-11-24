# -*- coding: utf-8 -*-
"""简单排版 + 一步到位的接口：``text_image`` / ``draw_text`` / ``measure``。

排版规则很直接：逐字符按 ``font.getlength`` 推进（不加字距调整），``\\n`` 换行，
行高 = (上升部 + 下降部) × ``line_spacing``。需要复杂排版的请自己算好每个字的位置，
再用 ``fontfx.draw_glyphs`` 画。
"""

import math
from typing import Optional, Tuple

from PIL import Image, ImageFont

from .core import ITALIC_SHEAR, draw_glyphs
from .styles import Style, default_decorations

_ANCHOR_H = "lmr"
_ANCHOR_V = "amsbd"   # 顶部(上升部) / 竖直居中 / 基线 / 末行下沿 / 整块下沿


def _line_width(line, font, spacing):
    return sum(font.getlength(ch) for ch in line) + spacing * max(0, len(line) - 1)


def _resolve_style(style, font, underline, strike):
    if style is None:
        style = Style()
    if not style.decorations and (underline or strike):
        style = Style(**{**style.__dict__,
                         "decorations": default_decorations(font, underline, strike)})
    return style


def _auto_padding(style, font) -> int:
    """让阴影、发光、描边、装饰线都不被裁掉所需的余量"""
    need = 3
    if style.stroke is not None:
        need = max(need, style.stroke.width + 3)
    if style.glow is not None:
        need = max(need, style.glow.radius * 0.9)       # 最大 σ = 0.3r，取 3σ
    if style.shadow is not None:
        s = style.shadow
        need = max(need, s.blur * 3 + max(abs(s.offset[0]), abs(s.offset[1])))
    for dec in style.decorations:
        need = max(need, abs(dec.offset) + 2 * dec.thickness + 2)
    return int(math.ceil(need)) + 2


def measure(text, font, spacing: float = 0.0, line_spacing: float = 1.0) -> Tuple[float, float]:
    """文字块的 (宽, 高)，单位像素，不含特效余量。"""
    lines = text.split("\n")
    ascent, descent = font.getmetrics()
    return (max((_line_width(line, font, spacing) for line in lines), default=0.0),
            len(lines) * (ascent + descent) * line_spacing)


def _layout(text, font, style, spacing, line_spacing, align, padding, underline, strike):
    if align not in ("left", "center", "right"):
        raise ValueError(f"align 只能是 'left' / 'center' / 'right': {align!r}")
    style = _resolve_style(style, font, underline, strike)
    lines = text.split("\n")
    ascent, descent = font.getmetrics()
    line_h = (ascent + descent) * line_spacing
    widths = [_line_width(line, font, spacing) for line in lines]
    width = max(widths, default=0.0)
    if style.synthetic_italic:
        width += ITALIC_SHEAR * ascent            # 斜切后右上角伸出的量
    height = len(lines) * line_h
    pad = _auto_padding(style, font) if padding == "auto" else int(padding)

    layer = Image.new("RGBA", (max(1, int(math.ceil(width)) + 2 * pad),
                               max(1, int(math.ceil(height)) + 2 * pad)), (0, 0, 0, 0))
    glyphs = []
    for i, line in enumerate(lines):
        yb = pad + ascent + i * line_h
        x = pad + (0.0 if align == "left"
                   else (width - widths[i]) / 2 if align == "center"
                   else width - widths[i])
        for ch in line:
            glyphs.append((ch, x, yb))
            x += font.getlength(ch) + spacing
    draw_glyphs(layer, glyphs, font, style)
    info = (pad, width, height, ascent, descent, line_h, len(lines))
    return layer, info


def text_image(text, font, style: Optional[Style] = None, *, spacing: float = 0.0,
               line_spacing: float = 1.0, align: str = "left", padding="auto",
               underline=None, strike: bool = False) -> Image.Image:
    """把文字渲染成一张独立的 RGBA 图片（透明底）。

    Args:
        text: 文字，``\\n`` 换行。
        font: ``ImageFont`` 字体。
        style: 样式（填充 / 阴影 / 发光 / 描边）。None = 黑色实心。
        spacing: 字距，像素（可为负）。
        line_spacing: 行距倍数，1.0 = 单倍。
        align: 多行时的对齐，``'left'`` / ``'center'`` / ``'right'``。
        padding: 四周留白。``'auto'``（默认）= 自动留够，保证阴影 / 发光 / 描边不被裁掉。
        underline: ``None`` / ``'single'`` / ``'double'``（``style`` 里没给装饰线时生效）。
        strike: 是否加删除线（同上）。

    Returns:
        RGBA 图片，文字块位于 ``(padding, padding)`` 起。
    """
    layer, _ = _layout(text, font, style, spacing, line_spacing, align, padding,
                       underline, strike)
    return layer


def draw_text(img: Image.Image, xy, text, font, style: Optional[Style] = None, *,
              anchor: str = "la", spacing: float = 0.0, line_spacing: float = 1.0,
              align: str = "left", padding="auto", underline=None,
              strike: bool = False) -> Image.Image:
    """把文字画到 ``img`` 上，越界部分自动裁剪。

    Args:
        xy: 锚点在 ``img`` 里的坐标。
        anchor: 两个字符，第一个是水平对齐 ``l`` / ``m`` / ``r``，第二个是竖直对齐：
            ``a`` 顶部、``m`` 竖直居中、``s`` 基线、``b`` 末行下沿、``d`` 整块下沿。
            默认 ``'la'`` = 左上角，和 Pillow 的 ``ImageDraw.text`` 一致。
        其余参数同 ``text_image``。

    Returns:
        入参 ``img`` 本身（就地修改，模式不变）。内部按 RGBA 合成，非 RGBA 的图片会
        先转 RGBA 再转回原模式，所以带透明度的特效在 RGB 等不含 alpha 的图上表现为
        与背景直接混合。
    """
    if len(anchor) != 2 or anchor[0] not in _ANCHOR_H or anchor[1] not in _ANCHOR_V:
        raise ValueError(
            f"anchor 需要是两个字符，水平 {_ANCHOR_H!r} + 竖直 {_ANCHOR_V!r}: {anchor!r}")
    layer, (pad, width, height, ascent, descent, line_h, n) = _layout(
        text, font, style, spacing, line_spacing, align, padding, underline, strike)
    x = xy[0] + {"l": -pad, "m": -(pad + width / 2), "r": -(pad + width)}[anchor[0]]
    y = xy[1] + {
        "a": -pad,
        "m": -(pad + height / 2),
        "s": -(pad + ascent),
        "b": -(pad + (n - 1) * line_h + ascent + descent),
        "d": -(pad + n * line_h),
    }[anchor[1]]

    out = img if img.mode == "RGBA" else img.convert("RGBA")
    ix, iy = int(round(x)), int(round(y))
    if ix < 0 or iy < 0 or ix + layer.width > out.width or iy + layer.height > out.height:
        sx0, sy0 = max(0, -ix), max(0, -iy)
        sx1 = min(layer.width, out.width - ix)
        sy1 = min(layer.height, out.height - iy)
        if sx1 <= sx0 or sy1 <= sy0:
            return img                       # 完全在图外
        layer = layer.crop((sx0, sy0, sx1, sy1))
        ix, iy = ix + sx0, iy + sy0
    out.alpha_composite(layer, (ix, iy))
    if out is not img:
        img.paste(out.convert(img.mode))
    return img
