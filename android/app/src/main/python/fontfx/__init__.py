# -*- coding: utf-8 -*-
"""fontfx —— 纯 Pillow 的文字特效库：渐变字、图片填充字、发光字、阴影字、描边字。

只依赖 Pillow；渐变在装有 numpy 时走向量化实现，没装也能跑。
横向单字体排版；字形逐个绘制，不依赖 HarfBuzz/Raqm。
"""

from .core import (
    GLOW_GAIN,
    GLOW_GAMMA,
    GLOW_SIGMA_KS,
    ITALIC_SHEAR,
    draw_glyphs,
    draw_italic_glyph,
    linear_gradient,
)
from .styles import (
    Decoration,
    Frame,
    Glyph,
    Glow,
    ImageFill,
    LinearGradient,
    Shadow,
    SolidFill,
    Stop,
    Stroke,
    Style,
    default_decorations,
    normalize_stops,
    to_rgb,
)
from .text import draw_text, measure, text_image

__version__ = "0.1.0"

__all__ = [
    "GLOW_GAIN", "GLOW_GAMMA", "GLOW_SIGMA_KS", "ITALIC_SHEAR",
    "Decoration", "Frame", "Glyph", "Glow", "ImageFill", "LinearGradient",
    "Shadow", "SolidFill", "Stop", "Stroke", "Style",
    "default_decorations", "normalize_stops", "to_rgb",
    "draw_glyphs", "draw_italic_glyph", "linear_gradient",
    "draw_text", "measure", "text_image",
    "__version__",
]
