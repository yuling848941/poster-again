# -*- coding: utf-8 -*-
"""样式定义：填充 / 阴影 / 发光 / 描边 / 装饰线。

约定：
- 颜色接受 ``'#RRGGBB'``、``'#RRGGBBAA'``、``(r, g, b)``、``(r, g, b, a)`` 以及 Pillow
  颜色名（``'red'``、``'white'``……），内部统一成 ``(r, g, b)``；透明度用 ``alpha``
  （0.0~1.0）单独表示。颜色自带透明度而 ``alpha`` 留空时，取颜色里的透明度。
- 所有长度单位都是像素：最终图按什么分辨率渲染，就按那个分辨率的像素给。
"""

from dataclasses import dataclass
from typing import Callable, NamedTuple, Optional, Sequence, Tuple, Union

from PIL import Image, ImageColor

RGB = Tuple[int, int, int]
Color = Union[str, Sequence[int]]
# 参照框：(原点 x, 原点 y, 宽, 高)，坐标属于最终图层
Frame = Tuple[float, float, float, float]
# 渐变 stop：(位置 0~1, 颜色[, 透明度])
Stop = Tuple[float, Color, float]


def _resolve(color, alpha) -> Tuple[RGB, float]:
    """(颜色, 透明度) -> ((r, g, b), 透明度)。透明度留空时用颜色自带的，否则为 1.0。"""
    if isinstance(color, str):
        r, g, b, a = ImageColor.getcolor(color, "RGBA")
        embedded = None if (len(color) == 7 and color.startswith("#")) else a / 255.0
    else:
        try:
            vals = tuple(color)
        except TypeError:
            raise ValueError(f"不支持的颜色: {color!r}") from None
        if len(vals) not in (3, 4):
            raise ValueError(f"颜色需要是 (r, g, b) 或 (r, g, b, a): {color!r}")
        r, g, b = vals[:3]
        embedded = vals[3] / 255.0 if len(vals) == 4 else None
    final = (1.0 if embedded is None else embedded) if alpha is None else float(alpha)
    return (int(r), int(g), int(b)), final


def to_rgb(color) -> RGB:
    """颜色 -> (r, g, b)（丢弃自带透明度）"""
    return _resolve(color, None)[0]


def normalize_stops(stops) -> Tuple[Stop, ...]:
    """渐变 stop 统一成 (位置, (r, g, b), 透明度)。已经是这个形式的原样保留数值。"""
    out = []
    for stop in stops:
        if len(stop) not in (2, 3):
            raise ValueError(f"stop 需要 (位置, 颜色[, 透明度]): {stop!r}")
        rgb, alpha = _resolve(stop[1], stop[2] if len(stop) == 3 else None)
        out.append((float(stop[0]), rgb, alpha))
    return tuple(out)


class Glyph(NamedTuple):
    """一个字形：字符、左基线锚点的 x、基线的 y（都相对最终图层，浮点像素）"""
    ch: str
    x: float
    baseline: float


# ---------------------------------------------------------------- 填充

@dataclass(frozen=True)
class SolidFill:
    """纯色填充"""
    color: Color = (0, 0, 0)
    alpha: Optional[float] = None

    def __post_init__(self):
        rgb, alpha = _resolve(self.color, self.alpha)
        object.__setattr__(self, "color", rgb)
        object.__setattr__(self, "alpha", alpha)


@dataclass(frozen=True)
class LinearGradient:
    """线性渐变填充。

    - ``angle``：方向角，单位度。0° = 左→右，90° = 上→下，45° = 左上→右下。
    - ``frame``：可选参照框 ``(x, y, 宽, 高)``，坐标属于最终图层。渐变按这个框归一化，
      框外的部分取两端颜色。多行文字想让渐变跨越整个文本框而不是每行各来一遍时，
      把 ``frame`` 设成文本框即可。留空 = 以渐变图自身为参照。
    - 参照框的原点取整到像素。
    """
    stops: Sequence[Stop]
    angle: float = 0.0
    frame: Optional[Frame] = None

    def __post_init__(self):
        if not self.stops:
            raise ValueError("渐变至少需要一个 stop")
        object.__setattr__(self, "stops", normalize_stops(self.stops))


@dataclass(frozen=True)
class ImageFill:
    """图片填充：图片拉伸铺满 ``frame``，文字超出 ``frame`` 的部分取图片边缘颜色延伸。

    ``image`` 可以是 ``PIL.Image.Image``（任意模式，内部转 RGBA），也可以是一个
    返回图片的零参函数——图片只在真正有字可画时才会去取。
    """
    image: Union[Image.Image, Callable[[], Image.Image]]
    frame: Frame


# ---------------------------------------------------------------- 效果

@dataclass(frozen=True)
class Shadow:
    """外阴影"""
    color: Color = (0, 0, 0)
    blur: float = 0.0                          # 高斯模糊 sigma（像素）；0 = 硬边
    offset: Tuple[float, float] = (0.0, 0.0)   # (dx, dy)，正值向右下
    alpha: Optional[float] = None

    def __post_init__(self):
        rgb, alpha = _resolve(self.color, self.alpha)
        object.__setattr__(self, "color", rgb)
        object.__setattr__(self, "alpha", alpha)


@dataclass(frozen=True)
class Glow:
    """外发光：字形向外扩散的光晕，叠在已有内容上做滤色（加亮），不会把背景压暗。

    观感按 PowerPoint 的文字发光校准（红/橙/黄/绿/蓝五色对照）。
    """
    color: Color
    radius: float                              # 发光半径（像素）
    alpha: Optional[float] = None

    def __post_init__(self):
        if self.radius < 0:
            raise ValueError("发光半径不能为负")
        rgb, alpha = _resolve(self.color, self.alpha)
        object.__setattr__(self, "color", rgb)
        object.__setattr__(self, "alpha", alpha)


@dataclass(frozen=True)
class Stroke:
    """描边"""
    color: Color = (0, 0, 0)
    width: int = 1                             # 单侧向外扩展的宽度（像素）
    alpha: Optional[float] = None

    def __post_init__(self):
        if self.width < 0:
            raise ValueError("描边宽度不能为负")
        rgb, alpha = _resolve(self.color, self.alpha)
        object.__setattr__(self, "color", rgb)
        object.__setattr__(self, "alpha", alpha)


# ---------------------------------------------------------------- 装饰线

@dataclass(frozen=True)
class Decoration:
    """删除线 / 下划线。相对基线定位，随填充、阴影、描边一起着色。

    - ``kind``：``'strike'`` 或 ``'underline'``
    - ``offset``：横条上沿相对基线的纵向偏移（像素）；删除线在基线上方，给负值
    - ``thickness``：横条粗细（像素，内部至少 1）
    - ``double``：仅下划线有效，画双线
    """
    kind: str
    offset: float
    thickness: float
    double: bool = False

    def __post_init__(self):
        if self.kind not in ("strike", "underline"):
            raise ValueError(f"kind 只能是 'strike' 或 'underline': {self.kind!r}")


def default_decorations(font, underline=None, strike=False) -> Tuple[Decoration, ...]:
    """按字号给出常用的装饰线位置（无衬线黑体的经验值，正文够用）。"""
    em = font.size
    thick = max(1, int(round(0.05 * em)))
    out = []
    if strike:
        out.append(Decoration("strike", -0.245 * em, thick))
    if underline:
        out.append(Decoration("underline", 0.09 * em, thick, double=(underline == "double")))
    return tuple(out)


# ---------------------------------------------------------------- 样式集合

@dataclass(frozen=True)
class Style:
    """一个文字样式的全部外观。

    叠放顺序固定为 阴影 → 发光 → 描边 → 填充，与 PowerPoint 一致。
    ``fill`` 为 None 时按黑色实心处理。
    ``synthetic_italic``：字体文件本身没有斜体时，把字形倾斜 12° 代替。
    """
    fill: Optional[Union[SolidFill, LinearGradient, ImageFill]] = None
    shadow: Optional[Shadow] = None
    glow: Optional[Glow] = None
    stroke: Optional[Stroke] = None
    synthetic_italic: bool = False
    decorations: Tuple[Decoration, ...] = ()
