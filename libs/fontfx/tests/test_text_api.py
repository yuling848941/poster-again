# -*- coding: utf-8 -*-
"""text_image / draw_text / measure：排版、锚点、余量、就地修改、越界裁剪、颜色解析。"""

import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _helpers import border_transparent, ink, load_font, run  # noqa: E402

from PIL import Image  # noqa: E402

import fontfx as fx  # noqa: E402

BG = (10, 20, 30, 255)


def _bbox(pts):
    xs = [x for x, _ in pts]
    ys = [y for _, y in pts]
    return min(xs), min(ys), max(xs), max(ys)


def test_text_image_tight_matches_measure():
    font = load_font(72)
    img = fx.text_image("Ag", font, padding=0)
    w, h = fx.measure("Ag", font)
    assert img.mode == "RGBA"
    assert img.size == (math.ceil(w), math.ceil(h)), (img.size, w, h)
    assert ink(img)


def test_effects_are_not_clipped():
    font = load_font(72)
    for name, style in (
            ("发光", fx.Style(glow=fx.Glow("#FF3DCF", alpha=0.9, radius=20))),
            ("阴影", fx.Style(shadow=fx.Shadow("black", alpha=0.8, blur=8, offset=(15, 20))),
                             ),
            ("描边", fx.Style(stroke=fx.Stroke("white", width=6))),
            ("全部", fx.Style(fill=fx.SolidFill("white"),
                              glow=fx.Glow("#00E5FF", radius=14),
                              shadow=fx.Shadow("black", alpha=0.7, blur=5, offset=(-9, 12)),
                              stroke=fx.Stroke("#D81E5B", width=3),
                              synthetic_italic=True,
                              decorations=fx.default_decorations(font, "double", True)))):
        img = fx.text_image("Hgply\n第二行", font, style, line_spacing=1.3)
        assert border_transparent(img), f"{name} 被裁掉了"
        assert ink(img)


def test_draw_text_modifies_rgba_in_place():
    font = load_font(72)
    img = Image.new("RGBA", (300, 200), BG)
    out = fx.draw_text(img, (50, 50), "Ag", font, fx.Style(fill=fx.SolidFill("white")))
    assert out is img
    changed = {(x, y) for y in range(200) for x in range(300) if img.getpixel((x, y)) != BG}
    assert changed, "应有像素被修改"
    x0, y0, x1, y1 = _bbox(changed)
    # 文本框左上角在 (50, 50)；字形离顶端还有上升部的空余
    assert 50 <= x0 <= 60 and 50 <= y0 < 100, (x0, y0)
    assert x1 < 260 and y1 < 170, (x1, y1)
    assert img.getpixel((0, 0)) == BG and img.getpixel((299, 199)) == BG


def test_draw_text_modifies_rgb_in_place():
    font = load_font(72)
    img = Image.new("RGB", (300, 200), BG[:3])
    out = fx.draw_text(img, (50, 50), "Ag", font, fx.Style(fill=fx.SolidFill("white")))
    assert out is img and img.mode == "RGB"
    assert any(img.getpixel((x, y)) == (255, 255, 255)
               for y in range(200) for x in range(300)), "不透明白字应与背景色完全不同"


def test_anchor_right_and_baseline():
    font = load_font(72)
    img = Image.new("RGBA", (300, 200), (0, 0, 0, 0))
    fx.draw_text(img, (250, 150), "H", font, fx.Style(fill=fx.SolidFill("white")), anchor="rs")
    x0, y0, x1, y1 = _bbox(ink(img))
    assert 240 <= x1 <= 251, x1            # 右对齐（字宽取整）
    assert 149 <= y1 <= 151, y1            # 基线：H 的底边贴在基线上
    assert y0 >= 150 - 80, y0


def test_anchor_variants_render():
    font = load_font(72)
    for anchor in ("la", "ms", "rd"):
        img = Image.new("RGBA", (300, 200), (0, 0, 0, 0))
        fx.draw_text(img, (150, 100), "H", font, anchor=anchor)
        assert ink(img), anchor


def test_multiline_and_center_align():
    font = load_font(72)
    img = fx.text_image("A\nBBBBB", font, padding=0, align="center")
    half = img.height // 2
    second = {(x, y) for x, y in ink(img) if y >= half}
    first = {(x, y) for x, y in ink(img) if y < half}
    assert first and second
    cx = sum(x for x, _ in second) / len(second)
    assert abs(cx - img.width / 2) <= 4, cx
    assert fx.text_image("A\nBBB", font, padding=0, align="right").width == \
        fx.text_image("A\nBBB", font, padding=0, align="left").width


def test_out_of_bounds_is_clipped_not_an_error():
    font = load_font(72)
    img = Image.new("RGBA", (100, 100), BG)
    fx.draw_text(img, (-500, -500), "Ag", font)
    assert all(img.getpixel((x, y)) == BG for y in range(100) for x in range(100))
    fx.draw_text(img, (95, 95), "Ag", font, fx.Style(fill=fx.SolidFill("white")))
    assert img.getpixel((0, 0)) == BG
    fx.draw_text(img, (40, 200), "Ag", font)            # 整体在图下方


def test_spacing_widens_text():
    font = load_font(72)
    plain = fx.measure("AB", font)[0]
    assert abs(fx.measure("AB", font, spacing=12)[0] - plain - 12) < 1e-6
    assert fx.text_image("AB", font, spacing=40).width > fx.text_image("AB", font).width + 30


def test_line_spacing_height():
    font = load_font(72)
    one = fx.text_image("A", font, padding=0).height
    two = fx.text_image("A\nA", font, padding=0, line_spacing=1.5).height
    assert abs(two - one * 3) <= 2, (one, two)


def test_underline_and_strike_shortcuts():
    font = load_font(72)
    ascent, descent = font.getmetrics()
    plain = ink(fx.text_image("A", font, padding=0))

    single = ink(fx.text_image("A", font, padding=0, underline="single")) - plain
    assert single and min(y for _, y in single) >= ascent + 5, "下划线应在基线下方"
    assert max(y for _, y in single) <= ascent + 10

    double = ink(fx.text_image("A", font, padding=0, underline="double")) - plain
    assert double and max(y for _, y in double) >= min(y for _, y in double) + 8, "双线应有间隔"

    strike = ink(fx.text_image("A", font, padding=0, strike=True)) - plain
    center = ascent - 0.245 * font.size
    assert strike and all(abs(y - center) <= 5 for _, y in strike), "删除线应横穿字身上部"


def test_decorations_follow_fill_color():
    font = load_font(72)
    img = fx.text_image("A", font, fx.Style(fill=fx.SolidFill("red")), padding=0, underline="single")
    assert {img.getpixel(p)[:3] for p in ink(img)} == {(255, 0, 0)}, "装饰线应与文字同色"


def test_colors_and_alpha():
    assert fx.to_rgb("#FF8000") == (255, 128, 0)
    assert fx.to_rgb("white") == (255, 255, 255)
    assert fx.SolidFill("#FF000080").alpha == 128 / 255
    assert abs(fx.SolidFill((1, 2, 3, 51)).alpha - 0.2) < 1e-6
    assert fx.SolidFill((1, 2, 3)).alpha == 1.0
    s = fx.SolidFill((1, 2, 3, 51), alpha=0.9)         # 显式 alpha 优先
    assert s.alpha == 0.9 and s.color == (1, 2, 3)
    assert fx.Shadow("black").alpha == 1.0
    img = fx.text_image("A", load_font(72), fx.Style(fill=fx.SolidFill("red", alpha=0.5)), padding=0)
    assert max(img.getchannel("A").getdata()) <= 128


def test_validation():
    font = load_font(20)
    img = Image.new("RGBA", (50, 50))
    for bad in (lambda: fx.draw_text(img, (1, 1), "A", font, anchor="x"),
                lambda: fx.draw_text(img, (1, 1), "A", font, anchor="lz"),
                lambda: fx.draw_text(img, (1, 1), "A", font, align="justify"),
                lambda: fx.Decoration("dotted", 1, 1),
                lambda: fx.Stroke("black", width=-1),
                lambda: fx.Glow("white", radius=-1),
                lambda: fx.SolidFill((1, 2))):
        try:
            bad()
        except ValueError:
            continue
        raise AssertionError("应当抛 ValueError")


if __name__ == "__main__":
    cases = [(k, v) for k, v in sorted(globals().items()) if k.startswith("test_")]
    sys.exit(run(cases, "文字接口"))
