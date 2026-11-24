# -*- coding: utf-8 -*-
"""draw_glyphs：纯色、阴影、发光、描边、图片填充、渐变填充、合成斜体、装饰线、叠放顺序。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _helpers import ink, load_font, run  # noqa: E402

from PIL import Image  # noqa: E402

import fontfx as fx  # noqa: E402

RED, GREEN, BLUE, WHITE, BLACK = (255, 0, 0), (0, 128, 0), (0, 0, 255), (255, 255, 255), (0, 0, 0)
X, BASE = 40, 140


def _layer(w=260, h=200):
    return Image.new("RGBA", (w, h), (0, 0, 0, 0))


def _colors(layer):
    """{(r, g, b): {坐标}} —— 只看完全不透明的像素"""
    out = {}
    px = layer.load()
    for y in range(layer.height):
        for x in range(layer.width):
            r, g, b, a = px[x, y]
            if a == 255:
                out.setdefault((r, g, b), set()).add((x, y))
    return out


def _plain(font, ch="H"):
    """不带任何特效的字形：(图层, 完全不透明的像素集合)"""
    layer = fx.draw_glyphs(_layer(), [(ch, X, BASE)], font)
    return layer, _colors(layer).get(BLACK, set())


def test_default_is_opaque_black_solid():
    layer = fx.draw_glyphs(_layer(), [("H", X, BASE)], load_font(80))
    px = layer.load()
    seen = {px[p][:3] for p in ink(layer)}
    assert seen == {BLACK}, seen


def test_solid_fill_and_alpha():
    layer = fx.draw_glyphs(_layer(), [("H", X, BASE)], load_font(80),
                           fx.Style(fill=fx.SolidFill(RED, alpha=0.5)))
    alphas = [layer.getchannel("A").getpixel(p) for p in ink(layer)]
    assert alphas and max(alphas) <= 129, max(alphas)
    assert {layer.getpixel(p)[:3] for p in ink(layer)} == {RED}


def test_shadow_is_shifted_copy():
    font = load_font(80)
    _, plain = _plain(font)
    dx = 90
    style = fx.Style(fill=fx.SolidFill(BLACK),
                     shadow=fx.Shadow(BLUE, alpha=1.0, blur=0.0, offset=(dx, 0)))
    colors = _colors(fx.draw_glyphs(_layer(), [("H", X, BASE)], font, style))
    assert colors.get(BLACK) == plain
    assert colors.get(BLUE) == {(x + dx, y) for x, y in plain}, "阴影应是字形整体平移一份"


def test_shadow_blur_softens_edges():
    font = load_font(80)
    counts = []
    for blur in (0.0, 5.0):
        layer = fx.draw_glyphs(_layer(), [("H", X, BASE)], font, fx.Style(
            fill=fx.SolidFill(BLACK), shadow=fx.Shadow(BLUE, alpha=1.0, blur=blur, offset=(90, 0))))
        alpha = layer.getchannel("A")
        counts.append(sum(1 for p in ink(layer) if 0 < alpha.getpixel(p) < 255))
    assert counts[1] > counts[0] + 20, counts


def test_glow_adds_light_around_glyphs():
    font = load_font(80)
    plain_layer, plain = _plain(font)
    base = _layer()
    out = fx.draw_glyphs(base, [("H", X, BASE)], font, fx.Style(
        fill=fx.SolidFill(BLACK), glow=fx.Glow(WHITE, radius=12)))
    assert out is base, "发光也必须就地修改、返回入参本身"
    extra = ink(out) - ink(plain_layer)        # 纯光晕像素（不含字形自身的抗锯齿边缘）
    assert len(extra) > 50, len(extra)
    for p in plain:                            # 字形本身仍是黑色
        assert out.getpixel(p)[:3] == BLACK, p
    assert all(out.getpixel(p)[:3] == WHITE for p in extra)


def test_stroke_wraps_glyph():
    font = load_font(80)
    _, plain = _plain(font)
    colors = _colors(fx.draw_glyphs(_layer(), [("H", X, BASE)], font, fx.Style(
        fill=fx.SolidFill(BLACK), stroke=fx.Stroke(WHITE, width=4))))
    assert colors.get(BLACK), "描边不应吃掉字形内部"
    ring = colors.get(WHITE, set())
    assert len(ring) > 50, len(ring)
    assert all(p not in plain for p in ring), "描边应落在字形轮廓之外"
    # 描边必须贴着字形（左基线锚点）：不能整体上移/下沉
    top = min(y for _, y in plain) - 7
    bottom = max(y for _, y in plain) + 7
    assert all(top <= y <= bottom for _, y in ring), "描边位置与字形不贴合"


def _half_half_image():
    img = Image.new("RGBA", (100, 100), (*RED, 255))
    img.paste((*GREEN, 255), (50, 0, 100, 100))
    return img


def _ink_colors(layer):
    return {layer.getpixel(p)[:3] for p in ink(layer)}


def test_image_fill_inside_frame():
    layer = fx.draw_glyphs(_layer(), [("I", 60, BASE)], load_font(80), fx.Style(
        fill=fx.ImageFill(_half_half_image(), frame=(0, 0, 100, 100))))
    assert _ink_colors(layer) == {GREEN}, _ink_colors(layer)


def test_image_fill_extends_edge_outside_frame():
    layer = fx.draw_glyphs(_layer(), [("I", 140, BASE)], load_font(80), fx.Style(
        fill=fx.ImageFill(_half_half_image(), frame=(0, 0, 100, 100))))
    assert _ink_colors(layer) == {GREEN}, "框外应取图片边缘（右半边）颜色"


def test_image_fill_resizes_to_frame_and_is_lazy():
    calls = []

    def loader():
        calls.append(1)
        return Image.new("RGBA", (7, 5), (*BLUE, 255))

    font = load_font(80)
    empty = _layer()
    fx.draw_glyphs(empty, [], font, fx.Style(fill=fx.ImageFill(loader, frame=(0, 0, 50, 50))))
    assert calls == [], "没有字可画时不应该去取图片"
    layer = fx.draw_glyphs(_layer(), [("I", 60, BASE)], font, fx.Style(
        fill=fx.ImageFill(loader, frame=(0, 0, 200, 200))))
    assert calls == [1]
    assert _ink_colors(layer) == {BLUE}


def test_gradient_fill_follows_frame():
    font = load_font(80)
    _, plain = _plain(font)
    row = sorted(x for x, y in plain if y == BASE - 40)     # 两条竖笔在该行的像素
    assert row, "该行应有字形像素"
    layer = fx.draw_glyphs(_layer(), [("H", X, BASE)], font, fx.Style(
        fill=fx.LinearGradient([(0.0, RED), (1.0, BLUE)], angle=0.0, frame=(0, 0, 260, 200))))
    first = layer.getpixel((row[0], BASE - 40))
    last = layer.getpixel((row[-1], BASE - 40))
    # 整条渐变横跨 260px 的参照框，字形只占左半边：左红右蓝，单调过渡
    assert first[0] > last[0] + 20 and last[2] > first[2] + 20, (first, last)


def test_synthetic_italic_slants_to_the_right():
    font = load_font(80)
    normal = ink(fx.draw_glyphs(_layer(), [("H", X, BASE)], font))
    italic = ink(fx.draw_glyphs(_layer(), [("H", X, BASE)], font, fx.Style(synthetic_italic=True)))
    top = min(y for _, y in normal)
    right = lambda pts: max(x for x, y in pts if y <= top + 4)  # noqa: E731
    assert right(italic) > right(normal) + 5, (right(normal), right(italic))


def _band(layer, x, y0, y1):
    return [layer.getchannel("A").getpixel((x, y)) for y in range(y0, y1 + 1)]


def test_underline_single_and_double():
    font = load_font(80)
    single = fx.draw_glyphs(_layer(), [(".", X, BASE)], font, fx.Style(
        fill=fx.SolidFill(BLACK), decorations=(fx.Decoration("underline", 3, 4),)))
    # 横条含首尾：offset=3、thickness=4 -> 基线下方第 3~7 行
    assert _band(single, X + 5, BASE + 3, BASE + 7) == [255] * 5
    assert _band(single, X + 5, BASE + 8, BASE + 20) == [0] * 13
    double = fx.draw_glyphs(_layer(), [(".", X, BASE)], font, fx.Style(
        fill=fx.SolidFill(BLACK), decorations=(fx.Decoration("underline", 3, 4, double=True),)))
    assert _band(double, X + 5, BASE + 3, BASE + 7) == [255] * 5
    assert _band(double, X + 5, BASE + 11, BASE + 15) == [255] * 5
    assert _band(double, X + 5, BASE + 8, BASE + 10) == [0] * 3, "双线之间应有空隙"


def test_strike_above_baseline():
    font = load_font(80)
    layer = fx.draw_glyphs(_layer(), [(".", X, BASE)], font, fx.Style(
        fill=fx.SolidFill(BLACK), decorations=(fx.Decoration("strike", -30, 4),)))
    assert _band(layer, X + 5, BASE - 30, BASE - 26) == [255] * 5


def test_stacking_order_shadow_stroke_fill():
    font = load_font(80)
    colors = _colors(fx.draw_glyphs(_layer(), [("H", X, BASE)], font, fx.Style(
        fill=fx.SolidFill((0, 255, 0)),
        shadow=fx.Shadow(RED, alpha=1.0, blur=0.0, offset=(40, 40)),
        stroke=fx.Stroke(BLUE, width=3))))
    assert colors.get((0, 255, 0)), "填充在最上面"
    assert colors.get(BLUE), "描边在填充下面"
    assert colors.get(RED), "阴影在最下面"


def test_nothing_to_draw_returns_layer_unchanged():
    font = load_font(80)
    layer = _layer()
    assert fx.draw_glyphs(layer, [], font) is layer
    assert fx.draw_glyphs(layer, [(" ", X, BASE)], font) is layer
    assert fx.draw_glyphs(layer, [("H", X, 500)], font) is layer     # 基线在图层外
    assert not ink(layer)


def test_layer_must_be_rgba():
    try:
        fx.draw_glyphs(Image.new("RGB", (10, 10)), [("H", 1, 8)], load_font(12))
    except ValueError:
        return
    raise AssertionError("应当抛 ValueError")


def test_accepts_glyph_namedtuples():
    layer = fx.draw_glyphs(_layer(), [fx.Glyph("H", X, BASE)], load_font(80))
    assert ink(layer)


if __name__ == "__main__":
    cases = [(k, v) for k, v in sorted(globals().items()) if k.startswith("test_")]
    sys.exit(run(cases, "特效合成"))
