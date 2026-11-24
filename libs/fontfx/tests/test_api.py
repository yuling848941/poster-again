# -*- coding: utf-8 -*-
"""公开接口的形状：导出名、不可变样式、stop 归一化保持数值、参数校验。"""

import dataclasses
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _helpers import ink, load_font, run  # noqa: E402

from PIL import Image  # noqa: E402

import fontfx  # noqa: E402


def test_all_exports_resolve():
    for name in fontfx.__all__:
        assert getattr(fontfx, name, None) is not None, name
    assert isinstance(fontfx.__version__, str) and fontfx.__version__.count(".") == 2


def test_style_objects_are_immutable():
    s = fontfx.Style(fill=fontfx.SolidFill("red"))
    try:
        s.fill = None
    except dataclasses.FrozenInstanceError:
        return
    raise AssertionError("Style 应当不可变")


def test_normalize_stops_keeps_numbers_and_is_idempotent():
    raw = [(0.5, (1, 2, 3), 0.25), (0.75, (200, 100, 50), 1.0)]
    once = fontfx.normalize_stops(raw)
    assert once == tuple(raw), once
    assert fontfx.normalize_stops(once) == once
    hexed = fontfx.normalize_stops([(0, "#FF000080"), (1, (0, 255, 0))])
    assert hexed == ((0.0, (255, 0, 0), 128 / 255), (1.0, (0, 255, 0), 1.0)), hexed


def test_default_decorations():
    font = load_font(40)
    assert fontfx.default_decorations(font) == ()
    u = fontfx.default_decorations(font, underline="double")
    assert len(u) == 1 and u[0].kind == "underline" and u[0].double
    assert u[0].offset > 0 and fontfx.default_decorations(font, strike=True)[0].offset < 0


def test_style_shortcut_fields_match_low_level():
    """Style(fill=LinearGradient(frame=...)) 与直接调 linear_gradient + 蒙版结果一致。"""
    font = load_font(72)
    stops = [(0.0, (255, 0, 0), 1.0), (1.0, (0, 0, 255), 1.0)]
    layer = Image.new("RGBA", (200, 120), (0, 0, 0, 0))
    a = fontfx.draw_glyphs(layer, [("A", 30, 80)], font, fontfx.Style(
        fill=fontfx.LinearGradient(stops, angle=0.0, frame=(0, 0, 200, 120))))
    layer_b = Image.new("RGBA", (200, 120), (0, 0, 0, 0))
    mask = Image.new("L", (200, 120), 0)
    from PIL import ImageDraw
    ImageDraw.Draw(mask).text((30, 80), "A", font=font, fill=255, anchor="ls")
    b = fontfx.draw_glyphs(layer_b, [("A", 30, 80)], font, fontfx.Style(
        fill=fontfx.ImageFill(fontfx.linear_gradient((200, 120), stops, 0.0),
                              frame=(0, 0, 200, 120))))
    assert ink(a) == ink(b)
    pa, pb = a.load(), b.load()
    for p in ink(a):
        assert pa[p] == pb[p], p


if __name__ == "__main__":
    cases = [(k, v) for k, v in sorted(globals().items()) if k.startswith("test_")]
    sys.exit(run(cases, "公开接口"))
