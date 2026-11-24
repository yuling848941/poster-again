# -*- coding: utf-8 -*-
"""线性渐变：方向、端点色、参照框、透明度、有无 numpy 两条实现的一致性。"""

import builtins
import os
import sys
from itertools import product

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _helpers import near, run  # noqa: E402

import fontfx as fx  # noqa: E402

RED, GREEN, BLUE = (255, 0, 0), (0, 255, 0), (0, 0, 255)
TWO = [(0.0, RED, 1.0), (1.0, BLUE, 1.0)]


class _NoNumpy:
    """让 `import numpy` 抛 ImportError，逼 linear_gradient 走纯 Pillow 路径"""

    def __enter__(self):
        self._orig = builtins.__import__

        def fake_import(name, *args, **kwargs):
            if name == "numpy":
                raise ImportError("numpy 被测试屏蔽")
            return self._orig(name, *args, **kwargs)

        builtins.__import__ = fake_import
        return self

    def __exit__(self, *exc):
        builtins.__import__ = self._orig


def _max_diff(a, b):
    assert a.size == b.size and a.mode == b.mode
    pa, pb = a.load(), b.load()
    return max(abs(x - y)
               for pos in product(range(a.width), range(a.height))
               for x, y in zip(pa[pos], pb[pos]))


def test_left_to_right():
    g = fx.linear_gradient((101, 8), TWO, angle=0)
    assert g.size == (101, 8) and g.mode == "RGBA"
    assert near(g.getpixel((0, 3)), (255, 0, 0, 255), 1), g.getpixel((0, 3))
    assert near(g.getpixel((100, 3)), (0, 0, 255, 255), 1), g.getpixel((100, 3))
    assert near(g.getpixel((50, 3)), (128, 0, 128, 255), 2), g.getpixel((50, 3))


def test_top_to_bottom():
    g = fx.linear_gradient((8, 101), TWO, angle=90)
    assert near(g.getpixel((3, 0))[:3], RED, 1)
    assert near(g.getpixel((3, 100))[:3], BLUE, 1)
    assert near(g.getpixel((3, 50))[:3], (128, 0, 128), 2)


def test_diagonal():
    g = fx.linear_gradient((101, 101), TWO, angle=45)
    assert near(g.getpixel((0, 0))[:3], RED, 1)
    assert near(g.getpixel((100, 100))[:3], BLUE, 1)


def test_reverse_angle():
    g = fx.linear_gradient((101, 8), TWO, angle=180)
    assert near(g.getpixel((0, 3))[:3], BLUE, 1)
    assert near(g.getpixel((100, 3))[:3], RED, 1)


def test_three_stops():
    g = fx.linear_gradient((101, 4), [(0.0, RED), (0.5, GREEN), (1.0, BLUE)], angle=0)
    assert near(g.getpixel((25, 0))[:3], (128, 128, 0), 2), g.getpixel((25, 0))
    assert near(g.getpixel((50, 0))[:3], GREEN, 1)
    assert near(g.getpixel((75, 0))[:3], (0, 128, 128), 2), g.getpixel((75, 0))


def test_extent_and_origin():
    """参照框比渐变图大：渐变图是整条渐变中的一段"""
    full = fx.linear_gradient((100, 4), TWO, angle=0)
    part = fx.linear_gradient((50, 4), TWO, angle=0, extent=(100, 100), origin=(-25, 0))
    for x in (0, 10, 25, 49):
        assert near(part.getpixel((x, 0)), full.getpixel((x + 25, 0)), 2), x


def test_outside_extent_takes_endpoint_colors():
    g = fx.linear_gradient((100, 4), TWO, angle=0, extent=(40, 4), origin=(30, 0))
    assert near(g.getpixel((0, 0))[:3], RED, 1)
    assert near(g.getpixel((29, 0))[:3], RED, 2)
    assert near(g.getpixel((70, 0))[:3], BLUE, 2)
    assert near(g.getpixel((99, 0))[:3], BLUE, 1)
    # 框宽 40 落在 x=30..69：x=49 对应 t=19/39
    assert near(g.getpixel((49, 0))[:3], (131, 0, 124), 3), g.getpixel((49, 0))


def test_alpha_channel():
    g = fx.linear_gradient((101, 2), [(0.0, (0, 0, 0), 0.5), (1.0, (0, 0, 0), 1.0)], angle=0)
    assert abs(g.getpixel((0, 0))[3] - 128) <= 2, g.getpixel((0, 0))
    assert abs(g.getpixel((100, 0))[3] - 255) <= 1


def test_hex_color_with_alpha_in_stops():
    g = fx.linear_gradient((11, 2), [(0, "#FF000080"), (1, "#00FF00")], angle=0)
    px = g.getpixel((0, 0))
    assert near(px, (255, 0, 0, 128), 1), px


def test_single_stop_is_uniform():
    g = fx.linear_gradient((30, 30), [(0.3, (12, 34, 56), 0.4)], angle=37)
    assert _max_diff(g, g.crop((0, 0, 1, 1)).resize(g.size)) <= 1


def test_pure_path_matches_numpy_path():
    try:
        import numpy  # noqa: F401
    except ImportError:
        print("      (未安装 numpy，跳过两条实现的对比)")
        return
    cases = [((120, 60), TWO, 0.0, None, (0.0, 0.0)),
             ((80, 80), TWO, 90.0, None, (0.0, 0.0)),
             ((90, 40), [(0.0, RED), (0.5, GREEN), (1.0, BLUE)], 315.0, None, (0.0, 0.0)),
             ((40, 40), [(0.0, (9, 9, 9), 0.3), (1.0, (250, 200, 100), 0.9)], 47.0,
              (80, 30), (-12.0, 7.0))]
    for size, stops, angle, extent, origin in cases:
        with_numpy = fx.linear_gradient(size, stops, angle, extent, origin)
        with _NoNumpy():
            pure = fx.linear_gradient(size, stops, angle, extent, origin)
        diff = _max_diff(with_numpy, pure)
        assert diff <= 3, (size, angle, diff)


def test_invalid_arguments():
    for bad in (lambda: fx.linear_gradient((10, 10), [], 0),
                lambda: fx.linear_gradient((0, 10), TWO, 0),
                lambda: fx.linear_gradient((10, -1), TWO, 0),
                lambda: fx.LinearGradient([])):
        try:
            bad()
        except ValueError:
            continue
        raise AssertionError("应当抛 ValueError")


if __name__ == "__main__":
    cases = [(k, v) for k, v in sorted(globals().items()) if k.startswith("test_")]
    sys.exit(run(cases, "渐变"))
