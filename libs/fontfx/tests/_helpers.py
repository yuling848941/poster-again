# -*- coding: utf-8 -*-
"""测试公用小工具：找字体、近似比较、逐条跑用例。"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "src")))

from PIL import ImageFont  # noqa: E402

# 各平台常见的无衬线字体；都找不到时用 Pillow 自带的字体（仅拉丁字形）
_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\arial.ttf",
    "/System/Library/Fonts/Helvetica.ttc",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
]


def load_font(size):
    for path in _CANDIDATES:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:          # Pillow < 10.1
        return ImageFont.load_default()


def near(a, b, tol=2):
    """逐分量近似相等（RGB 或 RGBA）"""
    return len(a) == len(b) and all(abs(x - y) <= tol for x, y in zip(a, b))


def ink(img):
    """alpha > 0 的像素坐标集合"""
    alpha = img.getchannel("A") if img.mode == "RGBA" else None
    if alpha is None:
        raise ValueError("ink() 需要 RGBA 图片")
    return {(x, y) for y in range(img.height) for x in range(img.width)
            if alpha.getpixel((x, y)) > 0}


def border_transparent(img):
    a = img.getchannel("A")
    w, h = img.size
    return all(a.getpixel(p) == 0 for p in
               [(x, 0) for x in range(w)] + [(x, h - 1) for x in range(w)] +
               [(0, y) for y in range(h)] + [(w - 1, y) for y in range(h)])


def run(cases, title):
    """cases: [(名称, 零参函数)]。返回退出码：有失败则 1。"""
    print(f"\n【{title}】" + "=" * (56 - len(title) * 2))
    failed = 0
    for name, fn in cases:
        try:
            fn()
            print(f"  ✅ {name}")
        except Exception as e:  # noqa: BLE001 - 测试框架
            failed += 1
            print(f"  ❌ {name}: {type(e).__name__}: {e}")
    print("=" * 60)
    print(f"结果: {len(cases) - failed} 通过, {failed} 失败")
    return 1 if failed else 0
