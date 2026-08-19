# -*- coding: utf-8 -*-
"""从安卓端存档的图标设计稿生成 PC 端 Windows .ico（多尺寸）。
用法: .venv/Scripts/python.exe resources/make_icon_ico.py [源图.png]
默认源图: android/art/icon_1024.png（Stitch 导出稿，外圈白描边会被裁掉）
输出: resources/icon.ico（覆盖，被 PosterAgain.spec 与 main.py 引用）
"""
import sys
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
DEFAULT_SRC = HERE.parent / "android" / "art" / "icon_1024.png"
SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
TRIM = 0.045  # 裁掉 Stitch 导出稿外圈的白描边，与安卓端 make_icons.py 一致


def main(src_path):
    src = Image.open(src_path).convert("RGBA")
    w, h = src.size
    t = int(w * TRIM)
    core = src.crop((t, t, w - t, h - t))
    core.save(HERE / "icon.ico", sizes=SIZES)
    print(f"OK -> {HERE / 'icon.ico'}  sizes: {[s[0] for s in SIZES]}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(DEFAULT_SRC))
