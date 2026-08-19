# -*- coding: utf-8 -*-
"""从 Stitch 图标稿生成安卓自适应图标资源。
用法: .venv/Scripts/python.exe android/make_icons.py <1024图标.png>

产出（写入 android/app/src/main/res/）：
- mipmap-*/ic_launcher.png / ic_launcher_round.png   —— API<26 兼容位图
- mipmap-*/ic_launcher_fg.png                        —— 自适应前景层（全出血）
- mipmap-anydpi-v26/ic_launcher.xml + ic_launcher_round.xml + colors（background 纯色）
"""
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageOps

RES = Path(__file__).resolve().parent / "app" / "src" / "main" / "res"
# {dpi: (legacy_px, foreground_px)}   legacy=48dp, foreground=108dp
DENSITIES = {
    "mdpi": (48, 108), "hdpi": (72, 162), "xhdpi": (96, 216),
    "xxhdpi": (144, 324), "xxxhdpi": (192, 432),
}
TRIM = 0.045          # 裁掉外圈白描边
BG_HEX = "#02698A"    # background 兜底色（取渐变右下角）


def rounded(im, radius_ratio):
    """圆角矩形裁切（alpha 蒙版）"""
    w, h = im.size
    r = int(min(w, h) * radius_ratio)
    mask = Image.new("L", (w, h), 0)
    d = ImageDraw.Draw(mask)
    d.rounded_rectangle([0, 0, w - 1, h - 1], radius=r, fill=255)
    out = im.convert("RGBA")
    out.putalpha(mask)
    return out


def circle(im):
    w, h = im.size
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, w - 1, h - 1], fill=255)
    out = im.convert("RGBA")
    out.putalpha(mask)
    return out


def main(src_path):
    src = Image.open(src_path).convert("RGB")
    w, h = src.size
    t = int(w * TRIM)
    core = src.crop((t, t, w - t, h - t))          # 去白边，全出血渐变+卡片

    for dpi, (legacy_px, fg_px) in DENSITIES.items():
        out = RES / f"mipmap-{dpi}"
        out.mkdir(exist_ok=True)

        fg = core.resize((fg_px, fg_px), Image.LANCZOS)
        fg.save(out / "ic_launcher_fg.png")

        # 兼容位图：整图设计本身就是全出血渐变+居中卡片，直接圆角/圆形
        full = core.resize((legacy_px, legacy_px), Image.LANCZOS)
        rounded(full, 0.22).save(out / "ic_launcher.png")
        circle(full).save(out / "ic_launcher_round.png")
        print(f"mipmap-{dpi}: legacy {legacy_px}px, fg {fg_px}px")

    # anydpi-v26 自适应图标定义
    anydpi = RES / "mipmap-anydpi-v26"
    anydpi.mkdir(exist_ok=True)
    for name in ("ic_launcher", "ic_launcher_round"):
        (anydpi / f"{name}.xml").write_text(
            f'<?xml version="1.0" encoding="utf-8"?>\n'
            f'<adaptive-icon xmlns:android="http://schemas.android.com/apk/res/android">\n'
            f'    <background android:drawable="@color/ic_launcher_bg"/>\n'
            f'    <foreground android:drawable="@mipmap/ic_launcher_fg"/>\n'
            f'</adaptive-icon>\n', encoding="utf-8")
    values = RES / "values"
    values.mkdir(exist_ok=True)
    colors = values / "ic_launcher_colors.xml"
    if colors.exists():
        colors.unlink()
    colors.write_text(
        f'<?xml version="1.0" encoding="utf-8"?>\n'
        f'<resources>\n'
        f'    <color name="ic_launcher_bg">{BG_HEX}</color>\n'
        f'</resources>\n', encoding="utf-8")
    print("anydpi-v26 adaptive icon xml + colors OK")
    print("done")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(
        Path(__file__).resolve().parent / "art" / "icon_1024.png"))
