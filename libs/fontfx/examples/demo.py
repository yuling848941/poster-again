# -*- coding: utf-8 -*-
"""生成一张包含全部特效的示例图：python examples/demo.py [输出路径]"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

from PIL import Image, ImageDraw, ImageFont  # noqa: E402

import fontfx as fx  # noqa: E402


def _font(size):
    for path in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyhbd.ttc",
                 "/System/Library/Fonts/PingFang.ttc",
                 "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"):
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default(size=size)


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "fontfx_demo.png"
    W, H = 1400, 900
    img = Image.new("RGB", (W, H))
    # 对角渐变底图，顺便展示 linear_gradient
    img.paste(fx.linear_gradient((W, H), [(0, "#2B1B4D"), (1, "#0B3550")], angle=60).convert("RGB"), (0, 0))
    cy = "#E8F1FF"
    big, mid, label = _font(150), _font(96), _font(30)

    # ---- 第一行 ----
    # 金色渐变 + 外阴影
    fx.draw_text(img, (70, 40), "恭喜发财", big, fx.Style(
        fill=fx.LinearGradient([(0.0, "#FFF7C2"), (0.45, "#FFD84A"), (1.0, "#C97A08")], angle=90),
        shadow=fx.Shadow("black", alpha=0.65, blur=10, offset=(7, 9))))
    # 霓虹发光
    fx.draw_text(img, (800, 70), "HAPPY 2026", mid, fx.Style(
        fill=fx.SolidFill("#FFFFFF"), glow=fx.Glow("#FF3DCF", alpha=0.95, radius=16)))

    # ---- 第二行 ----
    # 图片填充（整行共用一张图，边缘延伸盖住每个字）
    photo = fx.linear_gradient((420, 170), [(0, "#37E2D5"), (0.5, "#F9F871"), (1, "#FF61DC")],
                               angle=35)
    fx.draw_text(img, (70, 250), "图片填充", big, fx.Style(
        fill=fx.ImageFill(photo, frame=(70, 260, 420, 170))))
    # 描边
    fx.draw_text(img, (800, 280), "描边文字", mid, fx.Style(
        fill=fx.SolidFill(cy), stroke=fx.Stroke("#FF5A36", width=5)))

    # ---- 第三行 ----
    # 跨行对齐的一条渐变（frame 让两行共用同一条渐变）+ 双下划线 + 删除线
    fx.draw_text(img, (70, 460), "跨行渐变", mid, fx.Style(
        fill=fx.LinearGradient([(0, "#8EF6FF"), (1, "#FFD166")], angle=90, frame=(70, 460, 420, 300))),
                 line_spacing=1.15)
    fx.draw_text(img, (70, 600), "第二行", mid, fx.Style(
        fill=fx.LinearGradient([(0, "#8EF6FF"), (1, "#FFD166")], angle=90, frame=(70, 460, 420, 300))),
                 line_spacing=1.15)
    fx.draw_text(img, (600, 500), "双下划线 + 删除线", mid,
                 fx.Style(fill=fx.SolidFill(cy)), underline="double", strike=True)

    # ---- 第四行 ----
    # 合成斜体 + 半透明
    fx.draw_text(img, (70, 740), "Italic 72%", mid, fx.Style(
        fill=fx.SolidFill("#9BE564", alpha=0.72), synthetic_italic=True))
    # 右下角对齐
    fx.draw_text(img, (1330, 840), "右下角对齐", mid, fx.Style(fill=fx.SolidFill(cy)), anchor="rd")

    ImageDraw.Draw(img).text((W - 70, H - 8), "fontfx demo", font=label, fill="#7F8EA8", anchor="rd")
    img.save(out)
    print(f"已生成 {os.path.abspath(out)} ({W}x{H})")


if __name__ == "__main__":
    main()
