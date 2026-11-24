# fontfx

纯 Pillow 的文字特效库：**渐变字、图片填充字、发光字、阴影字、描边字**，一行调用。

- 只依赖 Pillow；渐变在装了 numpy 时走向量化实现，没装也能跑。
- 发光观感按 PowerPoint 的文字发光校准过（红/橙/黄/绿/蓝五色对照）。
- 不依赖 HarfBuzz/Raqm，Windows / macOS / Linux / Android（Chaquopy）都能用。

> 这个库是从 PPTX 海报渲染引擎里抽出来的底层：引擎负责"每个字该放在哪"，
> fontfx 负责"这个字长什么样"。任何用 Pillow 画字的项目都可以单独用它。

## 安装

```bash
pip install ./libs/fontfx              # 在本仓库里
pip install "fontfx[numpy]"            # 带渐变加速
```

## 快速上手

```python
from PIL import Image, ImageDraw, ImageFont
import fontfx as fx

img = Image.new("RGB", (1200, 600), "#1b2a49")
font = ImageFont.truetype("msyh.ttc", 160)

# 金色渐变 + 外阴影
gold = fx.Style(
    fill=fx.LinearGradient([(0.0, "#FFF3B0"), (0.5, "#FFD54A"), (1.0, "#C8860A")], angle=90),
    shadow=fx.Shadow(color="black", alpha=0.6, blur=10, offset=(6, 8)),
)
fx.draw_text(img, (80, 80), "恭喜发财", font, gold)

# 霓虹发光
neon = fx.Style(fill=fx.SolidFill("#FFFFFF"), glow=fx.Glow("#FF3DCF", alpha=0.9, radius=18))
fx.draw_text(img, (80, 300), "HAPPY NEW YEAR", font, neon, anchor="ls")

# 描边
outline = fx.Style(fill=fx.SolidFill("#FFFFFF"), stroke=fx.Stroke("#D81E5B", width=6))
fx.draw_text(img, (1120, 300), "2026", font, outline, anchor="rs")

img.convert("RGB").save("demo.png")
```

跑 `python examples/demo.py` 会生成一张包含全部特效的示例图。

## 效果

| 效果 | 写法 | 说明 |
|---|---|---|
| 纯色 | `SolidFill("#FF0000", alpha=0.8)` | |
| 线性渐变 | `LinearGradient(stops, angle=90, frame=(x, y, w, h))` | 0°=左→右，90°=上→下；`frame` 让多行文字共用一条渐变 |
| 图片填充 | `ImageFill(image, frame=(x, y, w, h))` | 图片拉伸铺满 `frame`，文字超出部分取图片边缘延伸 |
| 外阴影 | `Shadow(color, alpha, blur, offset)` | `blur` 是高斯模糊 sigma，0 = 硬边 |
| 外发光 | `Glow(color, alpha, radius)` | 滤色叠加，加亮背景而不压暗 |
| 描边 | `Stroke(color, width)` | `width` 是单侧向外扩展的像素数 |
| 下划线 / 删除线 | `draw_text(..., underline="double", strike=True)` | 跟着填充、阴影、描边一起着色 |
| 合成斜体 | `Style(synthetic_italic=True)` | 字体没有斜体文件时把字形倾斜 12° |

叠放顺序固定为 **阴影 → 发光 → 描边 → 填充**。所有样式对象都是不可变的，可以放心复用。

## 接口

一步到位：

```python
fx.draw_text(img, xy, text, font, style, anchor="la", spacing=0, line_spacing=1.0, align="left")
fx.text_image(text, font, style, ...)      # 返回独立的 RGBA 图片（透明底）
fx.measure(text, font)                     # (宽, 高)，不含特效余量
```

- `anchor` 和 Pillow 的 `ImageDraw.text` 一致：`l/m/r` + `a/m/s/b/d`，默认左上角。
- 多行用 `\n` 分隔；`padding="auto"`（默认）会自动留够余量，阴影 / 发光 / 描边不会被裁掉。
- `draw_text` **任何模式都就地修改**，返回入参本身，越界部分自动裁掉。非 RGBA 的图片
  内部先转 RGBA 再转回原模式，所以透明度表现为与背景直接混合。

自己排版时用底层接口：

```python
layer = Image.new("RGBA", size, (0, 0, 0, 0))
for glyphs, font, style in my_layout_engine(...):
    fx.draw_glyphs(layer, glyphs, font, style)     # glyphs: (字符, x, 基线 y)；就地修改
canvas.alpha_composite(layer)
```

`draw_glyphs` **永远就地修改**并返回入参本身——所有接口里唯一会动用滤色运算的发光
也在内部处理掉了，不需要接返回值。

单独生成渐变图：

```python
grad = fx.linear_gradient((800, 200), [(0, "#FFD54A"), (1, "#C8860A")], angle=90)
```

## 约定

- 颜色接受 `'#RRGGBB'`、`'#RRGGBBAA'`、`(r, g, b)`、Pillow 颜色名；透明度用 `alpha`（0~1）。
- **所有长度都是像素**。按什么分辨率出图，就按那个分辨率给字号、模糊、偏移。
- 逐字符按 `font.getlength` 推进，不做字距调整（kerning）；只支持横向文字。

## 测试

```bash
python libs/fontfx/tests/run_tests.py
```

每个测试文件也都是独立的脚本，可以单独 `python libs/fontfx/tests/test_gradient.py` 运行。
