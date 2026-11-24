# -*- coding: utf-8 -*-
"""生成一个带文字特效的示例 PPTX 并渲染成图片。

用法：python libs/render_core/examples/render_example.py [输出目录]
"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

from lxml import etree  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

import render_core  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PER_PT = 12700


def _xml(fragment):
    """一段 DrawingML（a: 命名空间）-> lxml 元素"""
    wrapper = etree.fromstring(f'<w xmlns:a="{A}">{fragment}</w>')
    return wrapper[0]


def _style_run(run, *, size_pt, gradient=None, glow=None, shadow=None, stroke=None):
    """用原生 OOXML 给 run 加特效——这正是 PowerPoint 自己保存文字特效的方式。

    rPr 子元素的 schema 顺序是：描边 ln -> 填充 -> 效果 effectLst。
    """
    run.font.size = Pt(size_pt)
    run.font.bold = True
    rPr = run._r.get_or_add_rPr()
    if stroke:
        width, color = stroke
        rPr.append(_xml(f'<a:ln w="{int(width * EMU_PER_PT)}"><a:solidFill>'
                        f'<a:srgbClr val="{color}"/></a:solidFill></a:ln>'))
    if gradient:
        stops = "".join(f'<a:gs pos="{int(p * 100000)}"><a:srgbClr val="{c}"/></a:gs>'
                        for p, c in gradient)
        rPr.append(_xml(f'<a:gradFill><a:gsLst>{stops}</a:gsLst>'
                        '<a:lin ang="5400000" scaled="0"/></a:gradFill>'))   # 90° = 上到下
    else:
        rPr.append(_xml('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'))
    if glow or shadow:
        eff = etree.Element(f"{{{A}}}effectLst")
        if shadow:                      # (模糊 pt, 距离 pt, 颜色, 不透明度)
            blur, dist, color, alpha = shadow
            shd = etree.SubElement(eff, f"{{{A}}}outerShdw",
                                   blurRad=str(int(blur * EMU_PER_PT)),
                                   dist=str(int(dist * EMU_PER_PT)),
                                   dir="2700000", algn="ctr")            # 右下方向
            clr = etree.SubElement(shd, f"{{{A}}}srgbClr", val=color)
            etree.SubElement(clr, f"{{{A}}}alpha", val=str(int(alpha * 100000)))
        if glow:                        # (半径 pt, 颜色, 不透明度)
            radius, color, alpha = glow
            g = etree.SubElement(eff, f"{{{A}}}glow", rad=str(int(radius * EMU_PER_PT)))
            clr = etree.SubElement(g, f"{{{A}}}srgbClr", val=color)
            etree.SubElement(clr, f"{{{A}}}alpha", val=str(int(alpha * 100000)))
        rPr.append(eff)


def build(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 深色背景
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(0x1B, 0x2A, 0x49)

    def text(left, top, width, height, alt_text=None):
        """加一个文本框，返回它的第一个 run。alt_text 即「替代文本」占位符。"""
        box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        if alt_text:
            box._element.nvSpPr.cNvPr.set("descr", alt_text)
        box.text_frame.word_wrap = True      # python-pptx 新建的文本框默认禁止换行(wrap=none)
        box.text_frame.text = ""
        return box.text_frame.paragraphs[0].add_run()

    r = text(0.6, 0.5, 8.8, 1.8)
    r.text = "恭喜发财"
    _style_run(r, size_pt=96, gradient=[(0.0, "FFF7C2"), (0.5, "FFD84A"), (1.0, "C97A08")],
               shadow=(5, 4, "000000", 0.6))

    r = text(0.6, 2.5, 8.8, 1.0)
    r.text = "HAPPY NEW YEAR"
    _style_run(r, size_pt=54, glow=(8, "FF3DCF", 0.9))

    r = text(0.6, 3.7, 4.5, 1.0)
    r.text = "2026"
    _style_run(r, size_pt=72, stroke=(2.5, "FF5A36"))

    # 替代文本 = ph_姓名：Personalizer 会把这个文本框整体换成数据里的「姓名」
    r = text(5.8, 3.85, 3.6, 0.9, alt_text="ph_姓名")
    r.text = "敬上"
    _style_run(r, size_pt=40)

    prs.save(path)
    return path


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.getcwd(), "render_core_example")
    os.makedirs(out_dir, exist_ok=True)
    pptx = build(os.path.join(out_dir, "example.pptx"))

    result = render_core.render(pptx, out_dir, "PNG", 2.0)
    print(f"渲染完成: {result.images[0]}")
    if result.issues:
        for issue in result.issues:
            print(f"  [预检] 第{issue.slide}页 {issue.shape!r} [{issue.code}] {issue.message}")
    else:
        print("  预检通过：模板里的元素全部在支持范围内")

    # 批量个性化：同一个模板只解析一次，逐行替换占位符
    pz = render_core.Personalizer(pptx)
    print("模板占位符:", pz.fields())
    batch_dir = os.path.join(out_dir, "batch")
    for name in ("张三", "李四"):
        images = pz.render_row({"ph_姓名": f"{name} 敬上"}, batch_dir)
        print(f"  {name} -> {images[0]}")


if __name__ == "__main__":
    main()
