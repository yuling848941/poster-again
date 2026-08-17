# -*- coding: utf-8 -*-
"""
精确校准脚本：构造白底合成 PPT（已知文字/字号/渐变/阴影），分别用
PowerPoint(COM) 和 Pillow PoC 渲染，逐行测量文本位置差异，
输出校准数据：首行基线偏移、行距比、字距宽度比、字形高度比。
"""
import os
import subprocess
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from PIL import Image
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
TEST_PPTX = os.path.join(HERE, 'out', 'calib_test.pptx')

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
GRAD_XML = (
    f'<a:gradFill xmlns:a="{A}">'
    '<a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="lt1"/></a:gs>'
    '<a:gs pos="100000"><a:srgbClr val="F9CE6F"/></a:gs>'
    '</a:gsLst>'
    '<a:lin ang="5400000" scaled="0"/>'
    '<a:tileRect/>'
    '</a:gradFill>'
)
SHADOW_XML = (
    f'<a:effectLst xmlns:a="{A}">'
    '<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">'
    '<a:srgbClr val="C00000"><a:alpha val="40000"/></a:srgbClr>'
    '</a:outerShdw>'
    '</a:effectLst>'
)


def style_run(run, sz, spc=300, shadow=True):
    run.font.size = Pt(sz)
    run.font.bold = True
    run.font.name = '微软雅黑'
    rPr = run._r.get_or_add_rPr()
    rPr.set('spc', str(spc))
    rPr.set('dirty', '0')
    rPr.append(etree.fromstring(GRAD_XML))
    if shadow:
        rPr.append(etree.fromstring(SHADOW_XML))


def build_test_pptx():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    def add_box(x, y, w, h):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        return tb

    # 1) 大号金额：单行 72pt + 字距 + 渐变 + 阴影
    tb = add_box(0.4, 0.2, 9.0, 1.5)
    style_run(tb.text_frame.paragraphs[0].add_run(), 72)
    tb.text_frame.paragraphs[0].runs[0].text = '88,888元'

    # 2) 中号单行 44pt
    tb = add_box(0.4, 1.8, 9.0, 1.1)
    style_run(tb.text_frame.paragraphs[0].add_run(), 44)
    tb.text_frame.paragraphs[0].runs[0].text = '递交一生中意（金耀版）1件'

    # 3) 小号多行 28pt 行距120%（长文本强制换行 -> 校准行距）
    tb = add_box(0.4, 3.2, 6.0, 3.0)
    tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = 1.2
    style_run(p.add_run(), 28)
    p.runs[0].text = '第一行校准文本ABC第二行内容defgh第三行文字ijklmn第四行尾行'

    # 4) 居中 36pt（校准居中）
    tb = add_box(0.4, 5.8, 9.0, 1.0)
    p = tb.text_frame.paragraphs[0]
    from pptx.enum.text import PP_ALIGN
    p.alignment = PP_ALIGN.CENTER
    style_run(p.add_run(), 36)
    p.runs[0].text = '居中对齐校准 CENTER TEST'

    prs.save(TEST_PPTX)
    print('合成测试 PPT:', TEST_PPTX)


def render_refs():
    # COM 参考图
    from src.core.processors.com_image_exporter import COMImageExporter
    exp = COMImageExporter()
    exp.export_to_images(TEST_PPTX, os.path.join(HERE, 'out', 'calib_ref'), 'PNG', 2.0)
    # PoC 渲染
    from pillow_render_poc import PillowRenderer
    img = PillowRenderer(TEST_PPTX, 2.0).render()
    out = os.path.join(HERE, 'out', 'calib_poc.png')
    img.convert('RGB').save(out)
    print('PoC 渲染:', out)


def measure(path, name):
    """白底图 -> 按行投影找文本行，返回 [(top, bottom, left, right, width, height)]"""
    import numpy as np
    a = np.array(Image.open(path).convert('L'), dtype=np.uint8)
    mask = a < 235  # 非白像素
    rows = mask.any(axis=1)
    lines, start = [], None
    for i, v in enumerate(rows):
        if v and start is None:
            start = i
        elif not v and start is not None:
            lines.append((start, i))
            start = None
    if start is not None:
        lines.append((start, len(rows)))
    out = []
    for t, b in lines:
        cols = mask[t:b].any(axis=0)
        xs = np.where(cols)[0]
        out.append((t, b, xs[0], xs[-1], xs[-1] - xs[0] + 1, b - t))
    print(f'--- {name}: {len(out)} 行 ---')
    for t, b, l, r, w, h in out:
        print(f'  top={t:4d} bottom={b:4d} left={l:4d} right={r:4d} w={w:4d} h={h:3d}')
    return out


def main():
    build_test_pptx()
    render_refs()
    ref = measure(os.path.join(HERE, 'out', 'calib_ref', 'slide_1.png'), 'PowerPoint参考')
    poc = measure(os.path.join(HERE, 'out', 'calib_poc.png'), 'PoC渲染')
    n = min(len(ref), len(poc))
    print('\n=== 对比（poc - ref, 单位 px @q2） ===')
    for i in range(n):
        rt, rb, rl, rr, rw, rh = ref[i]
        pt_, pb, pl, pr_, pw, ph = poc[i]
        print(f'行{i}: Δtop={pt_-rt:+4d} Δbottom={pb-rb:+4d} Δleft={pl-rl:+4d} '
              f'Δw={pw-rw:+4d} Δh={ph-rh:+3d}  (ref h={rh} w={rw})')
    # 行距比（多行块：行3-6 属于第3个文本框）
    if n >= 6:
        ref_pitch = ref[4][0] - ref[3][0]
        poc_pitch = poc[4][0] - poc[3][0]
        print(f'\n行距: ref={ref_pitch}px poc={poc_pitch}px 比值={poc_pitch/ref_pitch:.4f}')


if __name__ == '__main__':
    main()
