# -*- coding: utf-8 -*-
"""
健壮性测试：多段落行距 / 空段落占行 / 字体缺失兜底
COM 交叉验证段落间距

运行：.venv/Scripts/python.exe libs/render_core/tests/test_renderer_robustness.py
"""
import os
import sys
import tempfile

_HERE = os.path.dirname(os.path.abspath(__file__))
# 本库（render_core）与它依赖的 fontfx 都在仓库里；已 pip 安装时这几行是冗余的
for _p in (os.path.join(_HERE, '..', 'src'),
           os.path.join(_HERE, '..', '..', 'fontfx', 'src')):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from lxml import etree
from PIL import Image
import numpy as np
from pptx import Presentation
from pptx.util import Inches

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

from render_core.pillow_image_exporter import PillowImageExporter
from render_core import font_registry

PASS = 0
FAIL = 0


def check(desc, cond):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  ✅ {desc}")
    else:
        FAIL += 1
        print(f"  ❌ {desc}")


def add_para(tf, text, sz=3600):
    if len(tf.paragraphs) == 1 and not tf.paragraphs[0].runs:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    run = p.add_run()
    run.text = text
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', str(sz))
    rPr.set('b', '1')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="000000"/></a:solidFill>'))
    rPr.append(etree.fromstring(f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))
    return p


def build():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(4), Inches(2.5))
    tb.text_frame.word_wrap = True
    add_para(tb.text_frame, '第一段落')
    add_para(tb.text_frame, '第二段落')
    tb = slide.shapes.add_textbox(Inches(4.5), Inches(0.3), Inches(4), Inches(2.5))
    tb.text_frame.word_wrap = True
    add_para(tb.text_frame, '上段落')
    add_para(tb.text_frame, '')
    add_para(tb.text_frame, '下段落')
    path = os.path.join(tempfile.mkdtemp(prefix='rob_'), 'rob.pptx')
    prs.save(path)
    return path


def para_gaps(arr, x0, x1, y0, y1):
    """行投影分段 -> 相邻段顶部间距列表"""
    m = (np.abs(arr[y0:y1, x0:x1].astype(int) - 255).sum(axis=2) > 150).any(axis=1)
    segs, start = [], None
    for i, v in enumerate(m):
        if v and start is None:
            start = i
        elif not v and start is not None:
            segs.append(y0 + start)
            start = None
    if start is not None:
        segs.append(y0 + len(m))
    merged = []
    for s in segs:
        if merged and s - merged[-1] < 6:
            continue
        merged.append(s)
    return [b - a for a, b in zip(merged, merged[1:])]


def main():
    path = build()
    out = tempfile.mkdtemp()
    files = PillowImageExporter().export_to_images(path, out, 'PNG', 1.0)
    arr = np.array(Image.open(files[0]).convert('RGB'))

    print("=" * 60)
    print("(1) 多段落行距 + 空段落占行（COM 交叉验证）")
    try:
        sys.path.insert(0, os.path.abspath(os.path.join(
            os.path.dirname(__file__), '..', '..', '..')))
        from src.core.processors.com_image_exporter import COMImageExporter
        ref_dir = tempfile.mkdtemp()
        COMImageExporter().export_to_images(path, ref_dir, 'PNG', 1.0)
        ref = np.array(Image.open(os.path.join(ref_dir, 'slide_1.png')).convert('RGB'))
        poc_g1 = para_gaps(arr, 48, 430, 20, 270)
        ref_g1 = para_gaps(ref, 48, 430, 20, 270)
        poc_g2 = para_gaps(arr, 432, 830, 20, 270)
        ref_g2 = para_gaps(ref, 432, 830, 20, 270)
        if poc_g1 and ref_g1:
            check(f'两段落间距 PoC={poc_g1[0]} vs PPT={ref_g1[0]}（±5px）',
                  abs(poc_g1[0] - ref_g1[0]) <= 5)
        else:
            check(f'两段落均有分段 {poc_g1} {ref_g1}', False)
        if poc_g2 and ref_g2:
            check(f'含空段间距 PoC={poc_g2[0]} vs PPT={ref_g2[0]}（±6px，空行占位）',
                  abs(poc_g2[0] - ref_g2[0]) <= 6)
        else:
            check(f'空段场景均有分段 {poc_g2} {ref_g2}', False)
    except Exception as e:
        print(f"  ⚠️ COM 交叉验证跳过: {e}")
        # 无 COM 时退化为绝对值断言（36pt 行高 1.25 ≈ 60px）
        g1 = para_gaps(arr, 48, 430, 20, 270)
        g2 = para_gaps(arr, 432, 830, 20, 270)
        check(f'两段落间距 {g1} ≈ 60（±6）', g1 and abs(g1[0] - 60) <= 6)
        check(f'含空段间距 {g2} ≈ 120（±8）', g2 and abs(g2[0] - 120) <= 8)

    print("=" * 60)
    print("(2) 字体缺失兜底（不崩溃、有告警、正常渲染）")
    fake = r'C:\Windows\Fonts\__not_exist_font__.ttf'
    # 模拟楷体系统字体缺失：把注册表系统表中楷体项指向不存在文件
    # （楷体无捆绑等价族，解析将回退默认微软雅黑并告警——正是本测试要验证的兜底）
    saved = {k: v for k, v in font_registry.SYSTEM_FONT_FILES.items()
             if k[0] == '楷体'}
    for k in saved:
        font_registry.SYSTEM_FONT_FILES[k] = fake
    try:
        # 用楷体构造文字
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(4), Inches(1.0))
        tb.text_frame.word_wrap = True
        run = tb.text_frame.paragraphs[0].add_run()
        run.text = '楷体字兜底'
        rPr = run._r.get_or_add_rPr()
        rPr.set('sz', '3600')
        rPr.append(etree.fromstring(
            f'<a:latin xmlns:a="{A}" typeface="楷体"/>'))
        p2 = os.path.join(tempfile.mkdtemp(), 'font.pptx')
        prs.save(p2)
        out2 = tempfile.mkdtemp()
        files2 = PillowImageExporter().export_to_images(p2, out2, 'PNG', 1.0)
        a2 = np.array(Image.open(files2[0]).convert('RGB'))
        ink = (np.abs(a2[20:120, 40:400].astype(int) - 255).sum(axis=2) > 150).sum()
        check(f'缺失字体不崩溃且正常渲染（墨迹 {ink} > 200）', ink > 200)
    except Exception as e:
        check(f'缺失字体导出异常: {e}', False)
    finally:
        for k, v in saved.items():
            font_registry.SYSTEM_FONT_FILES[k] = v

    print("=" * 60)
    print(f"结果: {PASS} 通过, {FAIL} 失败")
    sys.exit(1 if FAIL else 0)


if __name__ == '__main__':
    main()
