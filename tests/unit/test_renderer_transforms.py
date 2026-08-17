# -*- coding: utf-8 -*-
"""
变换渲染测试：形状旋转 / 翻转 / 组合（含嵌套与缩放）
像素级断言 + PowerPoint COM 交叉验证旋转中心

运行：.venv/Scripts/python.exe tests/unit/test_renderer_transforms.py
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from lxml import etree
from PIL import Image
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Emu

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

try:
    from src.core.processors.pillow_image_exporter import PillowImageExporter
    from src.core.processors.render_precheck import RenderPrecheck
except ImportError:
    from core.processors.pillow_image_exporter import PillowImageExporter
    from core.processors.render_precheck import RenderPrecheck

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


def solid(shape, color):
    shape._element.spPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="{color}"/></a:solidFill>'))


def build_test_pptx():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 1) 长条矩形旋转 45°（3:1 长条 -> 旋转后接近正方形包围盒）
    sp = slide.shapes.add_shape(1, Inches(0.5), Inches(1.0), Inches(3), Inches(1))
    solid(sp, '0000FF')
    sp._element.spPr.xfrm.set('rot', '2700000')

    # 2) 水平长条旋转 90°（变为竖条：高>宽）
    sp = slide.shapes.add_shape(1, Inches(4.5), Inches(0.3), Inches(3), Inches(0.8))
    solid(sp, '00AA00')
    sp._element.spPr.xfrm.set('rot', '5400000')  # 90°

    # 3) flipH 箭头（右箭头翻转 -> 指向左）
    sp = slide.shapes.add_shape(33, Inches(0.5), Inches(4.4), Inches(2.2), Inches(0.9))
    solid(sp, 'FF8800')
    sp._element.spPr.xfrm.set('flipH', '1')

    # 4) 组合：两个错位矩形（子坐标系）
    grp = slide.shapes.add_group_shape()
    r1 = grp.shapes.add_shape(1, Inches(0), Inches(0), Inches(1), Inches(0.5))
    solid(r1, 'CC0000')
    r2 = grp.shapes.add_shape(1, Inches(1.2), Inches(0), Inches(1), Inches(0.5))
    solid(r2, '0000CC')
    # 移动组合到 (0.5, 3.2)
    grp.left, grp.top = Inches(0.5), Inches(3.2)

    # 5) 组合整体缩放 2x（ext = 2*chExt）+ 文字
    grp2 = slide.shapes.add_group_shape()
    box = grp2.shapes.add_textbox(Inches(0), Inches(0), Inches(2), Inches(0.5))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = '组字'
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', '2000')
    rPr.set('b', '1')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="000000"/></a:solidFill>'))
    rPr.append(etree.fromstring(f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))
    grp2.left, grp2.top = Inches(4.2), Inches(3.0)
    # ext 翻倍 -> 视觉放大 2x
    from pptx.oxml.ns import qn as _qn
    gx = grp2._element.find(_qn('p:grpSpPr') + '/' + _qn('a:xfrm'))
    ext = gx.find(_qn('a:ext'))
    chExt = gx.find(_qn('a:chExt'))
    ext.set('cx', str(int(chExt.get('cx')) * 2))
    ext.set('cy', str(int(chExt.get('cy')) * 2))

    path = os.path.join(tempfile.mkdtemp(prefix='xform_'), 'xform.pptx')
    prs.save(path)
    return path


def ink_bbox(arr, region):
    x0, y0, x1, y1 = region
    sub = arr[y0:y1, x0:x1]
    m = np.abs(sub.astype(int) - 255).sum(axis=2) > 150
    ys, xs = np.where(m)
    return (y0 + ys.min(), y0 + ys.max(), x0 + xs.min(), x0 + xs.max()) if len(ys) else None


def main():
    path = build_test_pptx()
    out_dir = tempfile.mkdtemp()
    files = PillowImageExporter().export_to_images(path, out_dir, 'PNG', 1.0)
    arr = np.array(Image.open(files[0]).convert('RGB'))

    print("=" * 60)
    print("(1) 长条旋转 45°（3:1 -> 包围盒接近方形）")
    # 中心 (2.0, 1.5)in = (192, 144)px；旋转 45° 后包围盒 271x271
    r = ink_bbox(arr, (40, 0, 350, 300))
    if r:
        w, h = r[3] - r[2], r[1] - r[0]
        check(f'旋转后宽{w} 高{h}，宽高比 {w/max(1,h):.2f} ≈ 1 (±0.15)',
              abs(w / max(1, h) - 1.0) <= 0.15)
    else:
        check('旋转矩形有墨迹', False)

    print("=" * 60)
    print("(2) 水平长条旋转 90°（变为竖条）")
    # 中心 (6.0, 0.7)in = (576, 67)px；旋转后 w=77(高变宽), h=288
    r = ink_bbox(arr, (500, 10, 660, 300))
    if r:
        w, h = r[3] - r[2], r[1] - r[0]
        check(f'旋转后宽{w} 高{h}，高/宽 {h/max(1,w):.1f} >= 2.5', h / max(1, w) >= 2.5)
    else:
        check('旋转竖条有墨迹', False)

    print("=" * 60)
    print("(3) flipH 箭头指向左")
    ax0, ay0 = int(0.5 * 96), int(4.4 * 96)
    ax1, ay1 = int(2.7 * 96), int(5.3 * 96)
    acy = (ay0 + ay1) // 2
    left_tip = arr[acy, ax0 + 6]
    right_notch = arr[ay0 + 4, ax1 - 45]
    check(f'左端为箭头尖(橙) {tuple(left_tip)}',
          left_tip[0] > 200 and 90 < left_tip[1] < 200 and left_tip[2] < 60)
    check(f'右上为翻转后凹口(白) {tuple(right_notch)}', right_notch.sum() > 720)

    print("=" * 60)
    print("(4) 组合：两个错位矩形按子坐标映射")
    # 组合 off=(0.5,3.2)in chOff=(0,0)：r1 在 (0.5,3.2)，r2 在 (1.7,3.2)
    r1c = arr[int(3.2 * 96) + 24, int(0.5 * 96) + 48]   # r1 中心 (1.0,3.45)
    r2c = arr[int(3.2 * 96) + 24, int(1.7 * 96) + 48]   # r2 中心 (2.2,3.45)
    gap = arr[int(3.2 * 96) + 24, int(1.55 * 96)]       # 两块之间空隙
    check(f'子形状1红色 {tuple(r1c)}', r1c[0] > 160 and r1c[2] < 90)
    check(f'子形状2蓝色 {tuple(r2c)}', r2c[2] > 160 and r2c[0] < 90)
    check(f'两块之间空白 {tuple(gap)}', gap.sum() > 720)

    print("=" * 60)
    print("(5) 组合 2x 缩放中的文字（20pt -> 视觉 40pt）")
    # 文字在 (4.2..8.2, 3.0..3.5) 区域放大后
    r = ink_bbox(arr, (int(4.0 * 96), int(2.9 * 96), int(9.4 * 96), int(4.4 * 96)))
    if r:
        glyph_h = r[1] - r[0]
        # 20pt 字在 96dpi 高约 27px（含上下沿），2x 后约 54px
        check(f'缩放后字高 {glyph_h}px 在 42~66 之间', 42 <= glyph_h <= 66)
    else:
        check('组合缩放文字有墨迹', False)

    print("=" * 60)
    print("(6) 预检：旋转/组合不再告警")
    issues = RenderPrecheck(installed_fonts=None).check_pptx(path)
    codes = sorted(set(i.code for i in issues))
    check('无 ROTATION/GROUP 告警', not any(c in ('ROTATION', 'GROUP') for c in codes))
    print('  全部问题码:', codes)

    print("=" * 60)
    print("(7) COM 交叉验证：旋转 45° 矩形中心与包围盒")
    try:
        sys.path.insert(0, os.path.abspath(os.path.join(
            os.path.dirname(__file__), '..', '..', '..', 'poster-again')))
        from src.core.processors.com_image_exporter import COMImageExporter
        ref_dir = tempfile.mkdtemp()
        COMImageExporter().export_to_images(path, ref_dir, 'PNG', 1.0)
        ref = np.array(Image.open(os.path.join(ref_dir, 'slide_1.png')).convert('RGB'))
        rr = ink_bbox(ref, (40, 0, 350, 300))
        pr = ink_bbox(arr, (40, 0, 350, 300))
        if rr and pr:
            dw = abs((rr[3] - rr[2]) - (pr[3] - pr[2]))
            dh = abs((rr[1] - rr[0]) - (pr[1] - pr[0]))
            dcx = abs((rr[2] + rr[3]) / 2 - (pr[2] + pr[3]) / 2)
            dcy = abs((rr[1] + rr[0]) / 2 - (pr[1] + pr[0]) / 2)
            check(f'宽差 {dw}px <= 12', dw <= 12)
            check(f'高差 {dh}px <= 8', dh <= 8)
            check(f'中心差 ({dcx:.0f},{dcy:.0f})px <= 8', dcx <= 8 and dcy <= 8)
        else:
            check('COM 参考有墨迹', False)
    except Exception as e:
        print(f"  ⚠️ COM 交叉验证跳过: {e}")

    print("=" * 60)
    print(f"结果: {PASS} 通过, {FAIL} 失败")
    sys.exit(1 if FAIL else 0)


if __name__ == '__main__':
    main()
