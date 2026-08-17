# -*- coding: utf-8 -*-
"""
形状渲染测试：预设几何 / 纯色/渐变/图片填充 / 形状阴影 / 描边
像素级断言 + 预检白名单验证

运行：.venv/Scripts/python.exe tests/unit/test_renderer_shapes.py
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from lxml import etree
from PIL import Image
import numpy as np
from pptx import Presentation
from pptx.util import Inches

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


def sp_fill(shape, xml):
    shape._element.spPr.append(etree.fromstring(xml))


def build_test_pptx():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 1) roundRect 纯蓝填充 + 红色阴影
    sp = slide.shapes.add_shape(5, Inches(0.5), Inches(0.3), Inches(2), Inches(1.2))  # 5=ROUNDED_RECTANGLE
    spFill = f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="0000FF"/></a:solidFill>'
    sp_fill(sp, spFill)
    spPr = sp._element.spPr
    eff = etree.SubElement(spPr, f'{{{A}}}effectLst')
    etree.SubElement(eff, f'{{{A}}}outerShdw',
                     {'blurRad': '50800', 'dist': '38100', 'dir': '2700000'})
    etree.SubElement(eff[0], f'{{{A}}}srgbClr', {'val': 'FF0000'})

    # 2) ellipse 垂直渐变 红到蓝
    sp = slide.shapes.add_shape(9, Inches(3.0), Inches(0.3), Inches(2), Inches(1.2))  # 9=OVAL
    sp_fill(sp, f'<a:gradFill xmlns:a="{A}"><a:gsLst>'
                f'<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>'
                f'<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>'
                f'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')

    # 3) rightArrow 黄色箭头
    sp = slide.shapes.add_shape(33, Inches(0.5), Inches(1.8), Inches(2.5), Inches(1.0))  # 33=RIGHT_ARROW
    sp_fill(sp, f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="FFFF00"/></a:solidFill>')

    # 4) roundRect 无填充 + 绿描边 3pt
    sp = slide.shapes.add_shape(5, Inches(3.2), Inches(1.8), Inches(2), Inches(1.0))
    sp_fill(sp, f'<a:noFill xmlns:a="{A}"/>')
    ln = etree.fromstring(
        f'<a:ln xmlns:a="{A}" w="38100"><a:solidFill>'
        f'<a:srgbClr val="00AA00"/></a:solidFill></a:ln>')
    sp._element.spPr.append(ln)

    # 5) rect 图片填充（左红右蓝图）
    img = Image.new('RGB', (200, 100))
    img.paste(Image.new('RGB', (100, 100), (255, 0, 0)), (0, 0))
    img.paste(Image.new('RGB', (100, 100), (0, 0, 255)), (100, 0))
    img_file = os.path.join(tempfile.mkdtemp(), 's.png')
    img.save(img_file)
    sp = slide.shapes.add_shape(1, Inches(0.5), Inches(3.1), Inches(2), Inches(1.0))  # 1=RECTANGLE
    # 替换默认填充为图片填充
    for old in list(sp._element.spPr):
        if old.tag.endswith('}solidFill') or old.tag.endswith('}noFill'):
            sp._element.spPr.remove(old)
    pic = slide.shapes.add_picture(img_file, Inches(5.5), Inches(3.1),
                                   Inches(2), Inches(1))
    # 手工把图片 rId 填到矩形上
    blip_rId = pic._element.blip_rId
    blipFill = etree.fromstring(
        f'<a:blipFill xmlns:a="{A}" xmlns:r="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships">'
        f'<a:blip r:embed="{blip_rId}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>')
    # blipFill 需要插在 prstGeom 之前
    prstGeom = sp._element.spPr.find(f'{{{A}}}prstGeom')
    sp._element.spPr.insert(list(sp._element.spPr).index(prstGeom), blipFill)
    slide.shapes._spTree.remove(pic._element)  # 图片本体不再单独显示

    # 6) 不支持的形状（cloud）-> 预检应报 SHAPE_GEOM
    try:
        sp = slide.shapes.add_shape(179, Inches(5.5), Inches(0.3), Inches(2), Inches(1.2))  # 179=CLOUD
        sp_fill(sp, f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="888888"/></a:solidFill>')
    except Exception:
        pass

    path = os.path.join(tempfile.mkdtemp(prefix='shapes_'), 'shapes.pptx')
    prs.save(path)
    return path


def main():
    path = build_test_pptx()
    out_dir = tempfile.mkdtemp()
    files = PillowImageExporter().export_to_images(path, out_dir, 'PNG', 1.0)
    arr = np.array(Image.open(files[0]).convert('RGB'))

    print("=" * 60)
    print("(1) roundRect 纯蓝填充 + 红色阴影")
    x0, y0 = int(0.5 * 96), int(0.3 * 96)          # 48, 28
    x1, y1 = int(2.5 * 96), int(1.5 * 96)          # 240, 144
    cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
    c = arr[cy, cx]
    check(f'中心为蓝色 {tuple(c)}', c[2] > 180 and c[0] < 80)
    corner = arr[y0 + 2, x0 + 2]                   # 圆角外应为白
    check(f'左上圆角外为白色 {tuple(corner)}', corner.sum() > 720)
    # 阴影：右下偏移处（形状外）应为红调
    sh = arr[y1 + 4, x1 - 8]
    check(f'右下阴影为红调 {tuple(sh)}', sh[0] > 120 and sh[0] > sh[2] + 30)

    print("=" * 60)
    print("(2) ellipse 垂直渐变红->蓝")
    ex0, ey0, ex1, ey1 = int(3.0 * 96), int(0.3 * 96), int(5.0 * 96), int(1.5 * 96)
    ecx, ecy = (ex0 + ex1) // 2, (ey0 + ey1) // 2
    top_c = arr[ey0 + 6, ecx]
    bot_c = arr[ey1 - 6, ecx]
    check(f'椭圆上部偏红 {tuple(top_c)}', top_c[0] > top_c[2] + 40)
    check(f'椭圆下部偏蓝 {tuple(bot_c)}', bot_c[2] > bot_c[0] + 40)
    corner = arr[ey0 + 2, ex0 + 2]
    check(f'椭圆外角落为白 {tuple(corner)}', corner.sum() > 720)

    print("=" * 60)
    print("(3) rightArrow 黄色")
    ax0, ay0 = int(0.5 * 96), int(1.8 * 96)        # 48, 172
    ax1, ay1 = int(3.0 * 96), int(2.8 * 96)        # 288, 268
    acx, acy = (ax0 + ax1) // 2, (ay0 + ay1) // 2
    tip = arr[acy, ax1 - 6]                        # 箭头尖端（右侧中部）
    notch = arr[ay0 + 3, ax1 - 40]                 # 箭头凹口（右上缺口）
    check(f'尖端为黄色 {tuple(tip)}', tip[0] > 200 and tip[1] > 200)
    check(f'上凹口为白色 {tuple(notch)}', notch.sum() > 720)

    print("=" * 60)
    print("(4) roundRect 无填充 + 绿描边")
    rx0, ry0, rx1, ry1 = int(3.2 * 96), int(1.8 * 96), int(5.2 * 96), int(2.8 * 96)
    mid = arr[(ry0 + ry1) // 2, (rx0 + rx1) // 2]
    check(f'内部无填充(白) {tuple(mid)}', mid.sum() > 720)
    border = arr[(ry0 + ry1) // 2, rx0 + 3]
    check(f'左边缘为绿色描边 {tuple(border)}', border[1] > 100 and border[1] > border[0] + 30)

    print("=" * 60)
    print("(5) rect 图片填充（左红右蓝）")
    bx0, by0, bx1 = int(0.5 * 96), int(3.1 * 96), int(2.5 * 96)
    bcy = by0 + 48
    left_c = arr[bcy, bx0 + 10]
    right_c = arr[bcy, bx1 - 10]
    check(f'左半为红 {tuple(left_c)}', left_c[0] > 200 and left_c[2] < 80)
    check(f'右半为蓝 {tuple(right_c)}', right_c[2] > 200 and right_c[0] < 80)

    print("=" * 60)
    print("(6) 预检白名单验证")
    issues = RenderPrecheck(installed_fonts=None).check_pptx(path)
    codes = [i.code for i in issues]
    geom_issues = [i for i in issues if i.code == 'SHAPE_GEOM']
    check('cloud 形状报 SHAPE_GEOM', len(geom_issues) >= 1)
    check('roundRect/ellipse/arrow/描边/阴影/图片填充均不告警',
          not any(i.code in ('SHAPE_FILL', 'SHAPE_EFFECTS', 'SHAPE_LINE')
                  for i in issues))
    print('  全部问题码:', sorted(set(codes)))

    print("=" * 60)
    print(f"结果: {PASS} 通过, {FAIL} 失败")
    sys.exit(1 if FAIL else 0)


if __name__ == '__main__':
    main()
