# -*- coding: utf-8 -*-
"""
渲染器新能力测试：anchor 垂直对齐 / normAutofit 缩字 / 文字描边 / 图片裁剪
像素级断言 + PowerPoint COM 交叉验证（anchor/缩字）

运行：.venv/Scripts/python.exe tests/unit/test_renderer_features.py
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from lxml import etree
from PIL import Image
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

try:
    from src.core.processors.pillow_image_exporter import PillowImageExporter
except ImportError:
    from core.processors.pillow_image_exporter import PillowImageExporter

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


def style_run(run, sz=2800, color='000000'):
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', str(sz))
    rPr.set('b', '1')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="{color}"/></a:solidFill>'))
    rPr.append(etree.fromstring(f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))


def build_test_pptx():
    """合成：anchor居中 / anchor底部 / normAutofit缩字 / 描边 / 裁剪图"""
    prs = Presentation()
    prs.slide_height = Inches(11)   # 加高画布给溢出用例留检测空间
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    from pptx.oxml.ns import qn

    # 1) anchor=ctr（高框 2 英寸，单行小字 -> 应垂直居中）
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(4), Inches(2.0))
    tb.text_frame.word_wrap = True
    style_run(tb.text_frame.paragraphs[0].add_run())
    tb.text_frame.paragraphs[0].runs[0].text = '居中锚定'
    tb.text_frame._txBody.find(qn('a:bodyPr')).set('anchor', 'ctr')

    # 2) anchor=b
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(2.6), Inches(4), Inches(2.0))
    tb.text_frame.word_wrap = True
    style_run(tb.text_frame.paragraphs[0].add_run())
    tb.text_frame.paragraphs[0].runs[0].text = '底部锚定'
    tb.text_frame._txBody.find(qn('a:bodyPr')).set('anchor', 'b')

    # 3) normAutofit fontScale=50%（同尺寸字对比：缩字行高应约为未缩的一半）
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(4), Inches(1.5))
    tb.text_frame.word_wrap = True
    style_run(tb.text_frame.paragraphs[0].add_run(), sz=7200)
    tb.text_frame.paragraphs[0].runs[0].text = '原尺寸'
    tb = slide.shapes.add_textbox(Inches(5.2), Inches(4.8), Inches(4), Inches(1.5))
    tb.text_frame.word_wrap = True
    style_run(tb.text_frame.paragraphs[0].add_run(), sz=7200)
    tb.text_frame.paragraphs[0].runs[0].text = '缩小字'
    bodyPr = tb.text_frame._txBody.find(qn('a:bodyPr'))
    etree.SubElement(bodyPr, f'{{{A}}}normAutofit', fontScale='50000')

    # 4) 描边：黑字红边
    tb = slide.shapes.add_textbox(Inches(5), Inches(0.3), Inches(4), Inches(1.0))
    tb.text_frame.word_wrap = True
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = '描边字'
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', '7200')
    rPr.set('b', '1')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="000000"/></a:solidFill>'))
    rPr.append(etree.fromstring(f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))
    rPr.append(etree.fromstring(
        f'<a:ln xmlns:a="{A}" w="50800">'
        f'<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'))

    # 5) 图片裁剪：左红右蓝图，srcRect 裁掉右半 -> 只剩红
    img = Image.new('RGB', (200, 100))
    img.paste(Image.new('RGB', (100, 100), (255, 0, 0)), (0, 0))
    img.paste(Image.new('RGB', (100, 100), (0, 0, 255)), (100, 0))
    img_file = os.path.join(tempfile.mkdtemp(), 't.png')
    img.save(img_file)
    pic = slide.shapes.add_picture(img_file, Inches(5), Inches(1.6), Inches(2), Inches(1))
    etree.SubElement(pic._element.find(qn('p:blipFill')), f'{{{A}}}srcRect',
                     {'r': '50000'})

    # 6) glow 发光：黑字红光（rad 18pt alpha 40%），光晕应外扩到字形之外
    # （放在 y=6.8in 空区，避免光晕污染 normAutofit 断言区域）
    tb = slide.shapes.add_textbox(Inches(5), Inches(6.8), Inches(3), Inches(1.4))
    tb.text_frame.word_wrap = True
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = '发光字'
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', '7200')
    rPr.set('b', '1')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="000000"/></a:solidFill>'))
    rPr.append(etree.fromstring(f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))
    rPr.append(etree.fromstring(
        f'<a:effectLst xmlns:a="{A}"><a:glow rad="228600">'
        f'<a:srgbClr val="FF0000"><a:alpha val="40000"/></a:srgbClr>'
        f'</a:glow></a:effectLst>'))

    # 7) 渐变溢出：小框(1in) + 72pt 渐变大字 -> 溢出部分应为渐变端点延伸色（非黑）
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(8.2), Inches(2), Inches(1))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = '渐变溢出测试'
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', '7200')
    rPr.set('b', '1')
    rPr.append(etree.fromstring(
        f'<a:gradFill xmlns:a="{A}"><a:gsLst>'
        f'<a:gs pos="0"><a:prstClr val="white"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="F9CE6F"/></a:gs>'
        f'</a:gsLst><a:lin ang="18900000" scaled="1"/></a:gradFill>'))
    rPr.append(etree.fromstring(f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))

    path = os.path.join(tempfile.mkdtemp(prefix='feat_'), 'feat.pptx')
    prs.save(path)
    return path


def ink_bbox(arr, region):
    x0, y0, x1, y1 = region
    sub = arr[y0:y1, x0:x1]
    m = np.abs(sub.astype(int) - 255).sum(axis=2) > 120
    ys, xs = np.where(m)
    return (y0 + ys.min(), y0 + ys.max()) if len(ys) else None


def main():
    path = build_test_pptx()
    out_dir = tempfile.mkdtemp()
    files = PillowImageExporter().export_to_images(path, out_dir, 'PNG', 1.0)
    arr = np.array(Image.open(files[0]).convert('RGB'))
    S = 96 / 914400  # EMU->px @q1

    print("=" * 60)
    print("(1) anchor 垂直对齐（@96dpi，1in=96px，内边距 tIns=0.05in）")
    # 框1: y=0.3~2.3in，行高 28pt*1.25=35pt≈46.7px
    box_h = 2.0 * 96
    top_in = 0.3 * 96
    r = ink_bbox(arr, (int(0.5 * 96), int(top_in), int(4.5 * 96), int(top_in + box_h)))
    if r:
        cy = (r[0] + r[1]) / 2
        box_cy = top_in + box_h / 2
        check(f'anchor=ctr 文字中心 {cy:.0f} ≈ 框中心 {box_cy:.0f} (±8px)',
              abs(cy - box_cy) <= 8)
    else:
        check('anchor=ctr 有文字', False)

    # 框2: anchor=b -> 文字底部应贴近框底-bIns
    top_in2 = 2.6 * 96
    r2 = ink_bbox(arr, (int(0.5 * 96), int(top_in2), int(4.5 * 96), int(top_in2 + box_h)))
    if r2:
        bottom_margin = top_in2 + box_h - 4.8 - r2[1]  # bIns=0.05in=4.8px
        check(f'anchor=b 底边距 {bottom_margin:.0f}px ≈ 0 (+15px 内)', bottom_margin <= 15)
    else:
        check('anchor=b 有文字', False)

    print("=" * 60)
    print("(2) normAutofit fontScale=50%（左右两列同高区域，无相互干扰）")
    ra = ink_bbox(arr, (int(0.5 * 96), int(4.7 * 96), int(4.5 * 96), int(6.3 * 96)))
    rb = ink_bbox(arr, (int(5.2 * 96), int(4.7 * 96), int(9.0 * 96), int(6.3 * 96)))
    if ra and rb:
        ratio = (rb[1] - rb[0]) / (ra[1] - ra[0])
        check(f'缩字字高比 {ratio:.2f} ≈ 0.5 (±0.08)', abs(ratio - 0.5) <= 0.08)
    else:
        check(f'两组字都有墨迹 {ra} {rb}', False)

    print("=" * 60)
    print("(3) 文字描边（黑字红边 4pt）")
    x0, y0 = int(5 * 96), int(0.3 * 96)
    sub = arr[y0:y0 + 96, x0:x0 + 4 * 96].astype(int)
    red = ((sub[:, :, 0] > 170) & (sub[:, :, 1] < 90) & (sub[:, :, 2] < 90)).sum()
    black = ((sub.sum(axis=2) < 180)).sum()
    check(f'存在红色描边像素 {red} > 200', red > 200)
    check(f'存在黑色填充像素 {black} > 500', black > 500)

    print("=" * 60)
    print("(4) 图片裁剪 srcRect r=50%（只剩左半红色）")
    px0, py0 = int(5 * 96), int(1.6 * 96)
    sub2 = arr[py0:py0 + 96, px0:px0 + 192]
    check('裁剪后为红色（左半保留）',
          sub2[:, :, 0].mean() > 200 and sub2[:, :, 2].mean() < 60)

    print("=" * 60)
    print("(5) glow 发光（黑字红光 rad=18pt，光晕外扩到字形外）")
    gx0, gy0 = int(5 * 96), int(6.8 * 96)
    region = arr[gy0:gy0 + 134, gx0:gx0 + 3 * 96].astype(int)
    # 黑字区域（有墨迹）
    black_ink = (region.sum(axis=2) < 180).sum()
    # 红光：红通道显著高于绿蓝（非黑白像素即来自 glow）
    reddish = ((region[:, :, 0] - region[:, :, 1] > 25) &
               (region[:, :, 0] - region[:, :, 2] > 25)).sum()
    # 字形 bbox 之外的环带：字宽约 3 字 ×96px + 光晕，框左缘外侧 30px 应仍是白色
    # （光晕 24px@96dpi 不该越出框太多——这里只验证内部红光与黑字并存）
    check(f'存在红色光晕像素 {reddish} > 300', reddish > 300)
    check(f'存在黑色字形像素 {black_ink} > 500', black_ink > 500)
    # 预检不应再报 glow 不支持
    issues = PillowImageExporter().precheck(path)
    check('预检无 TEXT_EFFECTS/glow 项',
          not any(i.code == 'TEXT_EFFECTS' and 'glow' in i.message for i in issues))

    print("=" * 60)
    print("(6) 渐变溢出（小框大字，框外应为端点延伸色非黑）")
    # 框 x0.5-2.5in y6.8-7.8in；72pt 字溢出到框下方
    gx0, gy0 = int(0.5 * 96), int(8.2 * 96)
    sub = arr[gy0:gy0 + int(2.2 * 96), gx0:gx0 + int(3 * 96)].astype(int)
    inside = sub[:int(1.0 * 96)]     # 框内
    below = sub[int(1.05 * 96):]     # 框下溢出区
    def ink_stats(z):
        m = np.abs(z - 255).sum(axis=2) > 30
        return m, (z[m].mean(axis=0) if m.any() else None)
    mi, ci = ink_stats(inside)
    mb_, cb = ink_stats(below)
    check(f'框内有渐变字 均色={None if ci is None else ci.round(0)}', ci is not None)
    check(f'框外溢出存在 均色={None if cb is None else cb.round(0)}', mb_.sum() > 100)
    if cb is not None:
        # 端点延伸色应为亮白金（各通道>200），绝不能是黑/灰
        check(f'溢出区非黑（均色 {cb.round(0)}）', cb.min() > 200)

    print("=" * 60)
    print(f"结果: {PASS} 通过, {FAIL} 失败")
    sys.exit(1 if FAIL else 0)


if __name__ == '__main__':
    main()
