# -*- coding: utf-8 -*-
"""
文字装饰测试：合成斜体 / 删除线 / 下划线（单双线）/ 文字描边验证
像素级断言 + PowerPoint COM 交叉验证删除线位置

运行：.venv/Scripts/python.exe libs/render_core/tests/test_renderer_text_decorations.py
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
from render_core.render_precheck import RenderPrecheck

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


def add_text(slide, x, y, text, sz=4800, **attrs):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(4), Inches(1.2))
    tb.text_frame.word_wrap = True
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = text
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', str(sz))
    rPr.set('b', '1')
    for k, v in attrs.items():
        rPr.set(k, v)
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="000000"/></a:solidFill>'))
    rPr.append(etree.fromstring(f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))
    return tb


def ink_bbox(arr, region):
    x0, y0, x1, y1 = region
    sub = arr[y0:y1, x0:x1]
    m = np.abs(sub.astype(int) - 255).sum(axis=2) > 150
    ys, xs = np.where(m)
    return (y0 + ys.min(), y0 + ys.max(), x0 + xs.min(), x0 + xs.max()) if len(ys) else None


def build_test_pptx():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 1) 斜体 vs 常规（同文本同字号同位置错开）
    add_text(slide, 0.4, 0.2, '国国国', i='1')
    add_text(slide, 0.4, 1.2, '国国国')

    # 2) 删除线（回字有内部空腔，72pt 空腔更大）
    add_text(slide, 3.5, 0.2, '回', sz=7200, strike='sngStrike')
    add_text(slide, 3.5, 1.4, '回', sz=7200)

    # 3) 下划线 单/双（拉丁字母贴基线，下方空）
    add_text(slide, 0.4, 2.4, 'HHHH', u='sng')
    add_text(slide, 0.4, 3.3, 'HHHH', u='dbl')
    add_text(slide, 0.4, 4.2, 'HHHH')

    # 4) 描边（已支持能力回归验证）+ 非纯色描边（预检应告警）
    tb = slide.shapes.add_textbox(Inches(3.5), Inches(2.4), Inches(4), Inches(1.2))
    tb.text_frame.word_wrap = True
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = '描边'
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', '4800')
    rPr.set('b', '1')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="000000"/></a:solidFill>'))
    rPr.append(etree.fromstring(f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))
    rPr.append(etree.fromstring(
        f'<a:ln xmlns:a="{A}" w="50800">'
        f'<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'))
    tb2 = slide.shapes.add_textbox(Inches(3.5), Inches(3.5), Inches(4), Inches(1.2))
    tb2.text_frame.word_wrap = True
    run = tb2.text_frame.paragraphs[0].add_run()
    run.text = '渐变边'
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', '2800')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="000000"/></a:solidFill>'))
    rPr.append(etree.fromstring(
        f'<a:ln xmlns:a="{A}" w="50800"><a:gradFill><a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>'
        f'</a:gsLst><a:lin ang="0"/></a:gradFill></a:ln>'))

    path = os.path.join(tempfile.mkdtemp(prefix='deco_'), 'deco.pptx')
    prs.save(path)
    return path


def main():
    path = build_test_pptx()
    out_dir = tempfile.mkdtemp()
    files = PillowImageExporter().export_to_images(path, out_dir, 'PNG', 1.0)
    arr = np.array(Image.open(files[0]).convert('RGB'))

    print("=" * 60)
    print("(1) 合成斜体（12° 斜切）")
    ra = ink_bbox(arr, (30, 10, 320, 110))    # 斜体行 y=0.2in（x<320 排除回字）
    rb = ink_bbox(arr, (30, 106, 320, 210))   # 常规行 y=1.2in
    if ra and rb:
        wa, ha = ra[3] - ra[2], ra[1] - ra[0]
        wb, hb = rb[3] - rb[2], rb[1] - rb[0]
        check(f'斜体字高 {ha} ≈ 常规 {hb}（±3px）', abs(ha - hb) <= 3)
        check(f'斜体比常规宽 {wa - wb}px（预期 8~20px 斜切扩展）', 8 <= wa - wb <= 20)
        # 顶部窄带右缘 vs 底部窄带右缘：右倾量 ≈ tan12° * 字高
        def band_right(y0, y1):
            m = np.abs(arr[y0:y1, 30:320].astype(int) - 255).sum(axis=2) > 150
            xs = np.where(m.any(axis=0))[0]
            return 30 + xs.max() if len(xs) else 0
        top_x = band_right(ra[0], ra[0] + 6)
        bot_x = band_right(ra[1] - 6, ra[1])
        check(f'顶部右缘 {top_x} - 底部右缘 {bot_x} = {top_x - bot_x}px（右倾 6~14px）',
              6 <= top_x - bot_x <= 14)
    else:
        check(f'两行都有墨迹 {ra} {rb}', False)

    print("=" * 60)
    print("(2) 删除线（回字空腔内出现连续横线段）")
    em = 96  # 72pt

    def max_run(a, y, x0, x1):
        row = (a[y, x0:x1].sum(axis=1) < 500)
        best = cur = 0
        for v in row:
            cur = cur + 1 if v else 0
            best = max(best, cur)
        return best

    yb = 0.2 * 96 + 4.8 + 0.72 * 1.25 * em
    gx = int(3.5 * 96 + 9.6)
    strike_y = int(round(yb - 0.245 * em))
    run1 = max(max_run(arr, y, gx + 15, gx + em - 15)
               for y in range(strike_y - 2, strike_y + 3))
    check(f'删除线连续段 {run1}px >= 60（横贯空腔，区别于内部横画42px）', run1 >= 60)
    yb2 = 1.4 * 96 + 4.8 + 0.72 * 1.25 * em
    strike_y2 = int(round(yb2 - 0.245 * em))
    run2 = max(max_run(arr, y, gx + 15, gx + em - 15)
               for y in range(strike_y2 - 2, strike_y2 + 3))
    check(f'对照组连续段 {run2}px < 50（仅内部横画）', run2 < 50)

    print("=" * 60)
    print("(3) 下划线（单/双）")
    for y_in, label, expect_bars in ((2.4, '单线', 1), (3.3, '双线', 2), (4.2, '无', 0)):
        yb = y_in * 96 + 4.8 + 0.72 * 1.25 * 64
        gx = 0.4 * 96 + 9.6
        # 基线下方 0.09em 起扫描三条候选线
        found = 0
        for dy in range(int(0.06 * em), int(0.22 * em)):
            row = arr[int(yb) + dy, int(gx): int(gx) + 150]
            if (row.sum(axis=1) < 500).sum() > 100:
                found += 1
        # 连续行合并：粗略统计着色行段数
        bars, prev = 0, False
        for dy in range(int(0.06 * em), int(0.25 * em)):
            row = arr[int(yb) + dy, int(gx): int(gx) + 150]
            c = (row.sum(axis=1) < 500).sum() > 100
            if c and not prev:
                bars += 1
            prev = c
        check(f'{label}：基线下方检测到 {bars} 条横线（预期 {expect_bars}）',
              bars == expect_bars)

    print("=" * 60)
    print("(4) 描边回归 + 非纯色描边预检")
    x0, y0 = int(3.5 * 96), int(2.4 * 96)
    sub = arr[y0:y0 + 96, x0:x0 + 400].astype(int)
    red = ((sub[:, :, 0] > 170) & (sub[:, :, 1] < 90) & (sub[:, :, 2] < 90)).sum()
    check(f'红色描边像素 {red} > 300', red > 300)
    issues = RenderPrecheck(installed_fonts=None).check_pptx(path)
    codes = [i.code for i in issues]
    check('非纯色文字描边报 TEXT_STROKE', 'TEXT_STROKE' in codes)
    check('无其他误报（除 NO_WRAP）',
          all(c in ('NO_WRAP', 'TEXT_STROKE') for c in codes))

    print("=" * 60)
    print("(5) COM 交叉验证：删除线行位置（±4px）")
    try:
        sys.path.insert(0, os.path.abspath(os.path.join(
            os.path.dirname(__file__), '..', '..', '..')))
        from src.core.processors.com_image_exporter import COMImageExporter
        ref_dir = tempfile.mkdtemp()
        COMImageExporter().export_to_images(path, ref_dir, 'PNG', 1.0)
        ref = np.array(Image.open(os.path.join(ref_dir, 'slide_1.png')).convert('RGB'))

        def find_strike_row(a, y_guess, gx, em):
            # 找删除线级别的连续粗带（>=60px），返回带的中点行
            def cont_run(y):
                row = (a[y, int(gx) + 15: int(gx) + em - 15].sum(axis=1) < 500)
                best = cur = 0
                for v in row:
                    cur = cur + 1 if v else 0
                    best = max(best, cur)
                return best

            band = [y for y in range(y_guess - 6, y_guess + 7) if cont_run(y) >= 60]
            return (band[0] + band[-1]) // 2 if band else None

        ref_y = find_strike_row(ref, strike_y, int(3.5 * 96 + 9.6), em)
        poc_y = find_strike_row(arr, strike_y, int(3.5 * 96 + 9.6), em)
        if ref_y is not None and poc_y is not None:
            check(f'删除线行位置 PoC={poc_y} vs PowerPoint={ref_y}（±4px）',
                  abs(ref_y - poc_y) <= 4)
        else:
            check(f'两边都找到删除线 {ref_y} {poc_y}', False)
    except Exception as e:
        print(f"  ⚠️ COM 交叉验证跳过: {e}")

    print("=" * 60)
    print(f"结果: {PASS} 通过, {FAIL} 失败")
    sys.exit(1 if FAIL else 0)


if __name__ == '__main__':
    main()
