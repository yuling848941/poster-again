# -*- coding: utf-8 -*-
"""
渲染预检（render_precheck）单元测试
验证：白名单预检能准确识别各类不支持元素；真实模板零误报；
渲染器背景/渐变新能力正常工作。

运行：.venv/Scripts/python.exe tests/unit/test_render_precheck.py
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

try:
    from src.core.processors.render_precheck import RenderPrecheck, PrecheckIssue
    from src.core.processors.pillow_image_exporter import PillowImageExporter
except ImportError:
    from core.processors.render_precheck import RenderPrecheck, PrecheckIssue
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


def style_run(run, sz=2800, fill=True):
    """给 run 写入显式格式，构造'干净'文本"""
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', str(sz))
    rPr.set('b', '1')
    if fill:
        rPr.append(etree.fromstring(
            f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="FF0000"/></a:solidFill>'))
        rPr.append(etree.fromstring(
            f'<a:latin xmlns:a="{A}" typeface="微软雅黑"/>'))


def set_wrap_square(tb):
    """python-pptx 新建文本框默认 wrap="none"，统一改为 square 消除无关告警"""
    tb.text_frame.word_wrap = True


def build_test_pptx():
    """构造包含各类不支持元素的合成 PPT"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 1) 干净文本框（显式格式）—— 不应产生任何问题
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(4), Inches(0.8))
    set_wrap_square(tb)
    style_run(tb.text_frame.paragraphs[0].add_run())
    tb.text_frame.paragraphs[0].runs[0].text = '干净文本'

    # 2) 旋转文本框 -> ROTATION
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4), Inches(0.8))
    set_wrap_square(tb)
    style_run(tb.text_frame.paragraphs[0].add_run())
    tb.text_frame.paragraphs[0].runs[0].text = '旋转'
    tb._element.spPr.xfrm.set('rot', '4500000')

    # 3) anchor 居中 -> ANCHOR
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(2.1), Inches(4), Inches(0.8))
    set_wrap_square(tb)
    style_run(tb.text_frame.paragraphs[0].add_run())
    tb.text_frame.paragraphs[0].runs[0].text = '居中锚'
    from pptx.oxml.ns import qn as _qn
    bodyPr = tb.text_frame._txBody.find(_qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('anchor', 'ctr')

    # 4) 未书写任何格式的 run -> INHERIT_RPR
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(3.0), Inches(4), Inches(0.8))
    set_wrap_square(tb)
    tb.text_frame.paragraphs[0].add_run().text = '完全继承'

    # 5) 有颜色没字号 -> INHERIT_SIZE
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(3.9), Inches(4), Inches(0.8))
    set_wrap_square(tb)
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = '只写颜色'
    rPr = run._r.get_or_add_rPr()
    rPr.set('b', '1')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="00FF00"/></a:solidFill>'))

    # 6) 表格 -> TABLE
    gf = slide.shapes.add_table(2, 2, Inches(5), Inches(0.3), Inches(3), Inches(1.5))

    # 7) 组合 -> GROUP
    grp = slide.shapes.add_group_shape()
    grp.shapes.add_textbox(Inches(5), Inches(2.1), Inches(3), Inches(0.8))
    grp.shapes.add_textbox(Inches(5), Inches(3.0), Inches(3), Inches(0.8))

    # 8) 未知字体 -> FONT_MISSING
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(4), Inches(0.8))
    set_wrap_square(tb)
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = '怪字体'
    rPr = run._r.get_or_add_rPr()
    rPr.set('sz', '2800')
    rPr.append(etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="0000FF"/></a:solidFill>'))
    rPr.append(etree.fromstring(
        f'<a:latin xmlns:a="{A}" typeface="不存在的字体XYZ"/>'))

    # 9) normAutofit 缩字 -> AUTOFIT_SCALE
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(5.7), Inches(4), Inches(0.8))
    set_wrap_square(tb)
    style_run(tb.text_frame.paragraphs[0].add_run())
    tb.text_frame.paragraphs[0].runs[0].text = '自动缩字'
    bodyPr = tb.text_frame._txBody.find(f'{{{A}}}bodyPr')
    etree.SubElement(bodyPr, f'{{{A}}}normAutofit', fontScale='80000')

    path = os.path.join(tempfile.mkdtemp(prefix='precheck_'), 'test.pptx')
    prs.save(path)
    return path


def build_bg_pptx():
    """构造带线性渐变背景的 PPT（验证新背景渲染能力）"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cSld = slide._element.find(f'{{{P}}}cSld')
    bg = etree.SubElement(cSld, f'{{{P}}}bg')
    bgPr = etree.SubElement(bg, f'{{{P}}}bgPr')
    gradFill = etree.fromstring(
        f'<a:gradFill xmlns:a="{A}" xmlns:r="{R}">'
        '<a:gsLst>'
        '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>'
        '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>'
        '</a:gsLst>'
        '<a:lin ang="0" scaled="0"/>'
        '</a:gradFill>')
    bgPr.append(gradFill)
    etree.SubElement(bgPr, f'{{{A}}}effectLst')
    tb = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(6), Inches(1))
    tb.text_frame.word_wrap = True
    style_run(tb.text_frame.paragraphs[0].add_run())
    tb.text_frame.paragraphs[0].runs[0].text = '渐变背景上的文字'
    path = os.path.join(tempfile.mkdtemp(prefix='precheck_bg_'), 'bg.pptx')
    prs.save(path)
    return path


def main():
    print("=" * 60)
    print("(1) 合成 PPT 预检（各类不支持元素应被识别）")
    path = build_test_pptx()
    # installed_fonts=None 跳过字体检查，单独用例测字体
    issues = RenderPrecheck(installed_fonts=None).check_pptx(path)
    codes = [i.code for i in issues]

    check('ROTATION 被识别', 'ROTATION' in codes)
    check('ANCHOR 被识别', 'ANCHOR' in codes)
    check('TABLE 被识别', 'TABLE' in codes)
    check('GROUP 被识别', 'GROUP' in codes)
    check('AUTOFIT_SCALE 已支持不再告警', 'AUTOFIT_SCALE' not in codes)
    # 继承解析器接管后：无格式 run 通过母版 otherStyle 解析（18pt 黑），
    # 与 PowerPoint 行为一致，不再告警
    check('继承类问题零误报（解析器已支持格式继承）',
          not any(i.code.startswith('INHERIT') for i in issues))
    check('无未解析继承链标记', 'INHERIT_UNRESOLVED' not in codes)

    print("=" * 60)
    print("(2) 字体检查")
    issues = RenderPrecheck(
        installed_fonts={'微软雅黑', 'simsun.ttc', '宋体'}).check_pptx(path)
    codes = [i.code for i in issues]
    check('FONT_MISSING 被识别（不存在的字体XYZ）', 'FONT_MISSING' in codes)
    check('已安装字体（微软雅黑）不误报（消息中仅回退提示提及）',
          not any(i.message.split('未安装')[0].strip().endswith("微软雅黑'")
                  for i in issues if i.code == 'FONT_MISSING'))

    print("=" * 60)
    print("(3) 渐变背景：预检不误报 + 渲染成功")
    bg_path = build_bg_pptx()
    issues = RenderPrecheck(installed_fonts=None).check_pptx(bg_path)
    check('渐变背景(bgPr+lin)零预检问题', len(issues) == 0)
    out_dir = tempfile.mkdtemp(prefix='precheck_render_')
    files = PillowImageExporter().export_to_images(bg_path, out_dir, 'PNG', 1.0)
    check('背景渲染输出文件', len(files) == 1 and os.path.getsize(files[0]) > 5000)
    from PIL import Image
    import numpy as np
    arr = np.array(Image.open(files[0]).convert('RGB'))
    left_red = arr[:, :50].mean(axis=(0, 1))
    right_blue = arr[:, -50:].mean(axis=(0, 1))
    check(f'左红右蓝渐变生效 左={left_red.round(0)} 右={right_blue.round(0)}',
          left_red[0] > 200 and left_red[2] < 80 and right_blue[2] > 200 and right_blue[0] < 80)

    print("=" * 60)
    print("(4) 真实模板预检（递交/承保应零问题；每日贺报空占位符不误报）")
    root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    alt_root = os.path.abspath(os.path.join(root, '..', 'poster-again'))
    for tpl in ['10月递交贺报(3).pptx', '10月承保贺报(3).pptx', '01、每日贺报 - 发2 .pptx']:
        tpl_path = os.path.join(root, tpl)
        if not os.path.exists(tpl_path) and os.path.exists(os.path.join(alt_root, tpl)):
            tpl_path = os.path.join(alt_root, tpl)
        if not os.path.exists(tpl_path):
            print(f"  ⚠️ 模板不存在，跳过: {tpl}")
            continue
        issues = RenderPrecheck(installed_fonts=None).check_pptx(tpl_path)
        status = '零问题 ✅' if not issues else f'{len(issues)} 项问题'
        print(f"  {tpl}: {status}")
        for i in issues[:5]:
            print(f"    [{i.code}] 第{i.slide}页 {i.shape!r}: {i.message}")
        if '递交' in tpl or '承保' in tpl:
            check(f'{tpl} 零误报（白名单必须覆盖现有模板家族）', len(issues) == 0)

    print("=" * 60)
    print(f"结果: {PASS} 通过, {FAIL} 失败")
    sys.exit(1 if FAIL else 0)


if __name__ == '__main__':
    main()
