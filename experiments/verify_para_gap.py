# -*- coding: utf-8 -*-
"""验证多段落/空段落行距：PoC vs PowerPoint COM 基线实测"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'tests', 'unit'))

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

from render_core.pillow_image_exporter import PillowImageExporter


def add_para(tf, text, sz=3600):
    from lxml import etree
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    p = tf.add_paragraph() if tf.paragraphs[0].runs or len(tf.paragraphs) > 1 else tf.paragraphs[0]
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
    # 框1：两个非空段落
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(4), Inches(2.5))
    tb.text_frame.word_wrap = True
    add_para(tb.text_frame, '第一段落')
    add_para(tb.text_frame, '第二段落')
    # 框2：段1 + 空段 + 段3
    tb = slide.shapes.add_textbox(Inches(4.5), Inches(0.3), Inches(4), Inches(2.5))
    tb.text_frame.word_wrap = True
    add_para(tb.text_frame, '上段落')
    add_para(tb.text_frame, '')
    add_para(tb.text_frame, '下段落')
    path = os.path.join(tempfile.mkdtemp(), 'para.pptx')
    prs.save(path)
    return path


def para_tops(arr, x0, x1, y0, y1):
    """按行投影分段，返回每段文字的顶部 y"""
    m = (np.abs(arr[y0:y1, x0:x1].astype(int) - 255).sum(axis=2) > 150).any(axis=1)
    segs, start = [], None
    for i, v in enumerate(m):
        if v and start is None:
            start = i
        elif not v and start is not None:
            segs.append((y0 + start, y0 + i))
            start = None
    if start is not None:
        segs.append((y0 + start, y0 + len(m)))
    # 合并近邻（<6px）
    merged = []
    for s in segs:
        if merged and s[0] - merged[-1][1] < 6:
            merged[-1] = (merged[-1][0], s[1])
        else:
            merged.append(list(s) if False else s)
    return merged


def main():
    path = build()
    out = tempfile.mkdtemp()
    files = PillowImageExporter().export_to_images(path, out, 'PNG', 1.0)
    poc = np.array(Image.open(files[0]).convert('RGB'))

    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'poster-again')))
    from src.core.processors.com_image_exporter import COMImageExporter
    ref_dir = tempfile.mkdtemp()
    COMImageExporter().export_to_images(path, ref_dir, 'PNG', 1.0)
    ref = np.array(Image.open(os.path.join(ref_dir, 'slide_1.png')).convert('RGB'))

    for name, a in (('PoC', poc), ('PowerPoint', ref)):
        s1 = para_tops(a, 48, 430, 20, 270)
        s2 = para_tops(a, 432, 830, 20, 270)
        gap1 = s1[1][0] - s1[0][0] if len(s1) >= 2 else None
        gap2 = s2[-1][0] - s2[0][0] if len(s2) >= 2 else None
        print(f'{name}: 两段落间距={gap1}px  含空段间距={gap2}px  分段数: {len(s1)}/{len(s2)}')


if __name__ == '__main__':
    main()
