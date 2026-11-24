# -*- coding: utf-8 -*-
"""
个性化引擎（personalize）单元测试
验证：三种占位符识别、替换语义与 pptx_processor 一致、
恢复后可复用（双行不同值互不污染）、render_row 端到端出图、
表头匹配（精确/去前缀/忽略大小写）。

运行：.venv/Scripts/python.exe libs/render_core/tests/test_personalize.py
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

from pptx import Presentation
from pptx.util import Pt, Inches

from render_core import Personalizer, match_fields
from render_core.personalize import PLACEHOLDER_PREFIX

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


def build_template(path):
    """三种植位符各一：形状名 ph_name / AltText ph_city / 行内 {{amount}}"""
    prs = Presentation()
    prs.slide_width = Inches(7.5)
    prs.slide_height = Inches(5.0)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tb1 = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(6.5), Inches(1.2))
    tb1.text_frame.word_wrap = True
    r = tb1.text_frame.paragraphs[0].add_run()
    r.text = '张三'
    r.font.size = Pt(40)
    tb1.name = 'ph_name'   # 占位符方式2：形状名

    tb2 = slide.shapes.add_textbox(Inches(0.5), Inches(2.0), Inches(6.5), Inches(1.2))
    r2 = tb2.text_frame.paragraphs[0].add_run()
    r2.text = '广州'
    r2.font.size = Pt(30)
    cNvPr = tb2._element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr')
    if cNvPr is None:  # p 命名空间
        cNvPr = tb2._element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr')
    cNvPr.set('descr', 'ph_city')   # 占位符方式1：Alt Text（WPS/PPT 兼容路径）

    tb3 = slide.shapes.add_textbox(Inches(0.5), Inches(3.4), Inches(6.5), Inches(1.2))
    r3 = tb3.text_frame.paragraphs[0].add_run()
    r3.text = '保费 {{amount}} 元'
    r3.font.size = Pt(30)

    prs.save(path)


def main():
    tmp = tempfile.mkdtemp(prefix='personalize_t_')
    tpl = os.path.join(tmp, 'tpl.pptx')
    build_template(tpl)

    print('(1) 占位符识别')
    pz = Personalizer(tpl)
    fields = pz.fields()
    check('识别全部三种植位符',
          set(fields) == {'ph_name', 'ph_city', 'ph_amount'})

    print('(2) 替换 + 恢复复用（核心：同一 Presentation 逐行不互相污染）')
    pz.apply({'ph_name': '李逵', 'ph_city': '深圳', 'ph_amount': '88,888'})
    xml1 = pz._prs.slides[0]._element
    txt1 = ''.join(xml1.itertext())
    check('行1 替换生效', '李逵' in txt1 and '深圳' in txt1 and '88,888' in txt1
          and '张三' not in txt1)
    pz.restore()
    txt0 = ''.join(pz._prs.slides[0]._element.itertext())
    check('恢复模板原状', '张三' in txt0 and '李逵' not in txt0 and '{{amount}}' in txt0)
    pz.apply({'ph_name': '王小明', 'ph_city': '东莞', 'ph_amount': '12,345'})
    txt2 = ''.join(pz._prs.slides[0]._element.itertext())
    check('行2 替换生效且无行1残留', '王小明' in txt2 and '李逵' not in txt2)

    print('(3) render_row 端到端（替换->渲染->自动恢复）')
    import numpy as np
    from PIL import Image
    out1 = os.path.join(tmp, 'row1')
    img1 = pz.render_row({'ph_name': '李逵', 'ph_city': '深圳', 'ph_amount': '88,888'}, out1)
    out2 = os.path.join(tmp, 'row2')
    img2 = pz.render_row({'ph_name': '王小明', 'ph_city': '东莞', 'ph_amount': '12,345'}, out2)
    check('两行各出一张图', len(img1) == 1 and len(img2) == 1)
    a = np.array(Image.open(img1[0]).convert('RGB'))
    b = np.array(Image.open(img2[0]).convert('RGB'))
    check('两行渲染结果不同（个性化生效）', bool((a != b).any()))
    txt0 = ''.join(pz._prs.slides[0]._element.itertext())
    check('render_row 后模板已恢复', '张三' in txt0)

    print('(4) 字节输入（安卓场景）')
    with open(tpl, 'rb') as f:
        pz2 = Personalizer(f.read())
    check('字节构造 + fields 正常', set(pz2.fields()) == {'ph_name', 'ph_city', 'ph_amount'})

    print('(5) 表头匹配')
    m = match_fields(['ph_name', 'ph_city', 'ph_amount'],
                     ['Name', 'CITY', '金额'])
    check('忽略大小写', m.get('ph_city') == 'CITY')
    m0 = match_fields(['ph_name'], ['ph_name'])
    check('精确匹配', m0.get('ph_name') == 'ph_name')
    m2 = match_fields(['ph_name'], ['name'])
    check('去 ph_ 前缀匹配', m2.get('ph_name') == 'name')
    m3 = match_fields(['ph_amount'], ['金额'])
    check('无匹配不产生条目', 'ph_amount' not in m3)

    print('(6) 输出文件基础名（与 PC ppt_generator 同规则）')
    from render_core.personalize import row_base_name
    headers = ['日期 ID', '支公司 DESC', '营销员或推广员 姓名', '标保']
    row = {'营销员或推广员 姓名': '陈春兰', '标保': '20097'}
    check('含"姓名"列', row_base_name(headers, row, 0) == '1_陈春兰')
    check('序号从1起', row_base_name(headers, row, 3) == '4_陈春兰')
    h2 = ['姓名', '家族']
    check('精确姓名列优先', row_base_name(h2, {'姓名': '李逵'}, 1) == '2_李逵')
    check('无姓名列兜底', row_base_name(['家族', '标保'], {'家族': '刘'}, 0) == '1_output')
    check('空值兜底', row_base_name(h2, {'姓名': ''}, 0) == '1_output')
    check('非法字符清洗', row_base_name(h2, {'姓名': '张:三*'}, 0) == '1_张_三_')
    check('首尾点空清理(仅整体首尾)', row_base_name(h2, {'姓名': '.王五.'}, 0) == '1_.王五')

    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
