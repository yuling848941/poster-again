# -*- coding: utf-8 -*-
"""
render_core（自包含渲染核心包）单元测试
验证：render() 一句话 API（路径/bytes 输入）、RenderResult 内容、
以及【安卓模拟】自包含性——把 render_core（含捆绑字体）+ fontfx 拷到
独立目录，在干净子进程（无仓库、无 pip 安装、bundled_first 策略）中
渲染，输出与仓库内渲染逐像素一致。

运行：.venv/Scripts/python.exe libs/render_core/tests/test_render_core.py
"""
import os
import shutil
import subprocess
import sys
import tempfile

_HERE = os.path.dirname(os.path.abspath(__file__))
# 本库（render_core）与它依赖的 fontfx 都在仓库里；已 pip 安装时这几行是冗余的
for _p in (os.path.join(_HERE, '..', 'src'),
           os.path.join(_HERE, '..', '..', 'fontfx', 'src')):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from PIL import Image
import numpy as np
from lxml import etree
from pptx import Presentation
from pptx.util import Pt, Inches

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

from render_core import render, RenderResult, font_registry

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


def build_sample_pptx(path):
    """合成含思源黑体文本 + 渐变 + 多行距的样本（覆盖仿射基线与字体解析）"""
    prs = Presentation()
    prs.slide_width = Inches(7.5)
    prs.slide_height = Inches(10.0)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(6.5), Inches(2.0))
    tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = 1.2
    run = p.add_run()
    run.text = 'render_core 自包含测试'
    run.font.size = Pt(44)
    run.font.bold = True
    run.font.name = '思源黑体'
    rPr = run._r.get_or_add_rPr()
    rPr.set('spc', '200')
    rPr.append(etree.fromstring(
        f'<a:gradFill xmlns:a="{A}"><a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="C00000"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="F9CE6F"/></a:gs>'
        f'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'))

    tb2 = slide.shapes.add_textbox(Inches(0.5), Inches(3.0), Inches(6.5), Inches(3.0))
    tb2.text_frame.word_wrap = True
    p2 = tb2.text_frame.paragraphs[0]
    p2.line_spacing = 1.5
    run2 = p2.add_run()
    run2.text = '第一行多行距文本ABC第二行内容defgh第三行文字ijklmn第四行尾行'
    run2.font.size = Pt(28)
    run2.font.name = '思源宋体'

    prs.save(path)


def main():
    tmp = tempfile.mkdtemp(prefix='render_core_t_')
    pptx_path = os.path.join(tmp, 'sample.pptx')
    build_sample_pptx(pptx_path)

    print('(1) render() 路径输入')
    font_registry.set_policy('bundled_first')   # 统一用捆绑思源（安卓实际路径）
    out1 = os.path.join(tmp, 'out_path')
    r1 = render(pptx_path, out1, 'PNG', 1.0)
    check('生成 1 张图片', len(r1.images) == 1 and os.path.exists(r1.images[0]))
    check('干净模板预检零问题', len(r1.issues) == 0)
    img1 = np.array(Image.open(r1.images[0]).convert('RGB'))
    check('画布有墨迹', (np.abs(img1.astype(int) - 255).sum(axis=2) > 150).sum() > 500)

    print('(2) render() bytes 输入（安卓 assets 场景）')
    with open(pptx_path, 'rb') as f:
        data = f.read()
    out2 = os.path.join(tmp, 'out_bytes')
    r2 = render(data, out2, 'PNG', 1.0)
    img2 = np.array(Image.open(r2.images[0]).convert('RGB'))
    check('bytes 与路径渲染逐像素一致', bool((img1 == img2).all()))

    print('(3) 安卓模拟：独立目录 + 干净子进程 + bundled_first')
    iso = os.path.join(tmp, 'isolated')
    pkg = os.path.join(iso, 'render_core')
    lib_src = os.path.abspath(os.path.join(_HERE, '..', 'src'))
    shutil.copytree(os.path.join(lib_src, 'render_core'), pkg,
                    ignore=shutil.ignore_patterns('__pycache__'))
    shutil.copytree(os.path.abspath(os.path.join(_HERE, '..', '..', 'fontfx', 'src', 'fontfx')),
                    os.path.join(iso, 'fontfx'), ignore=shutil.ignore_patterns('__pycache__'))
    check('拷贝了 5 个引擎模块 + __init__ + 捆绑字体 + fontfx',
          os.path.exists(os.path.join(pkg, '__init__.py'))
          and all(os.path.exists(os.path.join(pkg, m)) for m in (
              'pillow_image_exporter.py', 'format_inheritance.py',
              'shape_geometry.py', 'render_precheck.py', 'font_registry.py'))
          and os.path.exists(os.path.join(pkg, 'resources', 'fonts', 'SourceHanSansCN-Regular.otf'))
          and os.path.exists(os.path.join(iso, 'fontfx', '__init__.py')))

    out3 = os.path.join(tmp, 'out_iso')
    script = (
        "import sys, os\n"
        f"sys.path.insert(0, {iso!r})\n"
        "from render_core import render\n"
        f"r = render({pptx_path!r}, {out3!r}, 'PNG', 1.0)\n"
        "assert len(r.images) == 1 and len(r.issues) == 0, (r.images, r.issues)\n"
        "print('ISO_OK')\n"
    )
    env = dict(os.environ)
    env['POSTERAGAIN_FONT_POLICY'] = 'bundled_first'
    env['PYTHONPATH'] = ''  # 不继承仓库路径，验证自包含
    proc = subprocess.run([sys.executable, '-c', script], env=env,
                          capture_output=True, text=True, cwd=iso)
    check('子进程渲染成功（无仓库依赖）', 'ISO_OK' in proc.stdout)
    if 'ISO_OK' not in proc.stdout:
        print('    stderr:', proc.stderr[-500:])
    else:
        img3 = np.array(Image.open(os.path.join(out3, 'slide_1.png')).convert('RGB'))
        check('隔离环境与仓库渲染逐像素一致', bool((img1 == img3).all()))

    font_registry.set_policy('auto')
    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
