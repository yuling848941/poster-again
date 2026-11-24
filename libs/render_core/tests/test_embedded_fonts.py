# -*- coding: utf-8 -*-
"""
PPTX 内嵌字体（embeddedFontLst/fntdata）支持测试
================================================================
1. 手工构造嵌入字体的 pptx（字体数据 = 捆绑思源黑体，typeface 用系统没有的名字）
2. load_embedded_fonts 解包注册
3. resolve 命中 embedded；系统已有字体仍优先（PowerPoint 语义）
4. 渲染一致：嵌入字体渲染 == 直接用同一字体文件渲染
5. ODTTF 解混淆路径（前 32 字节 GUID 异或）

运行：.venv/Scripts/python.exe libs/render_core/tests/test_embedded_fonts.py
"""
import io
import os
import re
import sys
import tempfile
import uuid
import zipfile

_HERE = os.path.dirname(os.path.abspath(__file__))
# 本库（render_core）与它依赖的 fontfx 都在仓库里；已 pip 安装时这几行是冗余的
for _p in (os.path.join(_HERE, '..', 'src'),
           os.path.join(_HERE, '..', '..', 'fontfx', 'src')):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from pptx import Presentation
from pptx.util import Inches
from lxml import etree

from render_core import font_registry

PASS = 0
FAIL = 0
_BUNDLED = font_registry.bundled_font_dir() or ''   # 捆绑字体目录（由库自己定位）


def check(desc, cond):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  ✅ {desc}")
    else:
        FAIL += 1
        print(f"  ❌ {desc}")


FONT_FILE = os.path.join(_BUNDLED, 'SourceHanSansCN-Regular.otf')
FONT_BOLD = os.path.join(_BUNDLED, 'SourceHanSansCN-Bold.otf')
FACE = '测试嵌入字体XYZ'      # 系统一定没有的名字
FACE_SYS = '微软雅黑'          # 系统已有


def build_embedded_pptx(obfuscate=False, embed_sys_font=False):
    """最简 pptx + 注入 embeddedFontLst/fntdata/rels/[Content_Types]"""
    face = FACE_SYS if embed_sys_font else FACE
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(2))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = '嵌入字体测试字ABC'
    rPr = run._r.get_or_add_rPr()
    rPr.append(etree.fromstring(
        f'<a:latin xmlns:a="{A}" typeface="{face}"/>'))
    rPr.append(etree.fromstring(
        f'<a:ea xmlns:a="{A}" typeface="{face}"/>'))
    buf = io.BytesIO()
    prs.save(buf)

    guid = str(uuid.uuid4()).upper()
    src = io.BytesIO(buf.getvalue())
    out = io.BytesIO()

    # 先准备字体部件（名/数据），rels 引用要用同一批名字
    parts = {}
    for style, fname in (('regular', FONT_FILE), ('bold', FONT_BOLD)):
        data = open(fname, 'rb').read()
        if obfuscate:
            g = uuid.uuid4()  # 每个部件独立 GUID
            gb = g.bytes
            head = bytearray(data[:32])
            for i in range(32):
                head[i] ^= gb[i % 16]
            parts[style] = (f'ppt/fonts/{str(g).upper()}.odttf',
                            bytes(head) + data[32:])
        else:
            parts[style] = (f'ppt/fonts/font_{style}.fntdata', data)

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(
            out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item == 'ppt/presentation.xml':
                P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
                R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                root = etree.fromstring(data)
                lst = etree.SubElement(root, f'{{{P}}}embeddedFontLst')
                ef = etree.SubElement(lst, f'{{{P}}}embeddedFont')
                etree.SubElement(ef, f'{{{P}}}font', typeface=face)
                for style in parts:
                    el = etree.SubElement(ef, f'{{{P}}}{style}')
                    el.set(f'{{{R}}}id', f'rIdFont{style}')
                data = etree.tostring(root, xml_declaration=True,
                                      encoding='UTF-8', standalone=True)
            elif item == 'ppt/_rels/presentation.xml.rels':
                RNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
                root = etree.fromstring(data)
                for style, (part_name, _) in parts.items():
                    rel = etree.SubElement(root, f'{{{RNS}}}Relationship')
                    rel.set('Id', f'rIdFont{style}')
                    rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font')
                    rel.set('Target', part_name[len('ppt/'):])
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8')
            elif item == '[Content_Types].xml':
                CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
                root = etree.fromstring(data)
                for ext in ('fntdata', 'odttf'):
                    d = etree.SubElement(root, f'{{{CT}}}Default')
                    d.set('Extension', ext)
                    d.set('ContentType', 'application/x-fontdata')
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8')
            zout.writestr(item, data)
        for part_name, data in parts.values():
            zout.writestr(part_name, data)
    out.seek(0)
    return out.getvalue()


def main():
    global _unused
    print("(1) load_embedded_fonts 注册与 resolve 命中")
    pptx_bytes = build_embedded_pptx()
    font_registry.clear_embedded_fonts()
    n = font_registry.load_embedded_fonts(pptx_bytes)
    check(f'注册 2 个条目（regular+bold）: 实得 {n}', n == 2)
    spec = font_registry.resolve(FACE, False)
    check(f'未安装字体名命中 embedded (source={spec.source})', spec.source == 'embedded')
    spec_b = font_registry.resolve(FACE, True)
    check('bold 命中嵌入 bold 文件', spec_b.source == 'embedded'
          and spec_b.path != spec.path)
    check('is_resolvable 含嵌入', font_registry.is_resolvable(FACE))
    check('重复加载同文档为空操作', font_registry.load_embedded_fonts(pptx_bytes) == 2)

    print("(2) 系统已装字体优先（PowerPoint 语义，Windows 零回归）")
    pptx_sys = build_embedded_pptx(embed_sys_font=True)
    font_registry.clear_embedded_fonts()
    font_registry.load_embedded_fonts(pptx_sys)
    spec = font_registry.resolve(FACE_SYS, False)
    check(f'微软雅黑仍走 system (source={spec.source})',
          spec.source == 'system')

    print("(3) ODTTF 解混淆")
    pptx_odttf = build_embedded_pptx(obfuscate=True)
    font_registry.clear_embedded_fonts()
    n = font_registry.load_embedded_fonts(pptx_odttf)
    check(f'混淆字体解出注册 {n} 条', n == 2)
    spec = font_registry.resolve(FACE, False)
    check('解混淆后可命中', spec.source == 'embedded')

    print("(4) 渲染一致（嵌入字体 == 直接字体文件）")
    import logging
    logging.basicConfig(level=logging.ERROR)
    from PIL import Image, ImageDraw, ImageFont
    import numpy as np
    from render_core import render
    font_registry.clear_embedded_fonts()
    font_registry.load_embedded_fonts(pptx_bytes)
    r = render(pptx_bytes, tempfile.mkdtemp(prefix='emb_'), 'PNG', 1.0)
    ours = np.asarray(Image.open(r.images[0]).convert('L'), dtype=np.int16)
    # 直接用同一字体文件渲染参照（18pt @96dpi = 24px）
    W, H = ours.shape[1], ours.shape[0]
    ref = Image.new('L', (W, H), 255)
    d = ImageDraw.Draw(ref)
    f = ImageFont.truetype(FONT_FILE, 24)
    d.text((48, 48), '嵌入字体测试字ABC', font=f, fill=0)
    ref = np.asarray(ref, dtype=np.int16)
    # 引擎排版含内边距与基线模型，参照画在裸坐标 -> 质心对齐后比较字形（位置无关）
    o_ink = ours < 128
    r_ink = ref < 128
    oy, ox = np.where(o_ink)
    ry, rx = np.where(r_ink)
    dy = int(round(float(ry.mean() - oy.mean())))
    dx = int(round(float(rx.mean() - ox.mean())))
    o_shift = np.roll(np.roll(o_ink, dy, axis=0), dx, axis=1)
    inter = (o_shift & r_ink).sum()
    iou = inter / max(1, int(o_ink.sum() + r_ink.sum() - inter))
    check(f'字形 IoU {iou:.2f} > 0.90（嵌入字体真实生效，质心对齐）', iou > 0.90)

    print("(5) 清理")
    font_registry.clear_embedded_fonts()
    check('清空后不再命中', font_registry.resolve(FACE, False).source != 'embedded')

    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
