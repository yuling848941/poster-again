# -*- coding: utf-8 -*-
"""模板「恢复」与「重新加载」等价性测试。

批量生成现在逐行复用同一 Presentation：替换 -> 导出/保存 -> restore_template，
而不是每行重新 load_template。本测试确认两者完全等价，杜绝数据串行（H2）：

1. 替换后 restore，幻灯片 XML 与原始快照逐字节一致；
2. 多行批量：restore 循环 与 reload 循环 产出的 PPT 部件逐字节一致；
3. 多行批量图片：两种循环逐像素一致，且各行内容互不污染；
4. 真实模板（若存在）同样验证。

运行：.venv/Scripts/python.exe tests/unit/test_template_restore.py
"""
import os
import sys
import tempfile
import shutil
import zipfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
os.environ['POSTERAGAIN_IMAGE_ENGINE'] = 'python'   # 只测内置引擎，不驱动 Office

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image, ImageChops

from src.core.processors.pptx_processor import PPTXProcessor

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
    """多形状模板：alt-text 占位符 + {{标记}} + 多段落 + 组合，覆盖各种替换路径"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(8), Inches(6)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4), Inches(1))
    tb._element.nvSpPr.cNvPr.set('descr', 'ph_姓名')
    p = tb.text_frame.paragraphs[0]
    p.add_run().text = '姓名占位'
    p2 = tb.text_frame.add_paragraph()
    p2.add_run().text = '第二段会被清空'

    tb2 = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(6), Inches(1))
    r = tb2.text_frame.paragraphs[0].add_run()
    r.text = '金额：{{保费}} 元，日期 {{日期}}'
    r.font.size = Pt(28)

    tb3 = slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(4), Inches(1))
    tb3.name = 'ph_标题'
    tb3.text_frame.paragraphs[0].add_run().text = '标题占位'

    prs.save(path)
    return path


def slide_xmls(prs):
    return [etree.tostring(s._element) for s in prs.slides]


def pkg_parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def test_restore_equals_original(tpl):
    proc = PPTXProcessor()
    proc.load_template(tpl)
    original = slide_xmls(proc.presentation)
    fields = proc.find_placeholders()
    if not fields:
        check('无占位符模板：restore 为无害空操作', proc.restore_template() is True)
        return
    data = {f: f'{f[3:]}测试值' for f in fields}
    proc.replace_placeholders(data)
    check('替换确实改动了幻灯片', slide_xmls(proc.presentation) != original)
    proc.restore_template()
    check('restore 后幻灯片 XML 与原始逐字节一致', slide_xmls(proc.presentation) == original)
    # 再替换再恢复，反复多次仍一致
    for i in range(3):
        proc.replace_placeholders({f: f'{f[3:]}第{i}轮' for f in fields})
        proc.restore_template()
    check('反复 替换->恢复 后仍与原始一致', slide_xmls(proc.presentation) == original)


def make_rows(tpl):
    """按模板真实占位符构造 3 行数据，每个字段逐行取不同值（用于污染检查）。
    返回 (rows, name_key)；name_key 取第一个占位符。"""
    proc = PPTXProcessor()
    proc.load_template(tpl)
    fields = proc.find_placeholders()
    if not fields:
        return [], None
    rows = []
    for i in range(3):
        rows.append({f: f'{f[3:]}值{i}' for f in fields})
    return rows, fields[0]


def test_batch_ppt_restore_vs_reload(tpl, rows, name_key):
    if not rows:
        print('  （模板无占位符，跳过批量 PPT 对比）')
        return
    out = tempfile.mkdtemp(prefix='restore_ppt_')
    try:
        # restore 循环
        proc = PPTXProcessor()
        proc.load_template(tpl)
        restore_files = []
        for i, data in enumerate(rows):
            proc.replace_placeholders(data)
            p = os.path.join(out, f'restore_{i}.pptx')
            proc.save_presentation(p)
            restore_files.append(p)
            proc.restore_template()
        # reload 循环（旧行为）
        proc2 = PPTXProcessor()
        reload_files = []
        for i, data in enumerate(rows):
            proc2.load_template(tpl)
            proc2.replace_placeholders(data)
            p = os.path.join(out, f'reload_{i}.pptx')
            proc2.save_presentation(p)
            reload_files.append(p)
        for i in range(len(rows)):
            a, b = pkg_parts(restore_files[i]), pkg_parts(reload_files[i])
            # zip 成员名一致、各部件内容逐字节一致（压缩方式可能不同，比的是解压后内容）
            check(f'第{i+1}行 PPT 部件逐字节一致（restore vs reload）', a == b)
        # 各行互不污染：第 i 行含第 i 行取值、不含别行取值
        for i, data in enumerate(rows):
            parts = pkg_parts(restore_files[i])
            slide = parts['ppt/slides/slide1.xml'].decode('utf-8')
            mine = data[name_key]
            others = [d[name_key] for j, d in enumerate(rows) if j != i]
            check(f'第{i+1}行只含本行取值 {mine}', mine in slide and all(o not in slide for o in others))
    finally:
        shutil.rmtree(out, ignore_errors=True)


def test_batch_image_restore_vs_reload(tpl, rows, name_key):
    if not rows:
        print('  （模板无占位符，跳过批量图片对比）')
        return
    out = tempfile.mkdtemp(prefix='restore_img_')
    try:
        proc = PPTXProcessor()
        proc.load_template(tpl)
        restore_imgs = []
        for i, data in enumerate(rows):
            proc.replace_placeholders(data)
            files = proc.export_to_images(output_dir=out, format='PNG', quality=1.0,
                                          base_filename=f'r{i}')
            restore_imgs.append(files[0])
            proc.restore_template()
        proc2 = PPTXProcessor()
        reload_imgs = []
        for i, data in enumerate(rows):
            proc2.load_template(tpl)
            proc2.replace_placeholders(data)
            files = proc2.export_to_images(output_dir=out, format='PNG', quality=1.0,
                                           base_filename=f'l{i}')
            reload_imgs.append(files[0])
        for i in range(len(rows)):
            with Image.open(restore_imgs[i]) as a, Image.open(reload_imgs[i]) as b:
                same = a.size == b.size and ImageChops.difference(
                    a.convert('RGB'), b.convert('RGB')).getbbox() is None
            check(f'第{i+1}行图片逐像素一致（restore vs reload）', same)
        # 各行图片互不相同（数据确实逐行变化，没有缓存串味）
        distinct = True
        for i in range(len(rows)):
            for j in range(i + 1, len(rows)):
                with Image.open(restore_imgs[i]) as a, Image.open(restore_imgs[j]) as b:
                    if a.size == b.size and ImageChops.difference(
                            a.convert('RGB'), b.convert('RGB')).getbbox() is None:
                        distinct = False
        check('各行图片互不相同（无数据串行）', distinct)
    finally:
        shutil.rmtree(out, ignore_errors=True)


def main():
    tmp = tempfile.mkdtemp(prefix='restore_test_')
    try:
        synth = build_template(os.path.join(tmp, 'synth.pptx'))
        print('(1) 合成模板：restore 等价性')
        test_restore_equals_original(synth)
        srows, skey = make_rows(synth)
        print('(2) 合成模板：批量 PPT restore vs reload')
        test_batch_ppt_restore_vs_reload(synth, srows, skey)
        print('(3) 合成模板：批量图片 restore vs reload')
        test_batch_image_restore_vs_reload(synth, srows, skey)

        repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
        for name in ('10月承保贺报(3).pptx', '争霸赛承保贺报1.pptx'):
            real = next((os.path.join(d, name) for d in
                         (os.path.join(repo, 'samples', 'templates'), repo)
                         if os.path.exists(os.path.join(d, name))), None)
            if not real:
                print(f'(4) 真实模板 {name} 不存在，跳过')
                continue
            print(f'(4) 真实模板 {name}')
            test_restore_equals_original(real)
            rrows, rkey = make_rows(real)
            test_batch_ppt_restore_vs_reload(real, rrows, rkey)
            test_batch_image_restore_vs_reload(real, rrows, rkey)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print('=' * 60)
    print(f'结果: {PASS} 通过, {FAIL} 失败')
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
