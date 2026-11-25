# -*- coding: utf-8 -*-
"""
命令行入口（src.cli）单元测试

自包含：临时构造最小模板（ph_姓名 / ph_金额）与 CSV/Excel/JSON 数据，
不依赖 samples/ 私有业务文件，CI 可跑。覆盖：CLI 模式判定、预览、
PPT/图片生成、文件名前后缀、--rows 子集、CSV/Excel/JSON 加载、
错误退出码（2）、失败行退出码（1）。

运行：.venv/Scripts/python.exe tests/unit/test_cli.py
"""
import contextlib
import io
import os
import sys
import tempfile

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.abspath(os.path.join(_HERE, '..', '..'))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt

from src.cli import run, is_cli_mode

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


def call(argv):
    """在进程内调用 CLI，捕获 stdout/stderr 与退出码（含 argparse 的 SystemExit）"""
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            rc = run(argv)
    except SystemExit as e:
        rc = e.code if isinstance(e.code, int) else 1
    return rc, buf.getvalue()


def make_template(path):
    """ph_姓名 / ph_金额 能自动匹配；ph_备注 与任何列名都不同，需手工指定（--map/--preset）"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for i, (name, text) in enumerate((('ph_姓名', '姓名'), ('ph_金额', '金额'),
                                      ('ph_备注', '备注'))):
        box = slide.shapes.add_textbox(Inches(1), Inches(1 + i * 1.5), Inches(4), Inches(1))
        box.name = name
        run_ = box.text_frame.paragraphs[0].add_run()
        run_.text = text
        run_.font.size = Pt(28)
        run_.font.name = '微软雅黑'
    prs.save(path)


def make_data(tmp):
    df = pd.DataFrame({'姓名': ['张三', '李四', '坏数据'],
                       '金额': ['1,000', '2,000', '3,000'],
                       '备注信息': ['甲', '乙', '丙']})
    csv = os.path.join(tmp, 'data.csv')
    xlsx = os.path.join(tmp, 'data.xlsx')
    jsonf = os.path.join(tmp, 'data.json')
    df.to_csv(csv, index=False, encoding='utf-8-sig')
    df.to_excel(xlsx, index=False)
    df.to_json(jsonf, orient='records', force_ascii=False)
    return csv, xlsx, jsonf


def make_config(tmp):
    cfg = os.path.join(tmp, 'cfg.yaml')
    with open(cfg, 'w', encoding='utf-8') as f:
        f.write('processing:\n  image_export_engine: python\n')
    return cfg


def main():
    tmp = tempfile.mkdtemp(prefix='cli_test_')
    try:
        tpl = os.path.join(tmp, 'tpl.pptx')
        make_template(tpl)
        csv, xlsx, jsonf = make_data(tmp)
        cfg = make_config(tmp)
        base = ['-t', tpl, '--config', cfg]

        print('[1] CLI 模式判定')
        check('带 --template 判为 CLI', is_cli_mode(['--template', 'x.pptx']))
        check('带 -t / --preview / --help 判为 CLI',
              is_cli_mode(['-t', 'x']) and is_cli_mode(['--preview'])
              and is_cli_mode(['--help']))
        check('无参数判为 GUI', not is_cli_mode([]))

        print('[2] 预览（只匹配不生成）')
        out = os.path.join(tmp, 'prev')
        rc, text = call(base + ['-d', csv, '--preview'])
        check('预览退出码 0', rc == 0)
        check('预览打印匹配关系', 'ph_姓名' in text and '姓名' in text)
        check('预览不创建输出目录', not os.path.exists(out))

        print('[3] 生成 PPT（CSV）')
        out = os.path.join(tmp, 'ppt')
        rc, text = call(base + ['-d', csv, '-o', out, '-q'])
        check('PPT 退出码 0', rc == 0)
        check('摘要含成功 3/3', '成功 3/3' in text)
        check('生成 3 个 pptx', len([f for f in os.listdir(out) if f.endswith('.pptx')]) == 3)

        print('[4] 文件名前后缀')
        out = os.path.join(tmp, 'ps')
        rc, _ = call(base + ['-d', csv, '-o', out, '-q', '--prefix', '贺报_', '--suffix', '_v1'])
        files = sorted(os.listdir(out))
        check('前缀生效', rc == 0 and all(f.startswith('贺报_') for f in files))
        check('后缀生效（扩展名前）', all(f.endswith('_v1.pptx') for f in files))

        print('[5] 生成图片')
        out = os.path.join(tmp, 'img')
        rc, text = call(base + ['-d', csv, '-o', out, '--images', '-q'])
        pngs = [f for f in os.listdir(out) if f.endswith('.png')]
        check('图片退出码 0', rc == 0)
        check('生成 3 张 png', len(pngs) == 3)

        print('[6] Excel / JSON 数据')
        for label, data in (('Excel', xlsx), ('JSON', jsonf)):
            out = os.path.join(tmp, 'fmt_' + label)
            rc, text = call(base + ['-d', data, '-o', out, '-q'])
            check(f'{label} 加载并生成（退出码 0，3 个文件）',
                  rc == 0 and len(os.listdir(out)) == 3)

        print('[7] --rows 子集')
        out = os.path.join(tmp, 'rows')
        rc, text = call(base + ['-d', csv, '-o', out, '-q', '--rows', '0,2'])
        check('子集退出码 0 且成功 2/2', rc == 0 and '成功 2/2' in text)
        check('只生成 2 个文件', len(os.listdir(out)) == 2)

        print('[8] 错误退出码（2）')
        rc, _ = call(['-t', os.path.join(tmp, '不存在.pptx'), '-d', csv,
                      '-o', os.path.join(tmp, 'e1'), '--config', cfg])
        check('模板不存在 → 2', rc == 2)
        rc, _ = call(base + ['-d', os.path.join(tmp, '不存在.csv'),
                      '-o', os.path.join(tmp, 'e2')])
        check('数据不存在 → 2', rc == 2)
        rc, _ = call(base + ['-d', csv])
        check('生成但缺 --output → 2', rc == 2)
        rc, _ = call(base + ['-d', csv, '-o', os.path.join(tmp, 'e3'),
                      '--rows', 'abc'])
        check('--rows 非法 → 2', rc == 2)

        print('[9] 失败行退出码（1）')
        from src.core.processors import pptx_processor as pp
        orig = pp.PPTXProcessor.save_presentation

        def poison(self, output_path, *a, **k):
            if '坏数据' in os.path.basename(str(output_path)):
                raise RuntimeError('模拟生成失败')
            return orig(self, output_path, *a, **k)

        pp.PPTXProcessor.save_presentation = poison
        try:
            out = os.path.join(tmp, 'fail')
            rc, text = call(base + ['-d', csv, '-o', out, '-q'])
        finally:
            pp.PPTXProcessor.save_presentation = orig
        check('有失败行 → 退出码 1', rc == 1)
        check('摘要报告失败行', '失败 1 行' in text and '坏数据' in text)
        check('成功行仍生成（2 个）',
              len([f for f in os.listdir(out) if f.endswith('.pptx')]) == 2)
        check('提示可用 --rows 重跑', '--rows' in text)

        def texts_of(out_dir, name='1_张三.pptx'):
            prs = Presentation(os.path.join(out_dir, name))
            return [s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame]

        print('[10] 手工匹配 --map（自动匹配对不上的占位符）')
        out = os.path.join(tmp, 'nomap')
        rc, text = call(base + ['-d', csv, '-o', out, '-q'])
        check('未手工指定时：警告未匹配、占位符保留模板原文、退出码 0',
              rc == 0 and 'ph_备注' in text and '保留模板原文' in text
              and texts_of(out) == ['张三', '1,000', '备注'])
        out = os.path.join(tmp, 'map')
        rc, text = call(base + ['-d', csv, '-o', out, '-q', '--map', 'ph_备注=备注信息'])
        check('--map 后三处全部替换、无警告',
              rc == 0 and '保留模板原文' not in text
              and texts_of(out) == ['张三', '1,000', '甲'])
        rc, text = call(base + ['-d', csv, '-o', os.path.join(tmp, 'mapbad'), '-q',
                                '--map', 'ph_备注=没有这一列'])
        check('--map 指向不存在的列：提示并忽略，不影响其余（退出码 0）',
              rc == 0 and '已忽略' in text and '没有这一列' in text)

        print('[11] 文字前后缀 --add-prefix / --add-suffix')
        out = os.path.join(tmp, 'add')
        rc, _ = call(base + ['-d', csv, '-o', out, '-q',
                             '--add-prefix', 'ph_姓名=客户', '--add-suffix', 'ph_金额=元'])
        check('前缀、后缀都生效', rc == 0 and texts_of(out) == ['客户张三', '1,000元', '备注'])

        print('[12] 预设文件：界面「占位符配置」格式（matching_rules 列表 = 完整替换）')
        full = os.path.join(tmp, 'full.pptcfg')
        with open(full, 'w', encoding='utf-8') as f:
            import json
            json.dump({'version': '1.0.0', 'matching_rules': [
                {'placeholder': 'ph_姓名', 'column': '姓名',
                 'text_addition': {'prefix': '', 'suffix': '先生'}},
                {'placeholder': 'ph_备注', 'column': '备注信息'},
                {'placeholder': 'ph_改名了', 'column': '这一列不存在'},
            ]}, f, ensure_ascii=False)
        out = os.path.join(tmp, 'preset_full')
        rc, text = call(base + ['-d', csv, '-o', out, '-q', '--preset', full])
        check('预设规则与文字增加生效、列缺失的规则被提示',
              rc == 0 and '这一列不存在' in text
              and texts_of(out) == ['张三先生', '金额', '甲'])
        check('列表格式视为完整规则集：预设里没有的 ph_金额 不再自动匹配（保留原文）',
              texts_of(out)[1] == '金额')

        print('[13] 预设文件：界面「简化配置」格式（gui_settings = 叠加在自动匹配上）')
        simple = os.path.join(tmp, 'simple.pptcfg')
        with open(simple, 'w', encoding='utf-8') as f:
            import json
            json.dump({'version': '1.0.0', 'gui_settings': {
                'text_additions': {'ph_金额': {'prefix': '', 'suffix': '元整'}},
                'chengbao_mappings': {'ph_备注': '备注信息'},
            }}, f, ensure_ascii=False)
        out = os.path.join(tmp, 'preset_simple')
        rc, _ = call(base + ['-d', csv, '-o', out, '-q', '--preset', simple])
        check('自动匹配保留 + 趸期映射叠加 + 文字增加生效',
              rc == 0 and texts_of(out) == ['张三', '1,000元整', '甲'])

        print('[14] 预设 / 规则参数错误（2）')
        rc, _ = call(base + ['-d', csv, '-o', os.path.join(tmp, 'e4'),
                             '--preset', os.path.join(tmp, '没有这个.pptcfg')])
        check('预设文件不存在 → 2', rc == 2)
        bad = os.path.join(tmp, 'bad.pptcfg')
        with open(bad, 'w', encoding='utf-8') as f:
            f.write('{这不是JSON')
        rc, text = call(base + ['-d', csv, '-o', os.path.join(tmp, 'e5'), '--preset', bad])
        check('预设不是合法 JSON → 2 并给出原因', rc == 2 and '预设无效' in text)
        rc, _ = call(base + ['-d', csv, '-o', os.path.join(tmp, 'e6'), '--map', '少了等号'])
        check('--map 格式错误 → 2', rc == 2)
        rc, _ = call(base + ['-d', csv, '-o', os.path.join(tmp, 'e7'),
                             '--add-suffix', '也少了等号'])
        check('--add-suffix 格式错误 → 2', rc == 2)
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)

    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
