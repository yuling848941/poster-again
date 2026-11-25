# -*- coding: utf-8 -*-
"""
系统已安装字体（按名字匹配）单元测试

验证：name 表读取（与 fontTools 互校）、索引构建、名字归一化、字重/宽度词拆分、
思源 <-> Noto CJK / SC <-> CN 等价、可变字体实例选择、TTC 字面序号、
查找顺序（系统路径表 -> 扩展目录 -> 系统已安装 -> 捆绑）、索引缓存、
预检不误报、渲染器字体加载（字面序号/可变实例）与同一文本框内按字体分组绘制。

运行：.venv/Scripts/python.exe libs/render_core/tests/test_system_fonts.py

不依赖本机装了哪些字体：核心用例用临时目录里的字体副本和手工构造的索引；
标注「本机」的断言只在对应系统字体存在时才执行。
"""
import os
import shutil
import sys
import tempfile
import time

_HERE = os.path.dirname(os.path.abspath(__file__))
# 本库（render_core）与它依赖的 fontfx 都在仓库里；已 pip 安装时这几行是冗余的
for _p in (os.path.join(_HERE, '..', 'src'),
           os.path.join(_HERE, '..', '..', 'fontfx', 'src')):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from render_core import font_registry as fr

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


class _FakeSystem:
    """临时把「系统字体文件列表」换成指定文件（用例结束自动还原并清缓存）"""

    def __init__(self, files):
        self.files = list(files)

    def __enter__(self):
        self._real = fr._system_font_files
        fr._system_font_files = lambda: list(self.files)
        fr.invalidate_system_fonts()
        return self

    def __exit__(self, *exc):
        fr._system_font_files = self._real
        fr.invalidate_system_fonts()
        return False


def _bundled(*names):
    d = fr.bundled_font_dir()
    return [os.path.join(d, n) for n in names] if d else []


def _system_font(name):
    """本机系统字体路径（不存在返回 None）"""
    for d in fr._system_font_dirs():
        p = os.path.join(d, name)
        if os.path.isfile(p):
            return p
    return None


def test_read_face_info():
    print("[1] name 表读取（标准库 struct）")
    serif_bold, serif_reg, sans_reg = _bundled('SourceHanSerifCN-Bold.otf',
                                               'SourceHanSerifCN-Regular.otf',
                                               'SourceHanSansCN-Regular.otf')
    if not serif_bold:
        check('捆绑字体存在（跳过条件）', False)
        return
    info = fr._read_face_info(serif_bold)
    check('Bold 字面：家族名含中英文两种写法',
          info['names'][1] == {'Source Han Serif CN', '思源宋体 CN'})
    check('Bold 字面：全名 / PostScript 名',
          'Source Han Serif CN Bold' in info['names'][4]
          and info['names'][6] == {'SourceHanSerifCN-Bold'})
    check('Bold 字面：粗体位 + 字重 700 + 非斜体 + 非可变',
          info['bold'] and info['weight'] == 700
          and not info['italic'] and not info['variable'])
    info = fr._read_face_info(serif_reg)
    check('Regular 字面：非粗体 + 字重 400', not info['bold'] and info['weight'] == 400)
    check('思源黑体 CN 家族名',
          fr._read_face_info(sans_reg)['names'][1] == {'Source Han Sans CN', '思源黑体 CN'})
    check('普通字体字面数为 1', fr._face_count(serif_bold) == 1)

    tmp = tempfile.mkdtemp(prefix='badfont_')
    try:
        junk = os.path.join(tmp, 'junk.ttf')
        with open(junk, 'wb') as f:
            f.write(b'this is not a font file' * 16)
        check('不是字体的文件 -> None / 0',
              fr._read_face_info(junk) is None and fr._face_count(junk) == 0)
        check('不存在的文件 -> None', fr._read_face_info(os.path.join(tmp, 'nope.ttf')) is None)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        print("  (未装 fontTools，跳过互校)")
        return
    samples = [serif_bold, serif_reg, sans_reg]
    for name in ('calibri.ttf', 'calibril.ttf', 'arial.ttf', 'msyh.ttc',
                 'Deng.ttf', 'bahnschrift.ttf'):
        p = _system_font(name)
        if p:
            samples.append(p)
    n_bad = 0
    for p in samples:
        for idx in range(fr._face_count(p)):
            info = fr._read_face_info(p, idx)
            tt = TTFont(p, fontNumber=idx, lazy=True)
            try:
                for nid in (1, 2, 4, 6, 16, 17):
                    expected = {' '.join(r.toUnicode().split()) for r in tt['name'].names
                                if r.nameID == nid and r.platformID in (0, 3)}
                    ours = info['names'].get(nid, set())
                    # Unicode 平台的记录必须都读到；读到的非 ASCII 串必须都是真名字
                    if (not expected <= ours
                            or any(not s.isascii() and s not in expected for s in ours)):
                        n_bad += 1
                        print(f"     nameID {nid} 不一致 {os.path.basename(p)}: "
                              f"缺 {sorted(expected - ours)} / 多 {sorted(ours - expected)}")
                os2, head = tt['OS/2'], tt['head']
                if (info['bold'] != (bool(os2.fsSelection & 0x20) or bool(head.macStyle & 1))
                        or info['italic'] != (bool(os2.fsSelection & 0x01)
                                              or bool(head.macStyle & 2))
                        or info['variable'] != ('fvar' in tt)):
                    n_bad += 1
                    print(f"     粗体/斜体/可变标记不一致 {os.path.basename(p)}[{idx}]")
            finally:
                tt.close()
    check(f'与 fontTools 互校一致（{len(samples)} 个字体文件的各个字面）', n_bad == 0)


def test_table_and_lookup():
    print("[2] 索引构建与按名匹配（临时目录里的字体副本，文件名与字体名无关）")
    bundled = _bundled('SourceHanSerifCN-Bold.otf', 'SourceHanSerifCN-Regular.otf',
                       'SourceHanSansCN-Regular.otf')
    if not bundled:
        check('捆绑字体存在（跳过条件）', False)
        return
    tmp = tempfile.mkdtemp(prefix='sysfonts_')
    try:
        # 故意只装黑体的常规体：粗体请求应退到同一族的常规体而不是报错
        copies = {'f001.bin': 'SourceHanSerifCN-Bold.otf',      # 宋体 Bold
                  'f002.bin': 'SourceHanSerifCN-Regular.otf',   # 宋体 Regular
                  'f003.bin': 'SourceHanSansCN-Regular.otf'}    # 黑体 Regular
        for dst, src in copies.items():
            shutil.copy(os.path.join(fr.bundled_font_dir(), src), os.path.join(tmp, dst))
        junk = os.path.join(tmp, 'broken.ttf')
        with open(junk, 'wb') as f:
            f.write(b'\x00\x01\x00\x00broken' * 8)
        with _FakeSystem([os.path.join(tmp, n) for n in copies] + [junk]):
            lookup = fr._lookup_system_by_name
            cases = [
                # 家族名 + 粗体属性
                ('Source Han Serif CN', False, 'f002.bin'),
                ('Source Han Serif CN', True, 'f001.bin'),
                ('思源宋体 CN', True, 'f001.bin'),
                # PostScript 写法（大小写/连字符不敏感）
                ('sourcehanserifcn-bold', False, 'f001.bin'),
                ('SourceHanSerifCN-Bold', False, 'f001.bin'),
                # 简称与等价写法：思源 <-> Noto CJK、SC <-> CN
                ('思源宋体', False, 'f002.bin'),
                ('思源宋体 Bold', False, 'f001.bin'),
                ('source han serif sc', False, 'f002.bin'),
                ('SourceHanSerifSC-Bold', False, 'f001.bin'),
                ('SourceHanSerifSC-SemiBold', False, 'f001.bin'),     # 没装 SemiBold -> Bold
                ('SourceHanSerifSC-Medium', False, 'f002.bin'),       # 没装 Medium -> Regular
                ('Noto Serif CJK SC', False, 'f002.bin'),
                # 只装了常规体的家族：粗体请求退到常规体
                ('思源黑体', False, 'f003.bin'),
                ('思源黑体', True, 'f003.bin'),
                # 找不到
                ('完全不存在的字体XYZ', False, None),
                ('SourceHanSerifTC-Bold', False, None),               # 繁体写法不做等价
            ]
            for typeface, bold, want in cases:
                spec = lookup(typeface, bold)
                got = os.path.basename(spec.path) if spec else None
                check(f"{typeface!r} 粗体={bold} -> {want}", got == want)
            check('None / 空串不命中',
                  lookup(None, False) is None and lookup('', False) is None
                  and lookup('   ', True) is None)
            names = fr.system_font_names()
            check('名字集合含中英文家族名（小写、保留空格）',
                  {'source han serif cn', '思源宋体 cn', '思源黑体 cn',
                   'sourcehanserifcn-bold'} <= names)
            check('损坏文件不进索引',
                  all('broken' not in h.face.path
                      for hits in fr.system_font_table().values() for h in hits))
            spec = lookup('思源宋体', True)
            check('命中结果 source=system、字面序号 0、无实例',
                  spec.source == 'system' and spec.index == 0 and spec.variation == '')
            check('系统里装的思源沿用思源校准常数（宋体/黑体各归各）',
                  fr.metrics_for(lookup('思源宋体', False).path) is fr.METRICS_SOURCE_HAN_SERIF
                  and fr.metrics_for(lookup('思源黑体', False).path) is fr.METRICS_SOURCE_HAN_SANS)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# Bahnschrift（Windows 自带可变字体）的 15 个具名实例，按它真实的命名规律手工构造
_BAHNSCHRIFT = ('Light', 'SemiLight', 'Regular', 'SemiBold', 'Bold',
                'Light SemiCondensed', 'SemiLight SemiCondensed', 'SemiCondensed',
                'SemiBold SemiCondensed', 'Bold SemiCondensed', 'Light Condensed',
                'SemiLight Condensed', 'Condensed', 'SemiBold Condensed', 'Bold Condensed')


def _face(family, sub='Regular', full=None, ps=None, bold=False, italic=False,
          weight=None, typographic=None, variable=False, variations=()):
    """手工构造一个字面的 _read_face_info 结果"""
    names = {1: {family}, 2: {sub}}
    if full:
        names[4] = {full}
    if ps:
        names[6] = {ps}
    if typographic:
        names[16], names[17] = {typographic[0]}, {typographic[1]}
    return {'names': names, 'bold': bold, 'italic': italic,
            'weight': weight or (700 if bold else 400), 'variable': variable}, tuple(variations)


def _synthetic_table():
    """一个不依赖任何字体文件的索引：普通家族 4 个字面 + Light 变体、
    TTC 两个字面、可变字体、Noto CJK（用来测思源 <-> Noto 等价）"""
    items = [
        ('sans-r.ttf', 0, *_face('Test Sans', full='Test Sans', ps='TestSans-Regular')),
        ('sans-b.ttf', 0, *_face('Test Sans', 'Bold', 'Test Sans Bold', 'TestSans-Bold', bold=True)),
        ('sans-i.ttf', 0, *_face('Test Sans', 'Italic', 'Test Sans Italic', 'TestSans-Italic', italic=True)),
        ('sans-l.ttf', 0, *_face('Test Sans Light', full='Test Sans Light', ps='TestSans-Light',
                                  weight=300, typographic=('Test Sans', 'Light'))),
        ('coll.ttc', 0, *_face('Coll A', full='Coll A', ps='CollA')),
        ('coll.ttc', 1, *_face('Coll B', full='Coll B', ps='CollB')),
        ('var.ttf', 0, *_face('Var Sans', full='Var Sans', ps='VarSans-Regular',
                               variable=True, variations=_BAHNSCHRIFT)),
        ('noto.otf', 0, *_face('Noto Sans CJK SC', full='Noto Sans CJK SC', ps='NotoSansCJKsc-Regular')),
        ('noto-b.otf', 0, *_face('Noto Sans CJK SC', 'Bold', 'Noto Sans CJK SC Bold',
                                  'NotoSansCJKsc-Bold', bold=True)),
    ]
    return fr._index_faces(items)[0]


def test_matching_rules():
    print("[3] 匹配规则（手工索引：归一化、字重/宽度词、斜体、TTC、可变字体、等价写法）")
    table = _synthetic_table()

    def match(typeface, bold=False):
        return fr._match_system(table, typeface, bold)

    cases = [
        ('Test Sans', False, 'sans-r.ttf', 0, ''),
        ('Test Sans', True, 'sans-b.ttf', 0, ''),
        ('test sans', True, 'sans-b.ttf', 0, ''),          # 大小写不敏感
        ('TESTSANS', False, 'sans-r.ttf', 0, ''),          # 空格不敏感
        ('TestSans-Bold', False, 'sans-b.ttf', 0, ''),     # PostScript 写法
        ('Test Sans Bold', False, 'sans-b.ttf', 0, ''),    # 全名写法
        ('Test Sans Light', False, 'sans-l.ttf', 0, ''),
        ('Test Sans Light', True, 'sans-l.ttf', 0, ''),    # 家族里没有粗体 -> 用现有的
        ('TestSans-Light', False, 'sans-l.ttf', 0, ''),
        ('Test Sans Italic', False, 'sans-i.ttf', 0, ''),  # 点名要斜体才用斜体字面
        ('Test Sans Condensed', False, 'sans-r.ttf', 0, ''),  # 没有窄体 -> 常规体
        ('Coll B', False, 'coll.ttc', 1, ''),              # TTC 第二个字面
        ('CollA', False, 'coll.ttc', 0, ''),
        ('Var Sans', False, 'var.ttf', 0, ''),             # 可变字体默认实例
        ('Var Sans', True, 'var.ttf', 0, 'Bold'),
        ('VarSans-Bold', False, 'var.ttf', 0, 'Bold'),
        ('Var Sans Light', True, 'var.ttf', 0, 'SemiLight'),            # 加粗升一档
        ('Var Sans SemiBold SemiCondensed', False, 'var.ttf', 0, 'SemiBold SemiCondensed'),
        ('Var Sans SemiBold SemiCondensed', True, 'var.ttf', 0, 'Bold SemiCondensed'),
        ('Var Sans SemiLight SemiCondensed', True, 'var.ttf', 0, 'SemiCondensed'),
        ('Var Sans SemiCondensed', False, 'var.ttf', 0, 'SemiCondensed'),
        ('Var Sans Condensed', True, 'var.ttf', 0, 'Bold Condensed'),
        ('思源黑体', False, 'noto.otf', 0, ''),             # 装的是 Noto：同一套设计
        ('思源黑体', True, 'noto-b.otf', 0, ''),
        ('Source Han Sans SC', False, 'noto.otf', 0, ''),
        ('SourceHanSansTC-Regular', False, None, 0, ''),    # 繁体写法不做等价
        ('完全不存在', False, None, 0, ''),
        (None, False, None, 0, ''),
        ('', False, None, 0, ''),
    ]
    for typeface, bold, path, index, variation in cases:
        got = match(typeface, bold)
        want = None if path is None else (path, index, variation)
        check(f"{typeface!r} 粗体={bold} -> "
              f"{want if want else '未命中'}", got == want)


def test_resolve_order():
    print("[4] 查找顺序（resolve 层）")
    bundled = _bundled('SourceHanSerifCN-Regular.otf')
    if not bundled:
        check('捆绑字体存在（跳过条件）', False)
        return
    tmp = tempfile.mkdtemp(prefix='sysorder_')
    old_wps = os.environ.get('POSTERAGAIN_WPS_FONTS_DIR')
    copy = os.path.join(tmp, 'system_copy.otf')
    shutil.copy(bundled[0], copy)
    try:
        # 扩展字体目录置空，保证测的是「系统已安装」这一层
        os.environ['POSTERAGAIN_WPS_FONTS_DIR'] = tmp
        fr._extra_font_table = None
        with _FakeSystem([copy]):
            fr.set_policy('system_first')
            spec = fr.resolve('思源宋体', False)
            check('system_first：系统已安装的思源优先于捆绑',
                  spec.source == 'system' and os.path.normcase(spec.path) == os.path.normcase(copy))
            check('is_resolvable 认 PostScript 写法', fr.is_resolvable('SourceHanSerifSC-Bold'))
            check('available_typefaces 含系统字体名',
                  'source han serif cn' in fr.available_typefaces())
            fr.set_policy('bundled_first')
            spec = fr.resolve('思源宋体', False)
            check('bundled_first：捆绑优先于系统已安装', spec.source == 'bundled')
        with _FakeSystem([]):
            fr.set_policy('system_first')
            spec = fr.resolve('思源宋体', False)
            check('系统没装 -> 回退捆绑', spec.source == 'bundled')
            check('系统没装时 is_resolvable 仍为 True（捆绑兜底）', fr.is_resolvable('思源宋体'))
        if sys.platform == 'win32':
            msyh = fr.SYSTEM_FONT_FILES[('微软雅黑', False)]
            if os.path.exists(msyh):
                fr.set_policy('system_first')
                spec = fr.resolve('微软雅黑', False)
                check('系统路径表仍最优先（微软雅黑 -> msyh.ttc）',
                      os.path.normcase(spec.path) == os.path.normcase(msyh))
    finally:
        if old_wps is None:
            os.environ.pop('POSTERAGAIN_WPS_FONTS_DIR', None)
        else:
            os.environ['POSTERAGAIN_WPS_FONTS_DIR'] = old_wps
        fr._extra_font_table = None
        fr.set_policy('auto')
        shutil.rmtree(tmp, ignore_errors=True)


def test_cache():
    print("[5] 索引缓存")
    with _FakeSystem([]):
        first = fr.system_font_table()
        check('窗口期内复用同一张表', fr.system_font_table() is first)
        fr.invalidate_system_fonts()
        check('invalidate 后重建', fr.system_font_table() is not first)


def test_precheck_consistency():
    print("[6] 预检与渲染一致（能解析到文件的字体不报缺失）")
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from render_core.render_precheck import RenderPrecheck

    bundled = _bundled('SourceHanSerifCN-Bold.otf')
    if not bundled:
        check('捆绑字体存在（跳过条件）', False)
        return
    tmp = tempfile.mkdtemp(prefix='precheck_')
    try:
        copy = os.path.join(tmp, 'bold.otf')
        shutil.copy(bundled[0], copy)
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for i, name in enumerate(('SourceHanSerifSC-Bold', '不存在的字体XYZ', 'Calibri')):
            tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5 + i), Inches(6), Inches(0.8))
            run = tb.text_frame.paragraphs[0].add_run()
            run.text = '探针'
            run.font.size = Pt(28)
            run.font.name = name
        pptx = os.path.join(tmp, 'probe.pptx')
        prs.save(pptx)
        with _FakeSystem([copy]):

            def missing(**kw):
                return [i.message for i in RenderPrecheck(**kw).check_pptx(pptx)
                        if i.code == 'FONT_MISSING']

            msgs = missing(installed_fonts={'微软雅黑'}, font_resolver=fr.is_resolvable)
            check('带 resolver：能按 PostScript 名解析的字体不报缺失',
                  not any('SourceHanSerifSC-Bold' in m for m in msgs))
            check('带 resolver：真没有的字体照常报缺失',
                  any('不存在的字体XYZ' in m for m in msgs) and any('Calibri' in m for m in msgs))
            msgs = missing(installed_fonts={'微软雅黑'})
            check('不带 resolver：只认名字集合（会误报 PostScript 写法）',
                  any('SourceHanSerifSC-Bold' in m for m in msgs))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_renderer_font_loading():
    print("[7] 渲染器字体加载（本机字体存在时才断言）")
    from PIL import ImageFont
    from render_core.pillow_image_exporter import PillowImageExporter, _Style

    def style(path, **kw):
        return _Style(40, False, False, 0.0, None, None, None, None, False, None, path, **kw)

    a, b = style('a.ttf'), style('b.ttf')
    check('样式标识包含字体文件', a.key() != b.key())
    check('样式标识包含字面序号', style('a.ttf').key() != style('a.ttf', font_index=1).key())
    check('样式标识包含可变实例',
          style('a.ttf').key() != style('a.ttf', font_variation='Bold').key())

    msyh = _system_font('msyh.ttc')
    if msyh:
        exporter = PillowImageExporter()
        check('TTC 字面序号 0 = Microsoft YaHei',
              exporter._font(style(msyh)).getname()[0] == 'Microsoft YaHei')
        check('TTC 字面序号 1 = Microsoft YaHei UI',
              exporter._font(style(msyh, font_index=1)).getname()[0] == 'Microsoft YaHei UI')
        check('不同序号分别缓存', exporter._font(style(msyh)) is not exporter._font(style(msyh, font_index=1)))
    else:
        print("  (本机没有 msyh.ttc，跳过 TTC 断言)")

    bahnschrift = _system_font('bahnschrift.ttf')
    if bahnschrift:
        exporter = PillowImageExporter()
        default = exporter._font(style(bahnschrift))
        semi = exporter._font(style(bahnschrift, font_variation='SemiBold SemiCondensed'))
        text = 'Poster 2026'
        check('可变实例生效（SemiBold SemiCondensed 比默认实例窄）',
              semi.getlength(text) < default.getlength(text))
        ref = ImageFont.truetype(bahnschrift, 40 * exporter._dpi / 72.0)
        ref.set_variation_by_name('SemiBold SemiCondensed')
        check('与直接用 Pillow 设实例的结果一致', abs(semi.getlength(text) - ref.getlength(text)) < 0.01)
        bogus = exporter._font(style(bahnschrift, font_variation='NoSuchInstance'))
        check('不存在的实例名 -> 用默认实例，不抛异常', bogus.getlength(text) == default.getlength(text))
    else:
        print("  (本机没有 bahnschrift.ttf，跳过可变字体断言)")


def test_mixed_fonts_drawn_separately():
    print("[8] 同一文本框内不同字体各画各的")
    from PIL import Image, ImageChops
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from render_core import pillow_image_exporter as pie
    from render_core import render

    if not _bundled('SourceHanSerifCN-Regular.otf', 'SourceHanSansCN-Regular.otf'):
        check('捆绑字体存在（跳过条件）', False)
        return
    tmp = tempfile.mkdtemp(prefix='mixedfonts_')

    def deck(name, font_a, font_b):
        path = os.path.join(tmp, name)
        prs = Presentation()
        prs.slide_width = Inches(7.5)
        prs.slide_height = Inches(10.0)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(6.5), Inches(2.0))
        para = tb.text_frame.paragraphs[0]
        for text, font_name in (('永东', font_a), ('国风', font_b)):
            # 两段文字除字体外属性完全相同（同段、同字号、同颜色、都不加粗）
            run = para.add_run()
            run.text = text
            run.font.size = Pt(96)
            run.font.name = font_name
        prs.save(path)
        return path

    def pixels(pptx, out):
        return Image.open(render(pptx, os.path.join(tmp, out), 'PNG', 1.0).images[0]).convert('RGB')

    def differs(img1, img2):
        return (img1.size != img2.size
                or ImageChops.difference(img1, img2).getbbox() is not None)

    fr.set_policy('bundled_first')       # 两个名字都走捆绑字体，任何平台都可用
    try:
        img_serif = pixels(deck('serif.pptx', '思源宋体', '思源宋体'), 'a')
        img_sans = pixels(deck('sans.pptx', '思源黑体', '思源黑体'), 'b')
        mixed = deck('mixed.pptx', '思源宋体', '思源黑体')
        img_mixed = pixels(mixed, 'c')
        check('对照组：纯宋体与纯黑体的图不同', differs(img_serif, img_sans))
        check('混合文本既不同于纯宋体、也不同于纯黑体',
              differs(img_mixed, img_serif) and differs(img_mixed, img_sans))
        # 用例有效性：把字体身份从样式标识里拿掉（即旧行为），后一段就会画成前一段的字体
        real_key = pie._Style.key
        pie._Style.key = lambda self: real_key(self)[:-3]
        try:
            img_old = pixels(mixed, 'old')
        finally:
            pie._Style.key = real_key
        check('用例有效性：去掉字体身份后混合文本退化成纯宋体（本用例确实能发现该问题）',
              not differs(img_old, img_serif))
    finally:
        fr.set_policy('auto')
        shutil.rmtree(tmp, ignore_errors=True)


def test_local_system_fonts():
    print("[9] 本机系统字体（装了才断言）")
    if sys.platform != 'win32':
        print("  (非 Windows，系统字体索引为空)")
        fr.invalidate_system_fonts()
        check('非 Windows 索引为空且不报错', fr.system_font_table() == {})
        return
    fr.invalidate_system_fonts()
    t0 = time.perf_counter()
    table = fr.system_font_table()
    dt = time.perf_counter() - t0
    print(f"  本机索引 {len(table)} 个名字，构建 {dt * 1000:.0f}ms")
    check('本机索引非空', len(table) > 50)
    check('建索引耗时在 5 秒内', dt < 5.0)
    fr.set_policy('system_first')
    want = [('Calibri', True, 'calibrib.ttf'), ('Arial', True, 'arialbd.ttf'),
            ('等线', False, 'Deng.ttf'), ('等线 Light', False, 'dengl.ttf'),
            ('Segoe UI', True, 'segoeuib.ttf'), ('Microsoft YaHei UI', False, 'msyh.ttc')]
    try:
        for typeface, bold, filename in want:
            if not _system_font(filename):
                print(f"  (本机未装 {filename}，跳过 {typeface!r})")
                continue
            spec = fr.resolve(typeface, bold)
            check(f"{typeface!r} 粗体={bold} -> {filename}",
                  spec.source == 'system'
                  and os.path.basename(spec.path).lower() == filename.lower())
        if _system_font('msyh.ttc'):
            spec = fr.resolve('Microsoft YaHei UI', False)
            check('Microsoft YaHei UI 是 msyh.ttc 的第 2 个字面（序号 1）', spec.index == 1)
        if _system_font('bahnschrift.ttf'):
            spec = fr.resolve('Bahnschrift SemiBold SemiCondensed', False)
            check('Bahnschrift SemiBold SemiCondensed -> 可变实例',
                  spec.variation == 'SemiBold SemiCondensed')
    finally:
        fr.set_policy('auto')


def main():
    try:
        test_read_face_info()
        test_table_and_lookup()
        test_matching_rules()
        test_resolve_order()
        test_cache()
        test_precheck_consistency()
        test_renderer_font_loading()
        test_mixed_fonts_drawn_separately()
        test_local_system_fonts()
    finally:
        fr.set_policy('auto')
    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
