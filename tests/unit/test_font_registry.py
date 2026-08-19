# -*- coding: utf-8 -*-
"""
字体注册表（font_registry）单元测试
验证：策略顺序（auto/system_first/bundled_first）、回退链、
每物理字体校准常数绑定、捆绑目录解析、is_resolvable/available_typefaces。

运行：.venv/Scripts/python.exe tests/unit/test_font_registry.py
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

try:
    from src.core.processors import font_registry as fr
except ImportError:
    from core.processors import font_registry as fr

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


def test_extra_font_dir():
    """本机扩展字体目录（静默增强）：精确命中/近名唯一/多候选不猜"""
    import glob
    import logging
    import shutil
    import tempfile
    fonts = glob.glob(os.path.join(os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))), 'resources', 'fonts',
        'SourceHanSansCN-*.otf'))
    if not fonts:
        check('捆绑字体存在（跳过条件）', False)
        return
    d = tempfile.mkdtemp(prefix='extrafonts_')
    try:
        for p in fonts:
            shutil.copy(p, d)
        os.environ['POSTERAGAIN_WPS_FONTS_DIR'] = d
        fr._extra_font_table = None
        fr.set_policy('system_first')  # Windows 默认：本机层（系统+扩展目录）优先
        spec = fr.resolve('思源黑体 CN', False)
        check(f'扩展目录精确命中 (source={spec.source})', spec.source == 'local')
        spec2 = fr.resolve('思源黑体 CX', False)  # 近名（距离1，唯一族）
        check(f'近名唯一候选命中 (source={spec2.source})', spec2.source == 'local')
        spec3 = fr.resolve('完全不存在的字体XYZ', False)
        check('无近似回退默认链', spec3.source != 'local')
    finally:
        os.environ.pop('POSTERAGAIN_WPS_FONTS_DIR', None)
        fr._extra_font_table = None
        fr.set_policy('auto')
        shutil.rmtree(d, ignore_errors=True)


def main():
    import sys as _sys
    on_windows = _sys.platform == 'win32'

    print("[1] 捆绑字体目录解析")
    bundled_dir = fr.bundled_font_dir()
    check("捆绑目录存在（resources/fonts）", bundled_dir is not None)
    if bundled_dir:
        for f in ('SourceHanSansCN-Regular.otf', 'SourceHanSansCN-Bold.otf',
                  'SourceHanSerifCN-Regular.otf', 'SourceHanSerifCN-Bold.otf'):
            check(f"捆绑文件存在: {f}", os.path.exists(os.path.join(bundled_dir, f)))

    print("[2] 默认策略（auto）")
    check("Windows 上 auto 生效为 system_first" if on_windows
          else "非 Windows 上 auto 生效为 bundled_first",
          fr.effective_policy() == ('system_first' if on_windows else 'bundled_first'))

    print("[3] system_first：Windows 行为零回归")
    fr.set_policy('system_first')
    spec = fr.resolve('微软雅黑', False)
    msyh = fr.SYSTEM_FONT_FILES[('微软雅黑', False)]
    if on_windows:
        check("微软雅黑 -> 系统 msyh.ttc", os.path.abspath(spec.path).lower()
              == os.path.abspath(msyh).lower())
    else:
        check("无系统表命中（非 Windows）-> 捆绑思源黑体",
              spec.source == 'bundled' and spec.family == 'sans')
    m = fr.metrics_for(spec.path)
    check("雅黑校准常数 1.25/(a=0,g=0.9)/0.245 且 calibrated（仿射模型 ≡ 历史 0.72×行高）",
          (m.line_height_ratio, m.baseline_a, m.baseline_g, m.strike_pos_em, m.calibrated)
          == (1.25, 0.0, 0.9, 0.245, True))

    print("[4] bundled_first：思源替换链")
    fr.set_policy('bundled_first')
    spec = fr.resolve('微软雅黑', False)
    check("微软雅黑 -> 捆绑思源黑体 Regular",
          spec.source == 'bundled' and spec.family == 'sans'
          and spec.path.endswith('SourceHanSansCN-Regular.otf'))
    spec = fr.resolve('微软雅黑', True)
    check("微软雅黑粗体 -> 捆绑思源黑体 Bold",
          spec.path.endswith('SourceHanSansCN-Bold.otf'))
    spec = fr.resolve('宋体', False)
    check("宋体 -> 捆绑思源宋体（serif 族）",
          spec.family == 'serif'
          and spec.path.endswith('SourceHanSerifCN-Regular.otf'))
    spec = fr.resolve('思源黑体', True)
    check("思源黑体直连捆绑 Bold 文件",
          spec.source == 'bundled' and spec.path.endswith('SourceHanSansCN-Bold.otf'))
    m = fr.metrics_for(spec.path)
    check("思源黑体 COM 校准常数 1.1942/0.1852/0.7803 且 calibrated",
          (m.line_height_ratio, m.baseline_a, m.baseline_g) == (1.1942, 0.1852, 0.7803)
          and m.calibrated is True)
    spec = fr.resolve('思源宋体', False)
    check("思源宋体共享垂直度量常数且 calibrated",
          spec.path.endswith('SourceHanSerifCN-Regular.otf')
          and fr.metrics_for(spec.path).calibrated is True)

    print("[5] 回退链")
    spec = fr.resolve('不存在的字体XYZ', False)
    check("未知字体回退不抛异常且路径可解析",
          spec.path and (os.path.exists(spec.path) or 'Fonts' in spec.path))
    check("未知字体 is_resolvable=False", fr.is_resolvable('不存在的字体XYZ') is False)
    check("微软雅黑 is_resolvable=True", fr.is_resolvable('微软雅黑') is True)
    check("思源宋体 is_resolvable=True", fr.is_resolvable('思源宋体') is True)

    print("[6] available_typefaces")
    names = fr.available_typefaces()
    check("包含捆绑思源黑体/宋体名", '思源黑体' in names and '思源宋体' in names)
    check("包含系统表命中名（微软雅黑）", '微软雅黑' in names or not on_windows)

    print("[7] 本机扩展字体目录（静默增强）")
    test_extra_font_dir()

    fr.set_policy('auto')  # 恢复默认
    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())

