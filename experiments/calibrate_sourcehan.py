# -*- coding: utf-8 -*-
"""
思源字体排版常数校准（font_registry.METRICS_SOURCE_HAN_SANS/SERIF 的数据来源）
================================================================
用法：
    .venv/Scripts/python.exe experiments/calibrate_sourcehan.py sans    # 校准思源黑体
    .venv/Scripts/python.exe experiments/calibrate_sourcehan.py serif   # 校准思源宋体

前置条件（仅本机可做，PowerPoint COM 必须可用）：
1. resources/fonts/ 下思源字体已下载（渲染端用捆绑文件）
2. 同一批字体已安装到当前用户（HKCU Fonts，PowerPoint 渲染参考图用）
   —— 两边是同一文件，墨迹差异只反映排版常数差异

模型（经 2026-08-18 实测发现，PowerPoint 首行基线对 lnSpc 非纯比例）：
- 行距：line_h = pt2px(sz) × ratio × lnSpc            （比例模型，pitch 实测吻合）
- 首行基线：baseline = (a + c × lnSpc) × line_h 的基线分数（仿射模型，最小二乘拟合）
  msyh 的 a=0.72/c=0 即历史常数（比例模型是仿射的特例）

流程：合成竖版 PPT（单行 72/44pt + 多行块 lnSpc 1.0/1.2/1.5，纯黑文本）
→ COM 参考图 + 生产渲染器对比图 → 分块行投影测量：
- 各多行块 pitch → 验证 pitch/lnSpc 恒定 → ratio
- 各块首行 Δtop（同字体两侧 ascent 相同，只含基线放置误差）→ 最小二乘 a/c
输出建议的 FontMetrics(...)；更新 font_registry 后重跑验证残差（目标 ≤±5px@q2）。
"""
import os
import sys
import statistics
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from PIL import Image
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

TYPEFACES = {
    'sans': 'Source Han Sans CN',
    'serif': 'Source Han Serif CN',
}
Q = 2.0
SZ_MULTI = 28
LONG_TEXT = '第一行校准文本ABC第二行内容defgh第三行文字ijklmn第四行尾行'

# (标签, y, h, 字号, lnSpc or None, 文本, wrap)
BOXES = [
    ('single72', 0.15, 1.30, 72, None, '88,888元', False),
    ('multi10',  1.80, 2.00, 28, 1.0,  LONG_TEXT, True),
    ('multi12',  4.00, 2.20, 28, 1.2,  LONG_TEXT, True),
    ('multi15',  6.30, 2.40, 28, 1.5,  LONG_TEXT, True),
    ('single44', 8.90, 1.20, 44, None, '递交一生中意金耀版一件', False),
]


def build_test_pptx(typeface, path):
    prs = Presentation()
    prs.slide_width = Inches(7.5)
    prs.slide_height = Inches(10.0)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    for _tag, y, h, sz, lnspc, text, wrap in BOXES:
        tb = slide.shapes.add_textbox(Inches(0.3), Inches(y), Inches(6.8), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        if lnspc:
            p.line_spacing = lnspc
        run = p.add_run()
        run.text = text
        run.font.size = Pt(sz)
        run.font.name = typeface
        run.font.color.rgb = RGBColor(0, 0, 0)

    prs.save(path)
    print('合成测试 PPT:', path)


def measure(path):
    """白底图 -> 行投影，返回每行 (top, bottom)"""
    import numpy as np
    a = np.array(Image.open(path).convert('L'), dtype=np.uint8)
    mask = a < 235
    rows = mask.any(axis=1)
    lines, start = [], None
    for i, v in enumerate(rows):
        if v and start is None:
            start = i
        elif not v and start is not None:
            lines.append((start, i))
            start = None
    if start is not None:
        lines.append((start, len(rows)))
    return lines


def split_blocks(lines):
    """把测到的行按 BOXES 的 y 范围分块（px@q2，1in=96px×Q）"""
    blocks = {}
    for tag, y, h, *_ in BOXES:
        lo, hi = int(y * 96 * Q) - 30, int((y + h) * 96 * Q) + 60
        blocks[tag] = [t for t, _b in lines if lo <= t < hi]
    return blocks


def main():
    family = sys.argv[1] if len(sys.argv) > 1 else 'sans'
    typeface = TYPEFACES[family]

    try:
        from src.core.processors.com_image_exporter import COMImageExporter
        from src.core.processors.pillow_image_exporter import PillowImageExporter
        from src.core.processors import font_registry as fr
    except ImportError:
        from core.processors.com_image_exporter import COMImageExporter
        from core.processors.pillow_image_exporter import PillowImageExporter
        from core.processors import font_registry as fr

    m = fr._BUNDLED_FAMILIES[family]['metrics']
    print(f"当前 {family} 常数: ratio={m.line_height_ratio} a={m.baseline_a} "
          f"g={m.baseline_g} calibrated={m.calibrated}")

    tmp = tempfile.mkdtemp(prefix='calib_sh_')
    pptx_path = os.path.join(tmp, 'calib.pptx')
    build_test_pptx(typeface, pptx_path)

    ref_dir = os.path.join(tmp, 'ref')
    poc_dir = os.path.join(tmp, 'poc')
    print('COM 导出参考图（PowerPoint）...')
    COMImageExporter().export_to_images(pptx_path, ref_dir, 'PNG', Q)
    print('生产渲染器导出对比图...')
    PillowImageExporter().export_to_images(pptx_path, poc_dir, 'PNG', Q)

    ref = split_blocks(measure(os.path.join(ref_dir, 'slide_1.png')))
    poc = split_blocks(measure(os.path.join(poc_dir, 'slide_1.png')))

    print(f'\n--- 分块（px@q{Q:.0f}）---')
    for tag in ref:
        rn, pn = len(ref[tag]), len(poc[tag])
        if rn and pn:
            print(f'  {tag}: ref={rn}行 poc={pn}行 '
                  f'首行Δtop(poc-ref)={poc[tag][0] - ref[tag][0]:+5d}')
        else:
            print(f'  {tag}: ref={rn}行 poc={pn}行 (块缺失)')

    # 行距：各多行块 pitch / lnSpc 应恒定 -> ratio
    print('\n--- 行距 ---')
    ratio_pts = []
    for tag, _y, _h, sz, lnspc, _t, _w in BOXES:
        if lnspc is None:
            continue
        tops_r, tops_p = ref[tag], poc[tag]
        if len(tops_r) >= 2 and len(tops_p) >= 2:
            pr_ = statistics.median([tops_r[i+1] - tops_r[i] for i in range(len(tops_r)-1)])
            pp_ = statistics.median([tops_p[i+1] - tops_p[i] for i in range(len(tops_p)-1)])
            ratio_pts.append((tag, lnspc, pr_, pp_))
            print(f'  {tag}(L={lnspc}): ref_pitch={pr_:.0f} poc_pitch={pp_:.0f} '
                  f'ref_pitch/L/em={pr_/lnspc/(sz*4/3*Q):.4f}')
    # ratio 用参考侧直接计算（pitch_ref = sz_px × ratio × L）
    ratio_new = statistics.mean(
        r / (28 * 96 / 72 * Q) / l for _t, l, r, _p in ratio_pts)

    # 首行基线：仿射模型 baseline_em(L) = A + G×L（直接对 em 做最小二乘）
    # 每个测量点满足：A + G×L = 现值em(L) + d(L)，现值em = baseline_a + baseline_g×L
    print('\n--- 首行基线（em 单位）---')
    pts = []
    for tag, _y, _h, sz, lnspc, _t, _w in BOXES:
        L = lnspc if lnspc is not None else 1.0
        if not ref[tag] or not poc[tag]:
            continue
        d = (ref[tag][0] - poc[tag][0]) / (sz * 96 / 72 * Q)  # 正=需下移
        pts.append((tag, L, d))
        print(f'  {tag}(L={L}): 需修正 {d:+.4f} em')
    import numpy as np
    A = np.array([[1.0, L] for _t, L, _d in pts])
    b = np.array([m.baseline_a + m.baseline_g * L + d
                  for _t, L, d in pts])
    (a_est, g_est), _res, _rank, _sv = np.linalg.lstsq(A, b, rcond=None)
    print(f'  拟合: A={a_est:.4f} G={g_est:.4f}')
    for tag, L, d in pts:
        pred = a_est + g_est * L
        target = m.baseline_a + m.baseline_g * L + d
        print(f'  拟合残差 {tag}(L={L}): {pred - target:+.4f} em '
              f'({(pred - target) * 28 * 96 / 72 * Q:+.1f}px@q2/28pt)')

    print(f'\n=== 建议常数（{family}）===')
    print(f"FontMetrics(line_height_ratio={ratio_new:.4f}, "
          f"baseline_a={a_est:.4f}, baseline_g={g_est:.4f}, "
          f"strike_pos_em=0.245, calibrated=True)")
    print('（已是仿射模型时，拟合残差即预期收敛度；首行Δ列也应 ≤±5px@q2）')
    return 0


if __name__ == '__main__':
    sys.exit(main())
