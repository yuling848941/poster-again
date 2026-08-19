# -*- coding: utf-8 -*-
"""
渲染引擎内存路径优化（_load_rgba）单元测试
验证：与旧路径（open→convert('RGBA')→resize）逐像素等价（含 alpha/调色板图）、
JPEG draft 降采样正确性、以及大图渲染的峰值内存受控（移动端 OOM 防护）。

运行：.venv/Scripts/python.exe tests/unit/test_renderer_memory.py
"""
import io
import os
import subprocess
import sys
import tempfile
import threading
import time

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

import numpy as np
from PIL import Image

try:
    from src.render_core.pillow_image_exporter import _load_rgba
    from src.render_core import render
except ImportError:
    from render_core.pillow_image_exporter import _load_rgba
    from render_core import render

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


def old_path(blob, tw, crop_rect=None):
    """旧实现（对照）：open → convert('RGBA') → [crop] → resize"""
    img = Image.open(io.BytesIO(blob)).convert('RGBA')
    if crop_rect:
        w0, h0 = img.size
        l, t, r, b = crop_rect
        img = img.crop((int(w0 * l), int(h0 * t),
                        int(w0 * (1 - r)), int(h0 * (1 - b))))
    return img.resize((max(1, tw[0]), max(1, tw[1])), Image.LANCZOS)


def main():
    print('(1) 像素等价：不透明 RGB / 带 alpha / 调色板透明')
    tw = (333, 210)
    # RGB 渐变 PNG
    src = Image.fromarray(np.stack([
        np.tile(np.linspace(0, 255, 400, dtype=np.uint8), (300, 1))] * 3, axis=-1))
    buf = io.BytesIO(); src.save(buf, 'PNG')
    a = np.array(_load_rgba(buf.getvalue(), tw))
    b = np.array(old_path(buf.getvalue(), tw))
    check('RGB PNG 逐像素一致', bool((a == b).all()))
    # RGBA PNG（半透明渐变）
    arr = np.zeros((300, 400, 4), np.uint8)
    arr[..., 0] = 200; arr[..., 3] = np.tile(np.linspace(0, 255, 400, dtype=np.uint8), (300, 1))
    buf = io.BytesIO(); Image.fromarray(arr, 'RGBA').save(buf, 'PNG')
    a = np.array(_load_rgba(buf.getvalue(), tw))
    b = np.array(old_path(buf.getvalue(), tw))
    check('RGBA PNG 逐像素一致', bool((a == b).all()))
    # 调色板带透明 PNG
    p = src.convert('P')
    p.info['transparency'] = 100
    buf = io.BytesIO(); p.save(buf, 'PNG')
    a = np.array(_load_rgba(buf.getvalue(), tw))
    b = np.array(old_path(buf.getvalue(), tw))
    check('P+透明 PNG 逐像素一致', bool((a == b).all()))
    # srcRect 裁剪等价
    src2 = Image.fromarray(np.tile(np.linspace(0, 255, 500, dtype=np.uint8), (400, 1)))
    buf = io.BytesIO(); src2.save(buf, 'PNG')
    crop = (0.1, 0.2, 0.3, 0.0)
    a = np.array(_load_rgba(buf.getvalue(), tw, crop_rect=crop))
    b = np.array(old_path(buf.getvalue(), tw, crop_rect=crop))
    check('srcRect 裁剪路径逐像素一致', bool((a == b).all()))

    print('(2) JPEG draft 降采样')
    src3 = Image.fromarray(np.stack([
        np.tile(np.linspace(0, 255, 1600, dtype=np.uint8), (1200, 1))] * 3, axis=-1))
    buf = io.BytesIO(); src3.save(buf, 'JPEG', quality=92)
    fast = np.array(_load_rgba(buf.getvalue(), (400, 300)).convert('RGB'))
    ref = np.array(old_path(buf.getvalue(), (400, 300)).convert('RGB'))
    diff = np.abs(fast.astype(int) - ref.astype(int)).mean()
    check(f'JPEG draft 尺寸正确且视觉等价（平均偏差 {diff:.2f}/255 < 8）', diff < 8.0)

    print('(3) 大图渲染峰值内存（4000x6000 背景，子进程内核峰值计数）')
    print('    内存模型：峰值 ≈ 解释器+pptx急切加载(≈1.1×文件) + 全分辨率解码(W×H×3)')
    print('    —— python-pptx 打开时全量加载部件、Pillow PNG 解码不可降采样，为体系下限；')
    print('    本优化的收益 = 消除旧路径的全分辨率 RGBA 副本(+96MB)；JPEG 另有 draft 解码降采样。')
    import psutil  # noqa: F401  （子进程内使用）

    def measure_peak(bg_path, suffix):
        from pptx import Presentation
        from pptx.util import Inches
        prs = Presentation()
        prs.slide_width = Inches(7.5)
        prs.slide_height = Inches(10.0)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(bg_path, 0, 0, Inches(7.5), Inches(10.0))
        pptx_path = os.path.join(tempfile.mkdtemp(), f'mem_{suffix}.pptx')
        prs.save(pptx_path)
        size_mb = os.path.getsize(pptx_path) // 1048576

        script = (
            "import sys\n"
            f"sys.path.insert(0, {os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))!r})\n"
            "from src.render_core import render\n"
            f"r = render({pptx_path!r}, {tempfile.mkdtemp()!r}, 'PNG', 1.0)\n"
            "assert len(r.images) == 1\n"
            "import psutil\n"
            "mi = psutil.Process().memory_info()\n"
            "print('PEAK_MB=%d' % (getattr(mi, 'peak_wset', mi.rss) // 1048576))\n"
            "print('MEM_OK')\n"
        )
        proc = subprocess.run([sys.executable, '-c', script],
                              capture_output=True, text=True)
        ok = 'MEM_OK' in proc.stdout
        peak_mb = 0
        if ok:
            peak_mb = int([l for l in proc.stdout.splitlines()
                           if l.startswith('PEAK_MB=')][0].split('=')[1])
        return ok, peak_mb, size_mb

    big = np.random.default_rng(42).integers(0, 255, (6000, 4000, 3), dtype=np.uint8)
    tmpd = tempfile.mkdtemp()
    png_path = os.path.join(tmpd, 'big.png')
    Image.fromarray(big).save(png_path, 'PNG')
    jpg_path = os.path.join(tmpd, 'big.jpg')
    Image.fromarray(big).save(jpg_path, 'JPEG', quality=90)
    del big

    ok, png_peak, png_size = measure_peak(png_path, 'png')
    check(f'PNG 渲染成功（pptx {png_size}MB）', ok)
    if ok:
        # 模型下限 ≈ 45(解释器+库) + 1.1×68(pptx加载) + 91(解码) + 画布/杂项 ≈ 384 实测
        # 旧路径另需 +96MB 全分辨率 RGBA 副本（≈480）；420 给环境波动留余量
        check(f'PNG 峰值 {png_peak}MB < 420MB（旧路径量级 ~480MB）', png_peak < 420)

    ok, jpg_peak, jpg_size = measure_peak(jpg_path, 'jpg')
    check(f'JPEG 渲染成功（pptx {jpg_size}MB）', ok)
    if ok:
        # JPEG draft 解码降采样：跳过全分辨率解码（~91MB），仅按 draft 档位解码
        check(f'JPEG 峰值 {jpg_peak}MB < 250MB（draft 生效，显著低于 PNG 路径）',
              jpg_peak < 250)

    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
