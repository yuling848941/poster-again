#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pptx-render-core 全量测试：逐个跑 tests/test_*.py（每个都是独立脚本），汇总结果。

用法：python libs/render_core/tests/run_tests.py [--filter 关键字]
退出码：有失败则 1。

其中带「COM 交叉验证」的用例只在同时存在宿主应用（PosterAgain）且装了 PowerPoint 的
机器上才会真的跑；单独使用本库时自动退化为绝对值断言，不算失败。
"""

import argparse
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent


def main() -> int:
    ap = argparse.ArgumentParser(description="pptx-render-core 全量测试")
    ap.add_argument("--filter", default="", help="只跑文件名含该子串的用例")
    ap.add_argument("--timeout", type=int, default=600, help="单文件超时秒数")
    args = ap.parse_args()
    files = sorted(p for p in HERE.glob("test_*.py")
                   if args.filter in p.name and not p.name.startswith("_"))
    failed, total_sec = [], 0.0
    for i, path in enumerate(files, 1):
        t0 = time.time()
        try:
            proc = subprocess.run([sys.executable, str(path)], cwd=str(HERE.parent),
                                  capture_output=True, text=True, timeout=args.timeout,
                                  encoding="utf-8", errors="replace")
            ok = proc.returncode == 0
        except subprocess.TimeoutExpired:
            ok, proc = False, None
        dt = time.time() - t0
        total_sec += dt
        if not ok:
            failed.append(path.name)
        print(f"[{i}/{len(files)}] {'✅' if ok else '❌'} {path.name:<40} {dt:5.1f}s")
        if not ok:
            out = ((getattr(proc, "stdout", "") or "") + (getattr(proc, "stderr", "") or ""))
            print("      " + "\n      ".join(out.strip().splitlines()[-15:]))
    print("=" * 72)
    print(f"结果：{len(files) - len(failed)} 通过 / {len(failed)} 失败"
          f"（共 {len(files)}，耗时 {total_sec:.0f}s）")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
