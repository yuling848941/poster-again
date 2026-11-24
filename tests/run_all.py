#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全量单元测试一键回归：子进程逐个跑 tests/unit/*.py，汇总结果表；
然后接着跑仓库里两个库各自的测试（libs/fontfx、libs/render_core）。

用法：
    python tests/run_all.py                 # 默认（跳过 SKIP_LIST 中的用例）
    python tests/run_all.py --all           # 包含跳过清单
    python tests/run_all.py --filter mru    # 只跑文件名含 mru 的（仅应用用例）
    python tests/run_all.py --list          # 只列出将跑的用例
    python tests/run_all.py --timeout 420   # 单文件超时秒数（默认 300）
    python tests/run_all.py --app-only      # 不跑库测试

退出码：有 FAIL/TIMEOUT（或任一库测试失败）则 1，否则 0（SKIP 不影响）。
"""

import argparse
import subprocess
import sys
import time
from pathlib import Path

# Windows 控制台默认 GBK，打印中文会 UnicodeEncodeError；统一按 UTF-8 输出
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parents[1]
TESTS_DIR = REPO_ROOT / "tests" / "unit"

# 仓库内两个库的测试入口（各自独立，也可以单独运行）
LIB_RUNNERS = [
    REPO_ROOT / "libs" / "fontfx" / "tests" / "run_tests.py",
    REPO_ROOT / "libs" / "render_core" / "tests" / "run_tests.py",
]

# 默认跳过的用例（--all 可包含）：
# - 需要本机真实 PowerPoint COM / 特定业务文件的用例，CI 与无 Office 环境跑不了
SKIP_LIST = {
    "test_com_processor.py",        # 真实 COM（需本机 PowerPoint）
}


def discover(filter_str: str = "", include_skipped: bool = False):
    files = sorted(p for p in TESTS_DIR.glob("*.py")
                   if p.name != "__init__.py" and not p.name.startswith("_"))
    if filter_str:
        files = [p for p in files if filter_str in p.name]
    if not include_skipped:
        files = [p for p in files if p.name not in SKIP_LIST]
    return files


def main() -> int:
    parser = argparse.ArgumentParser(description="全量单元测试回归")
    parser.add_argument("--all", action="store_true", help="包含默认跳过清单")
    parser.add_argument("--filter", default="", help="只跑文件名含该子串的用例")
    parser.add_argument("--list", action="store_true", help="只列出用例")
    parser.add_argument("--timeout", type=int, default=300, help="单文件超时秒数")
    parser.add_argument("--app-only", action="store_true",
                        help="只跑应用用例，不跑 libs/ 下两个库的测试")
    args = parser.parse_args()

    files = discover(args.filter, args.all)
    if args.list:
        for p in files:
            print(p.name)
        return 0

    print(f"共 {len(files)} 个用例（python {sys.version.split()[0]}，"
          f"超时 {args.timeout}s/文件）\n" + "=" * 64)

    results = []  # (name, status, seconds, note)
    for i, path in enumerate(files, 1):
        t0 = time.time()
        try:
            proc = subprocess.run(
                [sys.executable, str(path)],
                cwd=str(REPO_ROOT),
                capture_output=True, text=True,
                timeout=args.timeout, encoding="utf-8", errors="replace")
            dt = time.time() - t0
            # 约定：测试脚本失败时 sys.exit(1)
            ok = (proc.returncode == 0)
            status = "PASS" if ok else "FAIL"
        except subprocess.TimeoutExpired as e:
            dt = time.time() - t0
            status = "TIMEOUT"
            proc = e
        results.append((path.name, status, dt, ""))

        mark = {"PASS": "✅", "FAIL": "❌", "TIMEOUT": "⏰"}[status]
        print(f"[{i:>2}/{len(files)}] {mark} {path.name:<44} {dt:6.1f}s")
        if status != "PASS":
            out = (getattr(proc, "output", None) or getattr(proc, "stdout", "") or "")
            err = (getattr(proc, "stderr", "") or "")
            tail = "\n".join(((out + err).strip().splitlines())[-15:])
            print("      " + tail.replace("\n", "\n      "))

    # 汇总
    passed = [r for r in results if r[1] == "PASS"]
    failed = [r for r in results if r[1] == "FAIL"]
    timeouts = [r for r in results if r[1] == "TIMEOUT"]
    total_sec = sum(r[2] for r in results)
    print("=" * 64)
    print(f"结果：{len(passed)} 通过 / {len(failed)} 失败 / {len(timeouts)} 超时"
          f"（共 {len(results)}，耗时 {total_sec:.0f}s，"
          f"跳过 {len(SKIP_LIST if not args.all else set())} 个）")
    if failed or timeouts:
        print("失败用例：" + ", ".join(r[0] for r in failed + timeouts))
    app_ok = not (failed or timeouts)

    # 两个库各自的测试（输出直接透传，它们自己会打印逐条结果）
    libs_failed = []
    if not args.app_only:
        for runner in LIB_RUNNERS:
            if not runner.exists():
                continue
            print(f"\n【库测试】{runner.relative_to(REPO_ROOT)}")
            proc = subprocess.run([sys.executable, str(runner)], cwd=str(REPO_ROOT))
            if proc.returncode != 0:
                libs_failed.append(runner.parent.parent.name)
        print("\n" + "=" * 64)
        print("库测试：" + ("全部通过 ✅" if not libs_failed else "失败 ❌ " + ", ".join(libs_failed)))

    print("=" * 64)
    if not app_ok or libs_failed:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
