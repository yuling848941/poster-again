#!/usr/bin/env python3
"""render_core 双副本单向同步脚本（openspec add-render-core-android 任务 4.3）。

把 ``src/render_core/`` 的 ``*.py`` 单向同步到安卓内嵌副本
``android/app/src/main/python/render_core/``：

- 只同步顶层 ``*.py``；跳过 ``__pycache__`` 与安卓副本专属的 ``resources/``
  （字体在 APK 内独立存放，不随引擎同步）
- 目标副本统一 LF 行尾（安卓侧现状）
- ``--check``：只比对不写入，存在漂移时退出码 1（可挂进 CI / run_all 防漂移）

用法：
    python tools/sync_render_core.py            # 执行同步
    python tools/sync_render_core.py --check    # 仅校验（漂移则 exit 1）

改完 ``src/render_core/`` 后必须跑一次本脚本，保证两份引擎一致。
"""

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = REPO_ROOT / "src" / "render_core"
DST_DIR = REPO_ROOT / "android" / "app" / "src" / "main" / "python" / "render_core"


def normalize(text: str) -> str:
    return text.replace("\r\n", "\n")


def collect_py_files(directory: Path):
    return sorted(p for p in directory.glob("*.py") if p.is_file())


def main() -> int:
    parser = argparse.ArgumentParser(description="render_core 双副本同步")
    parser.add_argument("--check", action="store_true",
                        help="只比对不写入，有漂移时退出码 1")
    args = parser.parse_args()

    if not SRC_DIR.is_dir():
        print(f"[错误] 源目录不存在: {SRC_DIR}")
        return 2
    if not DST_DIR.is_dir():
        print(f"[错误] 目标目录不存在: {DST_DIR}")
        return 2

    src_files = collect_py_files(SRC_DIR)
    dst_files = collect_py_files(DST_DIR)
    src_names = {p.name for p in src_files}
    dst_names = {p.name for p in dst_files}

    drift = []

    only_in_dst = sorted(dst_names - src_names)
    if only_in_dst:
        print(f"[警告] 安卓副本多出源侧没有的文件（不会删除，请人工确认）: "
              f"{', '.join(only_in_dst)}")
        drift.extend(only_in_dst)

    copied = []
    for src_path in src_files:
        dst_path = DST_DIR / src_path.name
        src_text = normalize(src_path.read_text(encoding="utf-8"))
        if dst_path.exists():
            dst_text = normalize(dst_path.read_text(encoding="utf-8"))
            if src_text == dst_text:
                continue
            status = "漂移" if args.check else "更新"
        else:
            status = "新增" if not args.check else "缺失"
        print(f"[{status}] {src_path.name}")
        if not args.check:
            dst_path.write_text(src_text, encoding="utf-8", newline="\n")
            copied.append(src_path.name)
        else:
            drift.append(src_path.name)

    if args.check:
        if drift:
            print(f"[结果] 检测到 {len(drift)} 个文件漂移，"
                  f"请运行 python tools/sync_render_core.py 同步")
            return 1
        print("[结果] 双副本一致")
        return 0

    print(f"[结果] 同步完成：更新 {len(copied)} 个文件，"
          f"共 {len(src_files)} 个源文件")
    return 0


if __name__ == "__main__":
    sys.exit(main())
