#!/usr/bin/env python3
"""渲染库双副本单向同步脚本。

安卓 APK 用不了 pip（Chaquopy 里是直接内置 Python 源码），所以把仓库里的两个库
各复制一份进去。本脚本负责单向同步并检查漂移：

    libs/render_core/src/render_core/*.py  ->  android/app/src/main/python/render_core/
    libs/fontfx/src/fontfx/*.py            ->  android/app/src/main/python/fontfx/

- 只同步顶层 ``*.py``；跳过 ``__pycache__`` 与安卓副本专属的 ``resources/``
  （字体在 APK 内独立存放，不随代码同步）
- 目标副本统一 LF 行尾
- ``--check``：只比对不写入，存在漂移时退出码 1（CI 用它防漂移）

用法：
    python tools/sync_render_core.py            # 执行同步
    python tools/sync_render_core.py --check    # 仅校验（漂移则 exit 1）

改完 ``libs/`` 下的库代码后必须跑一次本脚本，保证安卓副本与库一致。
"""

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
ANDROID_PY = REPO_ROOT / "android" / "app" / "src" / "main" / "python"
# (库名, 源目录, 安卓目标目录)
PACKAGES = [
    ("fontfx", REPO_ROOT / "libs" / "fontfx" / "src" / "fontfx", ANDROID_PY / "fontfx"),
    ("render_core", REPO_ROOT / "libs" / "render_core" / "src" / "render_core",
     ANDROID_PY / "render_core"),
]


def normalize(text: str) -> str:
    return text.replace("\r\n", "\n")


def collect_py_files(directory: Path):
    return sorted(p for p in directory.glob("*.py") if p.is_file())


def sync_one(name, src_dir, dst_dir, check_only: bool) -> int:
    """同步一个库。返回漂移的文件数（check 模式）；源目录不存在时返回 2。"""
    if not src_dir.is_dir():
        print(f"[错误] 源目录不存在: {src_dir}")
        return 2
    src_files = collect_py_files(src_dir)
    dst_names = {p.name for p in collect_py_files(dst_dir)} if dst_dir.is_dir() else set()
    src_names = {p.name for p in src_files}

    drift = 0
    only_in_dst = sorted(dst_names - src_names)
    if only_in_dst:
        print(f"[警告] {name}: 安卓副本多出源侧没有的文件（不会删除，请人工确认）: "
              f"{', '.join(only_in_dst)}")
        drift += len(only_in_dst)
    copied = 0
    for src_path in src_files:
        dst_path = dst_dir / src_path.name
        src_text = normalize(src_path.read_text(encoding="utf-8"))
        if dst_path.exists():
            if src_text == normalize(dst_path.read_text(encoding="utf-8")):
                continue
            status = "漂移" if check_only else "更新"
        else:
            status = "缺失" if check_only else "新增"
        print(f"[{name}] [{status}] {src_path.name}")
        if check_only:
            drift += 1
        else:
            dst_dir.mkdir(parents=True, exist_ok=True)
            dst_path.write_text(src_text, encoding="utf-8", newline="\n")
            copied += 1
    if not check_only:
        print(f"[{name}] 同步完成：更新/新增 {copied} 个文件，共 {len(src_files)} 个源文件")
    return drift


def main() -> int:
    parser = argparse.ArgumentParser(description="渲染库双副本同步")
    parser.add_argument("--check", action="store_true",
                        help="只比对不写入，有漂移时退出码 1")
    args = parser.parse_args()

    total = 0
    for name, src_dir, dst_dir in PACKAGES:
        total += sync_one(name, src_dir, dst_dir, args.check)

    if args.check:
        if total:
            print(f"[结果] 检测到 {total} 处漂移，请运行 python tools/sync_render_core.py 同步")
            return 1
        print("[结果] fontfx + render_core 双副本一致")
    return 0


if __name__ == "__main__":
    sys.exit(main())
