#!/usr/bin/env python
"""
简化版打包脚本 - 避免复杂参数
"""

import os
import sys
import shutil
from pathlib import Path

def simple_build():
    """简化打包流程"""
    print("开始简化打包...\n")

    # 清理
    for folder in ["dist", "build", "__pycache__"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"[OK] 已删除 {folder}")

    for file in ["PosterAgain.spec"]:
        if os.path.exists(file):
            os.remove(file)
            print(f"[OK] 已删除 {file}")

    print()

    # 直接执行系统命令
    os.system("""
python -m PyInstaller \
    --onefile \
    --windowed \
    --name=PosterAgain \
    --icon=resources/icon.ico \
    --clean \
    --add-data="config/config.yaml;config" \
    --add-data="src/*.pptcfg;config" \
    --hidden-import=PySide6.QtCore \
    --hidden-import=PySide6.QtGui \
    --hidden-import=PySide6.QtWidgets \
    --hidden-import=pandas \
    --hidden-import=openpyxl \
    --hidden-import=pptx \
    --hidden-import=pyyaml \
    main.py
    """)

    # 验证结果
    exe_path = Path("dist/PosterAgain.exe")
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024*1024)
        print(f"\n[SUCCESS] 打包成功!")
        print(f"文件位置: {exe_path}")
        print(f"文件大小: {size_mb:.1f} MB\n")

        # 复制配置
        if os.path.exists("config"):
            shutil.copytree("config", "dist/config")
            print("[OK] 已复制 config 目录")

        # 复制pptcfg
        for pptcfg in Path(".").glob("*.pptcfg"):
            shutil.copy(pptcfg, "dist/")
            print(f"[OK] 已复制 {pptcfg.name}")

        # 复制示例
        if os.path.exists("examples"):
            shutil.copytree("examples", "dist/examples")
            print("[OK] 已复制 examples 目录")

        print("\n[FINISHED] 打包完成!")
    else:
        print("\n[ERROR] 打包失败!")

if __name__ == "__main__":
    simple_build()