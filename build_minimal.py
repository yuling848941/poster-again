#!/usr/bin/env python
"""
极简打包脚本 - 只打包核心功能
"""

import os
import shutil
from pathlib import Path

print("开始极简打包...")

# 清理
for folder in ["dist", "build", "__pycache__"]:
    if os.path.exists(folder):
        shutil.rmtree(folder)
        print(f"[OK] 清理 {folder}")

# 简单的PyInstaller命令（无复杂参数）
cmd = '''
python -m PyInstaller --onefile --windowed --name=PosterAgain --icon=resources/icon.ico --clean main.py
'''

print("执行命令...")
os.system(cmd)

# 检查结果
exe_path = Path("dist/PosterAgain.exe")
if exe_path.exists():
    size_mb = exe_path.stat().st_size / (1024*1024)
    print(f"\n[SUCCESS] 打包成功！")
    print(f"位置: {exe_path}")
    print(f"大小: {size_mb:.1f} MB")
    print("\n正在复制配置文件...")

    # 复制配置（如果存在）
    if os.path.exists("config"):
        os.makedirs("dist/config", exist_ok=True)
        for f in Path("config").glob("*"):
            if f.is_file():
                shutil.copy(f, "dist/config/")
                print(f"[OK] 复制 {f.name}")

    # 复制pptcfg
    for pptcfg in Path(".").glob("*.pptcfg"):
        shutil.copy(pptcfg, "dist/")
        print(f"[OK] 复制 {pptcfg.name}")

    print("\n[FINISHED] 完成！dist/ 目录即可分发给用户")
else:
    print("\n[ERROR] 打包失败，请检查错误信息")