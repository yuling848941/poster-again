#!/usr/bin/env python
"""
PyInstaller 打包脚本
将Poster Again程序打包为Windows exe文件

使用方法:
    python build_exe.py

打包完成后，exe文件将在 dist/ 目录下
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def check_dependencies():
    """检查依赖"""
    try:
        import PyInstaller
        print("[OK] PyInstaller 已安装")
    except ImportError:
        print("[ERROR] PyInstaller 未安装")
        print("请运行: pip install pyinstaller")
        sys.exit(1)

    # 检查主程序文件
    if not Path("main.py").exists():
        print("[ERROR] 未找到 main.py 文件")
        sys.exit(1)

    # 检查图标文件
    icon_path = Path("resources/icon.ico")
    if not icon_path.exists():
        print("[WARNING] 未找到图标文件 resources/icon.ico")

    print("[OK] 依赖检查完成\n")

def build_exe():
    """构建exe文件"""
    project_root = Path(__file__).parent.absolute()
    print(f"项目根目录: {project_root}\n")

    # 清理旧文件
    dist_dir = project_root / "dist"
    build_dir = project_root / "build"
    spec_file = project_root / "PosterAgain.spec"

    if dist_dir.exists():
        shutil.rmtree(dist_dir)
        print("[OK] 已清理 dist 目录")
    if build_dir.exists():
        shutil.rmtree(build_dir)
        print("[OK] 已清理 build 目录")
    if spec_file.exists():
        spec_file.unlink()
        print("[OK] 已清理 spec 文件\n")

    # 构建 PyInstaller 命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # 打包为单个exe文件
        "--windowed",                   # Windows下不显示控制台
        "--name=PosterAgain",           # exe名称
        "--icon=resources/icon.ico",    # 图标文件
        f"--distpath={dist_dir}",       # 输出目录
        f"--workpath={build_dir}",      # 临时目录
        f"--specpath={project_root}",   # spec文件目录
        "--clean",                      # 清理临时文件
        # 添加资源文件（配置文件不打包，首次运行自动生成隐藏文件）
        "--add-data=resources;resources",
        # 隐藏导入（按需打包，最小化体积）
        "--hidden-import=PySide6.QtCore",
        "--hidden-import=PySide6.QtGui",
        "--hidden-import=PySide6.QtWidgets",
        "--hidden-import=PySide6.QtNetwork",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=pptx",
        "--hidden-import=numpy",
        "--hidden-import=yaml",
        "--hidden-import=psutil",
        # 排除不需要的模块，减小体积
        "--exclude-module=PySide6.Qt3DCore",
        "--exclude-module=PySide6.Qt3DRender",
        "--exclude-module=PySide6.Qt3DInput",
        "--exclude-module=PySide6.Qt3DLogic",
        "--exclude-module=PySide6.Qt3DAnimation",
        "--exclude-module=PySide6.Qt3DExtras",
        "--exclude-module=PySide6.QtWebEngineCore",
        "--exclude-module=PySide6.QtWebEngineWidgets",
        "--exclude-module=PySide6.QtWebEngine",
        "--exclude-module=PySide6.QtMultimedia",
        "--exclude-module=PySide6.QtMultimediaWidgets",
        "--exclude-module=PySide6.QtPositioning",
        "--exclude-module=PySide6.QtLocation",
        "--exclude-module=PySide6.QtSensors",
        "--exclude-module=PySide6.QtSerialPort",
        "--exclude-module=PySide6.QtSql",
        "--exclude-module=PySide6.QtTest",
        "--exclude-module=PySide6.QtBluetooth",
        "--exclude-module=PySide6.QtNfc",
        "--exclude-module=PySide6.QtDesigner",
        "--exclude-module=PySide6.QtHelp",
        "--exclude-module=PySide6.QtOpenGL",
        "--exclude-module=PySide6.QtOpenGLWidgets",
        "--exclude-module=PySide6.QtPdf",
        "--exclude-module=PySide6.QtPdfWidgets",
        "--exclude-module=PySide6.QtQuick",
        "--exclude-module=PySide6.QtQuickWidgets",
        "--exclude-module=PySide6.QtRemoteObjects",
        "--exclude-module=PySide6.QtScxml",
        "--exclude-module=PySide6.QtStateMachine",
        "--exclude-module=PySide6.QtSvg",
        "--exclude-module=PySide6.QtSvgWidgets",
        "--exclude-module=PySide6.QtTextToSpeech",
        "--exclude-module=PySide6.QtVirtualKeyboard",
        "--exclude-module=PySide6.QtWaylandClient",
        "--exclude-module=PySide6.QtWebChannel",
        "--exclude-module=PySide6.QtWebSockets",
        "--exclude-module=PySide6.QtXml",
        # 性能优化
        "--noconfirm",                  # 覆盖已存在的文件
        # 主程序入口（路径包含空格，用引号包围）
        "main.py"
    ]

    # 添加根目录下的 .pptcfg 文件
    for pptcfg in Path(".").glob("*.pptcfg"):
        cmd.insert(14, f"--add-data={pptcfg};config")

    print("=" * 60)
    print("开始打包...")
    print("=" * 60)
    print(f"执行命令: {' '.join(cmd)}\n")

    # 执行打包
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        print("\n[ERROR] 打包失败!")
        print(result.stderr)
        sys.exit(1)

    print(result.stdout)
    print("\n" + "=" * 60)
    print("打包完成!")
    print("=" * 60)

    # 验证exe文件
    exe_file = dist_dir / "PosterAgain.exe"
    if exe_file.exists():
        size_mb = exe_file.stat().st_size / (1024 * 1024)
        print(f"\n[OK] exe文件已生成: {exe_file}")
        print(f"[OK] 文件大小: {size_mb:.1f} MB")
    else:
        print("\n[ERROR] exe文件未生成!")
        sys.exit(1)

    # 复制额外文件
    copy_extra_files(project_root, dist_dir)

    # 创建发布说明
    create_readme(dist_dir)

    print("\n" + "=" * 60)
    print("[SUCCESS] 所有文件已准备就绪！")
    print("=" * 60)
    print(f"\n请将 dist/ 目录打包分发给用户")
    print(f"用户只需要: 双击 PosterAgain.exe 即可使用\n")

def copy_extra_files(project_root, dist_dir):
    """复制额外文件到dist目录"""
    print("\n正在复制额外文件...")

    # 复制示例文件目录
    examples_src = project_root / "examples"
    if examples_src.exists():
        examples_dest = dist_dir / "examples"
        if examples_dest.exists():
            shutil.rmtree(examples_dest)
        shutil.copytree(examples_src, examples_dest)
        print(f"[OK] 已复制示例文件目录")

    print(f"[OK] 所有文件复制完成")

def create_readme(dist_dir):
    """创建用户使用说明"""
    readme_content = """Poster Again 使用说明

【快速开始】
1. 双击 "PosterAgain.exe" 启动程序
2. 无需安装任何额外软件
3. 如遇到杀毒软件警告，请选择"允许"

【注意事项】
- 请保持 config/ 目录与 exe 在同一目录下
- 如需备份配置，请备份 config/ 目录
- 支持 Windows 10/11 操作系统

【故障排除】
- 如果程序无法启动，请以管理员身份运行
- 如遇问题，请查看程序日志或联系技术支持

版本: 1.0
更新日期: 2025-11-12
"""
    readme_file = dist_dir / "使用说明.txt"
    readme_file.write_text(readme_content, encoding="utf-8")
    print(f"[OK] 已创建使用说明: {readme_file}")

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("Poster Again 打包工具")
    print("=" * 60 + "\n")

    check_dependencies()
    build_exe()