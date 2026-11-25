#!/usr/bin/env python3
"""
PyInstaller 打包脚本 — 将 PosterAgain 打包为 Windows exe

设计要点：
1. 优先使用 .venv 虚拟环境的 Python（系统 Python 装了 PyTorch 等无关库，
   会污染打包结果，导致 exe 体积爆炸或打包失败）
2. 使用手工维护的 PosterAgain.spec（含体积优化：排除 torch/scipy 等，
   过滤 JDK DLL）。不要用命令行参数重新生成 spec，那会覆盖这些优化。
3. 日志实时输出（不捕获），便于诊断失败

使用方式：
    python build_exe.py        # 自动用 .venv（推荐）
    打包.bat                    # Windows 批处理封装

打包完成后，exe 文件在 dist/ 目录下。
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path


def find_venv_python(project_root: Path) -> str | None:
    """
    优先返回 .venv 的 Python 路径。
    若 .venv 不存在或损坏，返回 None（调用方决定是否回退）。
    """
    candidates = [
        project_root / ".venv" / "Scripts" / "python.exe",   # Windows
        project_root / ".venv" / "bin" / "python",            # Linux/macOS
    ]
    for p in candidates:
        if p.exists():
            return str(p)
    return None


def check_dependencies(python_exe: str) -> bool:
    """检查打包所需的依赖"""
    try:
        result = subprocess.run(
            [python_exe, "-c", "import PyInstaller; print(PyInstaller.__version__)"],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print("[ERROR] PyInstaller 未安装")
            print(f"请运行: {python_exe} -m pip install pyinstaller")
            return False
        print(f"[OK] PyInstaller {result.stdout.strip()}")

        # 检查关键运行依赖
        result = subprocess.run(
            [python_exe, "-c",
             "import PySide6, pptx, pandas, openpyxl, yaml, psutil; "
             "print('运行依赖齐全')"],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print("[ERROR] 运行依赖缺失:")
            print(result.stderr)
            print(f"请运行: {python_exe} -m pip install -r requirements.txt")
            return False
        print(f"[OK] {result.stdout.strip()}")
        return True

    except Exception as e:
        print(f"[ERROR] 依赖检查失败: {e}")
        return False


def build_exe(project_root: Path, python_exe: str):
    """使用 PosterAgain.spec 构建 exe"""
    dist_dir = project_root / "dist"
    build_dir = project_root / "build"
    spec_file = project_root / "PosterAgain.spec"

    print(f"\n项目根目录: {project_root}")
    print(f"使用 Python: {python_exe}\n")

    # 检查 spec 文件（这是手工维护的，包含体积优化）
    if not spec_file.exists():
        print(f"[ERROR] 未找到 {spec_file}")
        print("PosterAgain.spec 是手工维护的打包配置（含体积优化），不能缺失。")
        sys.exit(1)
    print(f"[OK] 使用 spec 文件: {spec_file.name}")

    # 清理旧产物（注意：只清理 build/dist，不动 spec）
    for d, name in [(dist_dir, "dist"), (build_dir, "build")]:
        if d.exists():
            shutil.rmtree(d)
            print(f"[OK] 已清理 {name} 目录")

    print("\n" + "=" * 60)
    print("开始打包（日志实时输出）...")
    print("=" * 60 + "\n")

    # 用 spec 打包，日志实时输出（不 capture，便于诊断）
    cmd = [python_exe, "-m", "PyInstaller", str(spec_file), "--noconfirm"]
    result = subprocess.run(cmd)

    if result.returncode != 0:
        print("\n" + "=" * 60)
        print("[ERROR] 打包失败！请查看上方日志")
        print("=" * 60)
        sys.exit(1)

    print("\n" + "=" * 60)
    print("打包完成！")
    print("=" * 60)

    # 验证 exe
    exe_file = dist_dir / "PosterAgain.exe"
    if not exe_file.exists():
        print(f"\n[ERROR] exe 未生成: {exe_file}")
        sys.exit(1)

    size_mb = exe_file.stat().st_size / (1024 * 1024)
    print(f"\n[OK] exe 已生成: {exe_file}")
    print(f"[OK] 文件大小: {size_mb:.1f} MB")

    if size_mb > 500:
        print(f"\n[警告] exe 体积偏大（{size_mb:.0f} MB > 500 MB）")
        print("可能原因：未用 .venv 打包，或环境里装了 PyTorch 等大型库。")
        print("建议用 .venv 重新打包。")

    # 复制额外文件 & 生成说明
    copy_extra_files(project_root, dist_dir)
    create_readme(dist_dir)

    print("\n" + "=" * 60)
    print("[SUCCESS] 所有文件已就绪！")
    print("=" * 60)
    print(f"\n分发方式：把 dist/ 目录整个打包给用户")
    print(f"用户使用：双击 PosterAgain.exe 即可\n")


def copy_extra_files(project_root: Path, dist_dir: Path):
    """复制额外文件到 dist 目录"""
    print("\n正在复制额外文件...")

    examples_src = project_root / "examples"
    if examples_src.exists():
        examples_dest = dist_dir / "examples"
        if examples_dest.exists():
            shutil.rmtree(examples_dest)
        shutil.copytree(examples_src, examples_dest)
        print(f"[OK] 已复制示例文件目录")

    print(f"[OK] 额外文件处理完成")


def create_readme(dist_dir: Path):
    """创建用户使用说明"""
    readme_content = """PosterAgain 使用说明（v1.1.0）

【快速开始】
1. 双击 "PosterAgain.exe" 启动程序
2. 无需安装 Microsoft Office、WPS 或任何其他软件，开箱即用
3. 如遇到杀毒软件警告，请选择"允许"

【基本流程】
1. 选择 PPT 模板文件、Excel/CSV 数据文件、输出目录
   - "最近"按钮可直接选用上次的文件/目录，通常三步点击即可开始
2. 点击"自动匹配"，核对模板占位符与数据列的对应关系
3. 勾选"直接生成图片"可输出 PNG/JPG 海报（默认引擎为内置纯 Python 渲染，
   不需要 Office）；不勾选则输出 PPT 文件
4. 点击"批量生成"，等待完成

【批量容错：单行失败不影响整批】
- 某一行数据异常时会跳过该行并继续生成其余行，完成后弹窗列出
  失败行（Excel 行号 + 原因）
- 点击弹窗中的"重试失败行"可只重新生成失败的行

【近似渲染提示】
- 部分模板效果（如艺术字变形、图表、特殊字体）为近似渲染，生成完成时
  会提示"本模板有 N 项近似渲染效果"
- 如需与 PowerPoint 完全一致的输出，可在配置文件中把
  processing.image_export_engine 设为 office（需要本机安装 Office）

【注意事项】
- 程序首次运行会在 exe 同目录生成隐藏配置文件（路径记忆等状态）
- 内置渲染引擎使用电脑上已安装的字体：模板里用到的字体请先安装到系统，
  没安装的会用微软雅黑代替（生成完成时会提示哪些字体缺失）
- 支持 Windows 10/11 操作系统
- 更新版本：用新版 exe 直接覆盖旧文件即可，配置与最近使用记录保留

【故障排除】
- 如果程序无法启动，请以管理员身份运行
- 生成的图片缺少部分效果？属近似渲染范围，见上方"近似渲染提示"
- 如遇其他问题，请联系技术支持
"""
    readme_file = dist_dir / "使用说明.txt"
    readme_file.write_text(readme_content, encoding="utf-8")
    print(f"[OK] 已创建使用说明: {readme_file.name}")


def main():
    project_root = Path(__file__).parent.absolute()

    print("=" * 60)
    print("PosterAgain 打包工具")
    print("=" * 60 + "\n")

    # 优先用 .venv（避免系统 Python 的 PyTorch 等无关库污染打包）
    venv_python = find_venv_python(project_root)
    if venv_python:
        print(f"[OK] 检测到虚拟环境: {venv_python}")
        python_exe = venv_python
    else:
        print("[警告] 未找到 .venv 虚拟环境")
        print("[警告] 将使用系统 Python 打包 — 可能卷入 PyTorch 等无关库，")
        print("[警告] 导致 exe 体积异常巨大（GB 级）或打包失败。")
        print("[警告] 强烈建议先创建虚拟环境：")
        print("        python -m venv .venv")
        print(f"        .venv\\Scripts\\python.exe -m pip install -r requirements.txt")
        print()
        resp = input("仍要用系统 Python 继续？(y/N): ").strip().lower()
        if resp != "y":
            print("已取消。请先创建 .venv。")
            sys.exit(0)
        python_exe = sys.executable
        print(f"[继续] 使用系统 Python: {python_exe}\n")

    # 依赖检查
    if not check_dependencies(python_exe):
        sys.exit(1)

    # 打包
    build_exe(project_root, python_exe)


if __name__ == "__main__":
    main()
