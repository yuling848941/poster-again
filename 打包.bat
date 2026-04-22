@echo off
chcp 65001 >nul
title Poster Again - 自动打包工具

echo.
echo ========================================
echo     Poster Again 自动打包工具
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [OK] Python 环境检查通过
echo.

REM 检查是否在项目根目录
if not exist "main.py" (
    echo [错误] 请在项目根目录运行此脚本
    pause
    exit /b 1
)

echo [OK] 项目目录检查通过
echo.

REM 安装/更新 PyInstaller
echo 正在检查 PyInstaller...
python -m pip install --upgrade pyinstaller -q
if errorlevel 1 (
    echo [错误] PyInstaller 安装失败
    pause
    exit /b 1
)

echo [OK] PyInstaller 已就绪
echo.

REM 生成图标（如果不存在）
if not exist "resources\icon.ico" (
    echo 正在生成图标...
    python create_icon.py
    echo.
)

REM 运行打包脚本
echo 开始打包...
echo ========================================
echo.
python build_exe.py

echo.
echo ========================================
if exist "dist\PosterAgain.exe" (
    echo [成功] 打包完成！
    echo.
    echo 打包文件位置: dist\PosterAgain.exe
    echo 文件大小:
    dir "dist\PosterAgain.exe" | find "PosterAgain.exe"
    echo.
    echo 按任意键打开输出目录...
    pause >nul
    explorer dist
) else (
    echo [失败] 打包失败，请检查错误信息
    echo.
    pause
)
echo ========================================