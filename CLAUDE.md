# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**本项目的完整开发指南见 [AGENTS.md](AGENTS.md)**（命令、架构、模块边界、占位符约定、测试、打包、注意事项）。两份文档不再各自维护，以 AGENTS.md 为准。

## 一句话概述

PosterAgain（v1.1.0）：根据 Excel/CSV/JSON 数据把 PPT 模板批量生成为个性化贺报。Windows 桌面端（PySide6）产出 .pptx 与图片，Android 端（Kotlin/Compose + Chaquopy）现场出图。

## 最关键的事实

- **渲染分两层库**（`libs/`，可编辑安装进 `.venv`）：`fontfx`（纯 Pillow 文字特效：渐变/发光/阴影/描边/图片填充）+ `render_core`（PPTX→图片引擎，依赖 fontfx）。安卓内嵌两者副本，改完跑 `tools/sync_render_core.py` 同步。
- **图片导出引擎** `processing.image_export_engine`：`auto`（默认，有 Office 用 COM 高保真，否则内置渲染）/ `office` / `python`。
- **所有命令用 `.venv`**（系统 Python 有 PyTorch 等会污染打包）。
- **全量回归**：`.venv\Scripts\python.exe tests\run_all.py`（应用 + 两个库）。改渲染引擎后跑 `tools/render_regression.py` 像素回归。
- 业务样例在 `samples/`（templates/data/presets，gitignored）；运行时配置快照在根目录（gitignored，勿提交）。
