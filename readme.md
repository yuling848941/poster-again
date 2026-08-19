# PosterAgain — 保险贺报批量生成工具（Windows + Android）

根据 Excel/CSV/JSON 数据，把 PPT 模板批量生成为个性化海报：Windows 桌面端产出 .pptx 与图片，Android 端现场生成图片并直接转发。核心渲染引擎 `render_core` 两端共用，纯 Python 实现——**批量生成与图片渲染默认无需安装 Office/WPS**。

## 双端能力总览

| 能力 | Windows 桌面端 | Android 端 |
|---|---|---|
| 批量个性化 .pptx | ✅ | — |
| 批量生成图片 | ✅ 纯 Python 引擎 / Office COM 高保真（可选） | ✅ 内嵌同一引擎 |
| 数据格式 | Excel / CSV / JSON（pandas） | Excel / CSV / JSON（无 pandas，openpyxl） |
| 字段匹配 | 自动 + 手动 + 前后缀文本 | 自动 + 手动绑定/固定值/前后缀 + 记住配置 |
| 输出命名 | `序号_姓名.png/.pptx` | 同一规则，与 PC 一致 |
| 保存/分享 | 输出目录 | 相册（MediaStore）/ 系统分享（微信等） |
| 依赖 | 无 Office 也可用（COM 为可选增强） | 无任何外部依赖 |

## 渲染引擎（src/render_core/）

自包含、跨平台的 PPTX→图片渲染核心，对照 PowerPoint COM 真值逐项校准：

- **文字效果**：glow 发光（五色探针统一校准，强度比 ±4%）、外阴影、纯色/渐变/图片填充、描边、斜体、删除线、单/双下划线
- **颜色系统**：srgbClr/schemeClr/sysClr/prstClr + lumMod/lumOff/satMod、渐变同位置 stop 覆盖（PowerPoint 语义）
- **排版**：typographic 行距模型（相邻行推进 = descent+ascent）、基线仿射模型（思源/雅黑分别校准）、normAutofit 缩字
- **形状**：18 种预设几何、旋转/翻转、组合递归、纯色/线性渐变/图片填充、形状阴影
- **字体**：捆绑思源黑体/宋体（SIL OFL）+ 系统字体 + PPTX 内嵌字体（含 ODTTF 解混淆）+ 本机扩展字体目录静默扫描
- **内存受限路径**：JPEG draft 降采样、先缩放后补 alpha（安卓 50MB 模板可用）
- **溢出语义**：文本溢出文本框时渐变端点色延伸（与 PowerPoint 一致）
- **渲染预检**：白名单制，不支持的效果明确提示（倒影/3D/艺术字变形等）

跨端一致性实测：设备 vs PC（同字体）差异 0.011~0.049% 像素，纯抗锯齿级。

## 快速开始（Windows）

```bash
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py            # 运行 GUI
.venv\Scripts\python.exe -m PyInstaller PosterAgain.spec --noconfirm   # 打包 exe
```

图片导出引擎可选：`config/config.yaml` → `processing.image_export_engine`
- `auto`（默认）：检测到 Office 用 COM（100% 保真），否则纯 Python
- `office` / `python`：强制其一；COM 失败自动降级保持出图

## 快速开始（Android）

```bash
cd android
gradle assembleDebug     # 需 JDK 17 + Android SDK（local.properties 指定）
# 产物 android/app/build/outputs/apk/debug/app-debug.apk
```

技术栈：Kotlin + Jetpack Compose（Material 3，底部导航：任务/模板/设置 + 流程页绑定/渲染/预览）+ Chaquopy 16.1 嵌入 Python 3.12 渲染引擎。

界面流程：首页选模板+数据 → 字段绑定（自动匹配，未匹配可点设置：数据列/固定值/前后缀）→ 批量渲染（进度+画廊）→ 预览/保存相册/分享。字段配置按"模板名+字段指纹"记忆，模板改动自动作废。

使用文档：[docs/安卓端使用指南.docx](docs/安卓端使用指南.docx)（带界面截图）

## 占位符约定

PPT 模板三种写法（优先级从高到低）：

1. **Alt Text（推荐）**：右键形状 → 编辑替代文本 → `ph_字段名`
2. **形状名称**：选择窗格中命名 `ph_字段名`
3. **文本标记**：`{{字段名}}`（行内替换）

字段名与数据列名匹配：精确 → 去 ph_ 前缀 → 忽略大小写 → 唯一子串兜底（多候选不猜，留手动）。

## 项目结构

```
poster-again/
├── main.py                    # 桌面端入口
├── src/
│   ├── render_core/           # ★ 跨平台渲染核心（自包含，可独立嵌入）
│   │   ├── pillow_image_exporter.py   # 渲染引擎（效果/排版/内存路径）
│   │   ├── personalize.py     # Personalizer 批量个性化管线
│   │   ├── data_rows.py       # 无 pandas 数据读取
│   │   ├── font_registry.py   # 字体解析（捆绑/系统/嵌入/扩展目录）
│   │   ├── format_inheritance.py      # OOXML 格式继承解析
│   │   ├── shape_geometry.py  # 预设形状几何
│   │   └── render_precheck.py # 渲染预检
│   ├── gui/                   # PySide6 界面
│   ├── ppt_generator.py       # 批量生成编排
│   └── core/                  # 处理器（COM 分层导出等）
├── android/                   # ★ Android 工程
│   ├── app/src/main/java/     # Kotlin/Compose UI（screens + AppState）
│   ├── app/src/main/python/   # poster_bridge 桥接 + render_core 副本
│   └── e2e_batch.py           # 模拟器自动化测试
├── resources/fonts/           # 捆绑思源字体（SIL OFL）
├── tests/unit/                # 40+ 独立测试脚本（python 直接运行）
├── openspec/                  # 规格变更记录
└── docs/                      # 文档与截图
```

## 测试

无测试框架配置，全部为独立脚本：

```bash
.venv\Scripts\python.exe tests\unit\test_render_core.py        # 引擎
.venv\Scripts\python.exe tests\unit\test_font_registry.py      # 字体（23 项）
.venv\Scripts\python.exe tests\unit\test_poster_bridge.py      # 安卓桥接
# ...共 40+ 文件，改哪测哪
```

## 版本历史

### v3.x（2026-08）
- 跨平台：render_core 引擎成包，Android 端（Chaquopy）端到端
- 渲染保真：glow/渐变颜色/溢出/typographic 行距等对照 COM 真值修复
- 字体系统：捆绑思源 + PPTX 内嵌字体 + 近名容错
- Windows 图片导出 COM 分层（auto/office/python）
- Android：批量 UI、字段配置（绑定/固定值/前后缀/记住配置）、相册/分享

### v2.0（2026-03）
- 纯 Python PPT 生成，脱离 Office/WPS

### v1.x
- COM 接口时代

## 许可证

MIT License（捆绑思源字体遵循 SIL OFL 1.1）
