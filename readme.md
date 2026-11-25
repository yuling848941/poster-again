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

## 渲染库（libs/，两层，均可 pip 安装）

- **`libs/fontfx`（fontfx）**：纯 Pillow 文字特效库——线性渐变 / 图片填充 / 发光 / 外阴影 / 描边 / 删除线 / 下划线 / 合成斜体。不依赖 PPTX，任何 Pillow 项目都能单独用。
- **`libs/render_core`（render_core，发行名 pptx-render-core）**：PPTX→图片渲染引擎，负责格式继承、排版、预检、形状几何、字体注册、批量个性化；「字长什么样」交给 fontfx。

两者对照 PowerPoint COM 真值逐项校准：

- **文字效果**：glow 发光（五色探针统一校准，强度比 ±4%）、外阴影、纯色/渐变/图片填充、描边、斜体、删除线、单/双下划线
- **颜色系统**：srgbClr/schemeClr/sysClr/prstClr + lumMod/lumOff/satMod、渐变同位置 stop 覆盖（PowerPoint 语义）
- **排版**：typographic 行距模型（相邻行推进 = descent+ascent）、基线仿射模型（思源/雅黑分别校准）、normAutofit 缩字
- **形状**：18 种预设几何、旋转/翻转、组合递归、纯色/线性渐变/图片填充、形状阴影
- **字体**：捆绑思源黑体/宋体（SIL OFL，在 `render_core/resources/fonts/`）+ 系统字体 + PPTX 内嵌字体（含 ODTTF 解混淆）+ 本机扩展字体目录静默扫描
- **批量提速**：模板只解析一次、逐行「替换→渲染→恢复」；背景图缩放后跨行缓存；已压缩媒体保存时不再重复 deflate（50MB 模板约 6.1s/行 → 1.1s/行，输出逐像素不变）
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

图片导出引擎：界面「直接生成图片」里有下拉框可直接切换，也对应运行时配置 `processing.image_export_engine`
- `python`（默认）：内置渲染，快、无需 Office；贺报模板已对照 PowerPoint 校准，效果基本一致
- `office`：Office/WPS COM 高保真（100% 保真，但每行有自动化开销，较慢）
- `auto`：检测到 Office 用 COM，否则内置
- 环境变量 `POSTERAGAIN_IMAGE_ENGINE` 可临时覆盖；office/auto 下 COM 失败自动降级内置保持出图

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
├── main.py                    # 桌面端入口（文件日志 + 全局异常钩子）
├── libs/
│   ├── fontfx/                # ★ 纯 Pillow 文字特效库（渐变/发光/阴影/描边/图片填充）
│   └── render_core/           # ★ PPTX→图片渲染库（依赖 fontfx；含捆绑字体 resources/fonts/）
├── src/
│   ├── gui/                   # PySide6 界面（含取消生成）
│   ├── ppt_generator.py       # 批量生成编排（恢复式复用 + 容错 + 降级）
│   ├── core/                  # 处理器（快速保存、COM 分层导出等）
│   ├── data_reader.py         # pandas 数据读取
│   └── exact_matcher.py       # 占位符匹配
├── android/                   # ★ Android 工程
│   ├── app/src/main/java/     # Kotlin/Compose UI（screens + AppState）
│   ├── app/src/main/python/   # poster_bridge 桥接 + render_core/fontfx 副本
│   └── e2e_batch.py           # 模拟器自动化测试
├── samples/                   # 业务样例（gitignored）：templates/ data/ presets/
├── resources/                 # 图标等（字体已随 render_core 库）
├── tests/                     # 应用独立测试脚本 + run_all.py 全量回归
├── tools/                     # sync_render_core.py（安卓副本同步）/ render_regression.py（像素回归）
└── docs/                      # 文档与截图（handoffs/ packaging/ 等）
```

## 测试

无测试框架配置，全部为独立脚本：

```bash
.venv\Scripts\python.exe tests\run_all.py                    # 全量：应用 + 两个库
.venv\Scripts\python.exe libs\render_core\tests\run_tests.py # 渲染库（12 个）
.venv\Scripts\python.exe libs\fontfx\tests\run_tests.py      # 特效库（4 个）
.venv\Scripts\python.exe tests\unit\test_template_restore.py # 恢复式批量等价性
# 改渲染引擎后跑像素回归：tools/render_regression.py
```

## 版本历史

### v1.1.0（2026-08-19）
- 批量容错（双端）：单行数据失败跳过并汇报行号/原因，可只重试失败行
- 预检结果上界面（双端）：模板近似渲染效果/字体回退提示直接展示给用户
- 桌面易用性：模板/数据/输出"最近使用"（MRU 8 条）一键选择
- 安卓：内置示例模板（assets/builtin_templates/ 出厂预置机制）
- 安卓正式分发：release 签名（keystore 离线管理）+ 单 ABI 77MB +
  版本对齐 1.1.0 + APK 命名 PosterAgain-<version>-<type>.apk
- 工程化：render_core 双副本同步脚本、tests/run_all.py 全量回归、GitHub Actions CI

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
