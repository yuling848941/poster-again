# AGENTS.md — PosterAgain

## Quick Commands

**重要：所有命令都应通过 `.venv` 虚拟环境运行**，避免系统 Python 里无关库（如 PyTorch）干扰。两种方式：
- 激活后运行：`.venv\Scripts\activate`（之后可直接用 `python`）
- 或显式调用：`.venv\Scripts\python.exe ...`

```bash
# 首次设置虚拟环境（已存在则跳过）
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements.txt   # 含两个本地库的可编辑安装（libs/fontfx、libs/render_core）

# 运行程序 / 测试 / 打包（统一用 .venv）
.venv/Scripts/python.exe main.py                       # Run GUI app
.venv/Scripts/python.exe tests\unit\test_xxx.py        # Run a single test (standalone script)
.venv/Scripts/python.exe tests\run_all.py              # 全量回归：应用用例 + 两个库的测试
.venv/Scripts/python.exe -m PyInstaller PosterAgain.spec --noconfirm  # Build exe via spec
打包.bat                                                # Build exe (Windows batch wrapper)
跑测试.bat                                              # 全量回归（tests/run_all.py 的批处理封装）

# 改完 libs/ 下的库代码后，同步安卓内嵌副本（CI 用 --check 防漂移）
.venv/Scripts/python.exe tools/sync_render_core.py
.venv/Scripts/python.exe tools/sync_render_core.py --check

# 渲染像素回归（改引擎前后各采一次，逐像素确认输出没变）
.venv/Scripts/python.exe tools/render_regression.py corpus --out before.jsonl "samples/templates/xxx.pptx"
.venv/Scripts/python.exe tools/render_regression.py verify before.jsonl after.jsonl

# 安卓端（android/ 工程；SDK 在 %LOCALAPPDATA%\Android\Sdk，Gradle 8.6 在 %LOCALAPPDATA%\gradle-8.6）
cd android && "$LOCALAPPDATA/gradle-8.6/bin/gradle.bat" assembleDebug   # Build debug APK
.venv/Scripts/python.exe android/e2e_batch.py                         # 模拟器全流程 E2E（选模板→数据→字段绑定→渲染→预览；需已启动模拟器，会 pm clear 应用）
```

⚠️ **打包必须用 `.venv`**：系统 Python 装了 PyTorch 等大量无关库，用系统 Python 打包会让 PyInstaller 误卷入 7GB+ 垃圾库，导致 exe 体积爆炸或打包失败。

应用测试是独立脚本（无 pytest 配置），直接用 `python` 跑。两个库（`libs/fontfx`、`libs/render_core`）各有 `pyproject.toml`，以可编辑方式装进 `.venv`，测试也是独立脚本。

## Architecture

**What it does:** Desktop GUI app (PySide6) that reads Excel/CSV/JSON data and batch-generates personalized .pptx files from a template.

**Current version: 1.1.0.** PPT generation is pure Python (`python-pptx`). Image export defaults to `python` (the built-in Pillow renderer `render_core`, fast and calibrated for the 贺报 template family); `office` (COM high-fidelity) and `auto` (COM when Office present) are selectable via the GUI engine dropdown / runtime config `processing.image_export_engine` / `POSTERAGAIN_IMAGE_ENGINE`. Engine config is read from the **runtime** ConfigManager (works in the frozen exe too). The COM exporter (`com_image_exporter.py`) also serves pixel-level cross-validation in tests and calibration on machines that have PowerPoint.

**Two-layer rendering libraries (under `libs/`, pip-installable, editable-installed into `.venv`):**
- `libs/fontfx` (`fontfx`) — pure-Pillow text effects: linear gradient / image fill / glow / shadow / stroke / decorations / synthetic italic. No PPTX dependency; usable by any Pillow project.
- `libs/render_core` (`render_core`, distribution name `pptx-render-core`) — PPTX→image engine: format inheritance, layout, precheck, shape geometry, font registry, batch personalization. Delegates "how a glyph looks" to `fontfx`. Bundled Source Han fonts live in `render_core/resources/fonts/` (SIL OFL) + per-font calibrated layout metrics in `font_registry.py`.

**Cross-platform note:** the Android app (`android/`, Kotlin/Compose + Chaquopy) embeds copies of both packages under `app/src/main/python/` (pip unavailable there). Font policy: `auto` (Windows system-first, other platforms bundled-first), override via `image_generation.font_policy` or `POSTERAGAIN_FONT_POLICY`. Image cache budget via `POSTERAGAIN_PIC_CACHE_MB` (default 256; Android bridge sets 64).

### Module boundaries

| Path | Role |
|---|---|
| `main.py` | Entry point — inits `ConfigManager`, sets up file logging + excepthooks, launches `MainWindow` |
| `src/gui/main_window.py` | PySide6 main window. UI logic, file selection, matching, generation orchestration, cancel |
| `src/gui/ppt_worker_thread.py` | `QThread` for non-blocking generation. Signals: progress, log, match_completed, generation_completed, error; cooperative cancel |
| `src/ppt_generator.py` | Coordinates `PPTXProcessor` + `DataReader` + `ExactMatcher` for batch generation (per-row replace→export→restore, fault tolerance, degraded rows) |
| `src/core/processors/pptx_processor.py` | Template processor: placeholder replace, fast save, in-memory render / COM export, engine caching |
| `src/core/utils/pptx_fast_save.py` | Fast .pptx save (already-compressed media stored, not re-deflated); self-verifies against standard save |
| `libs/render_core/` | **Render library** (PPTX→image). Public API: `render(...)`, `Personalizer`, `PillowImageExporter`, `RenderPrecheck`, `font_registry` |
| `libs/fontfx/` | **Text-effects library** (pure Pillow). Public API: `draw_text`, `text_image`, `draw_glyphs`, `linear_gradient`, style dataclasses |
| `src/data_reader.py` | Reads Excel/CSV/JSON into pandas DataFrames |
| `src/exact_matcher.py` | Matches PPT placeholders (`ph_字段名`) to Excel column names |
| `src/memory_management/` | In-memory data caching and optimization |
| `config_manager.py` | Root-level compat shim — delegates to `src/config/` |
| `config/config.yaml` | App config defaults. Loaded by `ConfigManager` |
| `samples/` | User business files (gitignored): `templates/` (.pptx), `data/` (.xlsx/.csv), `presets/` (.pptcfg) |
| `android/` | Android shell app (Kotlin + Compose + Chaquopy 16.1.0 / Python 3.12) embedding copies of `render_core` + `fontfx` + bundled fonts under `app/src/main/python/`. UI 为多屏 Compose（`AppState.kt` + `ui/screens/` 六屏）；桥接层 `poster_bridge.py`（PC 侧单测 `tests/unit/test_poster_bridge.py`，模拟器 E2E `android/e2e_batch.py`） |

### Import pattern

Most modules use a dual try/except import to support both `src.`-prefixed and flat imports:
```python
try:
    from src.ppt_generator import PPTGenerator
except ImportError:
    from ppt_generator import PPTGenerator
```
This supports running from root and from inside `src/`. Keep this pattern when adding modules.

## Placeholder Convention

PPT templates use **one of three** placeholder identification methods (checked in order):
1. **Alt Text** (替代文本) — `ph_字段名` — recommended, invisible to users
2. **Shape Name** — `ph_字段名` — set via PowerPoint Selection Pane
3. **Text markers** — `{{字段名}}` — inline text replacement

The `ph_` prefix is defined in `PPTXProcessor.PLACEHOLDER_PREFIX`. Field names must exactly match Excel column headers (case-insensitive matching is configurable).

## Testing

- 52 test files in `tests/unit/`, all standalone scripts; plus 4 in `libs/fontfx/tests/` and 12 in `libs/render_core/tests/`
- Each test adds the needed roots to `sys.path` itself, so they run directly from the repo root: `.venv\Scripts\python.exe tests\unit\test_xxx.py`
- No test framework config — just `python tests\unit\test_xxx.py`
- **全量回归**：`.venv\Scripts\python.exe tests\run_all.py`（先跑应用用例，再跑两个库的测试；`--app-only` 只跑应用）。CI 用它 + `tools/sync_render_core.py --check` + wheel 构建
- Many tests are historical "fix verification" scripts (e.g. `test_config_sync_fix.py`, `test_placeholder_fix.py`)
- The most relevant for new work: `test_pptx_processor.py`, `test_template_restore.py`（恢复式批量等价性）, `test_batch_fault_tolerance.py`（容错/降级/取消）, `libs/render_core/tests/`（渲染）
- **改渲染引擎后务必跑像素回归**：`tools/render_regression.py`（corpus 采样 / verify 对比），确认输出逐像素不变
- **Always run the specific test file related to your change**

## OpenSpec Workflow

This repo uses OpenSpec for spec-driven development. See `openspec/AGENTS.md` for full details.

**When to use:** New features, breaking changes, architecture shifts, performance optimization.
**When to skip:** Bug fixes, typos, formatting, dependency updates, tests for existing behavior.

Before any task: run `openspec list` and `openspec list --specs` to check for existing work and conflicts.

## Build / Distribution

- PyInstaller spec: `PosterAgain.spec`（`pathex` 指向 `libs/*/src`，`datas` 含 `resources/` 与库内捆绑字体 `render_core/resources/fonts`）
- Build script: `build_exe.py` (cleans old dist/build, runs pyinstaller)
- Windows shortcut: `打包.bat`
- Output: `dist/PosterAgain.exe`
- Hidden imports in spec include: `yaml`, `pandas`, `pptx`, `numpy`, `fontTools`, `win32com`, `pythoncom`, `pywintypes`
- Icon: `resources/icon.ico`

## Gotchas

- `main.py` creates a module-level `config_manager` singleton at import time — importing `main` has side effects
- `ConfigManager()` called without args defaults to `config/config.yaml`
- `config/config.yaml` (the tracked template) holds defaults. A runtime snapshot is written to the **repo-root `config.yaml`** (or `PosterAgain_config.yaml`) containing absolute paths from the last user session — these root-level snapshots are gitignored and must not be committed
- The `office_manager` parameter on `PPTWorkerThread` is deprecated (always set to `None`) — kept for backward compat
- **Logging:** `main.setup_logging()` calls `logging.basicConfig` once at startup — level from runtime config `advanced.log_level` (default INFO), outputs to console **and** a rotating file (`%LOCALAPPDATA%/PosterAgain/logs/posteragain.log`, size/count from `advanced.max_log_size`/`log_file_count`), and installs `sys.excepthook` + `threading.excepthook`. Every other module should use `logging.getLogger(__name__)` and never call `basicConfig` itself
- `ppt_worker_thread.py` no longer swallows exceptions silently — chengbao-term processing errors are now logged at WARNING level instead of `except: pass`
- **渲染库在两处**：`libs/render_core/src/render_core/`（桌面/可编辑安装）与 `android/app/src/main/python/render_core/`（安卓 APK 内副本，字体在包内 `resources/fonts/`）；`fontfx` 同理。改完库代码跑 `tools/sync_render_core.py` 同步安卓副本（CI `--check` 防漂移）
- **批量生成用「恢复式复用」**：`PPTXProcessor.restore_template()` 逐行撤销占位符替换（等价于重新加载但无需重新解包），使渲染器字体/图片缓存跨行生效。`pptx_fast_save.py` 对已压缩媒体不再重复 deflate（首次使用自动与标准保存逐成员校验）。两者都经 `test_template_restore.py` 与像素回归验证
- **排版校准常数是字体绑定的**：`font_registry.py` 按物理字体提供 `line_height_ratio + 首行基线仿射 (baseline_a, baseline_g)×lnSpc`。换字体必须用 `experiments/calibrate_sourcehan.py`（需本机 PowerPoint COM + 字体同时装入用户字体）重校，勿手调
- **安卓构建版本矩阵（实测可用）**：AGP 8.2.2 + Gradle 8.6 + JDK 17 + Chaquopy 16.1.0 + 目标 Python 3.12。Chaquopy 15.x 在构建机 Python 3.13 上会因 `cgi` 模块移除而失败
- **安卓模拟器内存**：largeHeap 已开；背景图建议 JPEG（draft 解码降采样，峰值 ~143MB vs PNG ~384MB @4000×6000）；图片缓存上限安卓桥接设为 64MB
