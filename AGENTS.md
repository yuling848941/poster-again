# AGENTS.md — PosterAgain

## Quick Commands

**重要：所有命令都应通过 `.venv` 虚拟环境运行**，避免系统 Python 里无关库（如 PyTorch）干扰。两种方式：
- 激活后运行：`.venv\Scripts\activate`（之后可直接用 `python`）
- 或显式调用：`.venv\Scripts\python.exe ...`

```bash
# 首次设置虚拟环境（已存在则跳过）
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements.txt

# 运行程序 / 测试 / 打包（统一用 .venv）
.venv/Scripts/python.exe main.py                       # Run GUI app
.venv/Scripts/python.exe tests\unit\test_xxx.py        # Run a single test (standalone script)
.venv/Scripts/python.exe -m PyInstaller PosterAgain.spec --noconfirm  # Build exe via spec
打包.bat                                                # Build exe (Windows batch wrapper)

# 安卓端（android/ 工程；SDK 在 %LOCALAPPDATA%\Android\Sdk，Gradle 8.6 在 %LOCALAPPDATA%\gradle-8.6）
cd android && "$LOCALAPPDATA/gradle-8.6/bin/gradle.bat" assembleDebug   # Build debug APK
.venv/Scripts/python.exe android/e2e_batch.py                         # 模拟器全流程 E2E（选模板→数据→字段绑定→渲染→预览；需已启动模拟器，会 pm clear 应用）
```

⚠️ **打包必须用 `.venv`**：系统 Python 装了 PyTorch 等大量无关库，用系统 Python 打包会让 PyInstaller 误卷入 7GB+ 垃圾库，导致 exe 体积爆炸或打包失败。

No pytest config, no `conftest.py`, no `pyproject.toml`. Tests are standalone Python scripts — run them directly with `python`.

## Architecture

**What it does:** Desktop GUI app (PySide6) that reads Excel/CSV/JSON data and batch-generates personalized .pptx files from a template.

**Current version: 1.1.0.** PPT generation AND image export are pure Python (`python-pptx` + Pillow renderer) — **no Microsoft Office or WPS needed**. The COM-based exporter (`com_image_exporter.py`) is dormant, kept only for pixel-level cross-validation in tests and calibration on machines that have PowerPoint.

**Cross-platform note:** the rendering engine lives in `src/render_core/` (self-contained: stdlib + python-pptx/lxml + Pillow + optional numpy for gradient text). Bundled Source Han fonts in `resources/fonts/` (SIL OFL) + per-font calibrated layout metrics in `font_registry.py`. The same package is embedded in the Android app (`android/`, Kotlin/Compose + Chaquopy). Font policy: `auto` (Windows system-first, other platforms bundled-first), override via `image_generation.font_policy` or `POSTERAGAIN_FONT_POLICY`.

### Module boundaries

| Path | Role |
|---|---|
| `main.py` | Entry point — inits `MemoryDataManager` + `ConfigManager`, then launches `MainWindow` |
| `src/gui/main_window.py` | PySide6 main window (~1500 lines). All UI logic, file selection, matching, generation orchestration |
| `src/gui/ppt_worker_thread.py` | `QThread` for non-blocking PPT generation. Signals: progress, log, match_completed, error |
| `src/ppt_generator.py` | Coordinates `PPTXProcessor` + `DataReader` + `ExactMatcher` for batch generation |
| `src/render_core/` | **Self-contained render package** (PPTX→PNG engine, precheck, format inheritance, shape geometry, font registry). Public API: `render(pptx_or_bytes, out_dir, ...) -> RenderResult`. Copy-paste embeddable (Android uses a copy) |
| `src/core/processors/*.py` | Template processor (`pptx_processor.py`) + **forwarding shims** for the render modules that moved to `src/render_core/` — old import paths keep working |
| `src/data_reader.py` | Reads Excel/CSV/JSON into pandas DataFrames |
| `src/exact_matcher.py` | Matches PPT placeholders (`ph_字段名`) to Excel column names |
| `src/memory_management/` | In-memory data caching and optimization |
| `config_manager.py` | Root-level compat shim — delegates to `src/config/` |
| `config/config.yaml` | App config (paths, matching rules, UI state). Loaded by `ConfigManager` |
| `android/` | Android shell app (Kotlin + Compose + Chaquopy 16.1.0 / Python 3.12) embedding a copy of `render_core` + bundled fonts under `app/src/main/python/`. UI 为多屏 Compose（`AppState.kt` 状态容器 + `ui/screens/` 六屏：首页向导/字段绑定/渲染/大图预览/模板/设置），主题色板来自 Stitch 设计稿（`ui/theme/Theme.kt`，深青 primary + success 扩展色）；桥接层 `app/src/main/python/poster_bridge.py`（PC 侧单测 `tests/unit/test_poster_bridge.py`，模拟器 E2E `android/e2e_batch.py`） |

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

- 47 test files in `tests/unit/`, all standalone scripts
- Each test adds the project root to `sys.path` itself (via `os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))`), so they run directly from the repo root: `.venv\Scripts\python.exe tests\unit\test_xxx.py`
- No test framework config — just `python tests\unit\test_xxx.py`
- Many tests are historical "fix verification" scripts (e.g. `test_config_sync_fix.py`, `test_placeholder_fix.py`)
- The most relevant for new work: `test_pptx_processor.py`, `test_complete_workflow.py`, `test_com_processor.py`
- **Always run the specific test file related to your change** — there is no "run all tests" target

## OpenSpec Workflow

This repo uses OpenSpec for spec-driven development. See `openspec/AGENTS.md` for full details.

**When to use:** New features, breaking changes, architecture shifts, performance optimization.
**When to skip:** Bug fixes, typos, formatting, dependency updates, tests for existing behavior.

Before any task: run `openspec list` and `openspec list --specs` to check for existing work and conflicts.

## Build / Distribution

- PyInstaller spec: `PosterAgain.spec`
- Build script: `build_exe.py` (cleans old dist/build, runs pyinstaller)
- Windows shortcut: `打包.bat`
- Output: `dist/PosterAgain.exe`
- Hidden imports in spec include: `yaml`, `pandas`, `pptx`, `win32com`, `pythoncom`, `pywintypes`
- Icon: `resources/icon.ico`

## Gotchas

- `main.py` creates module-level singletons (`memory_manager`, `config_manager`) at import time — importing `main` has side effects
- `ConfigManager()` called without args defaults to `config/config.yaml`
- `config/config.yaml` (the tracked template) holds defaults. A runtime snapshot is written to the **repo-root `config.yaml`** (or `PosterAgain_config.yaml`) containing absolute paths from the last user session — these root-level snapshots are gitignored and must not be committed
- The `office_manager` parameter on `PPTWorkerThread` is deprecated (always set to `None`) — kept for backward compat
- **Logging:** `main.setup_logging()` calls `logging.basicConfig` once at startup (reads `advanced.log_level` from `config/config.yaml`, default INFO, outputs to console). Every other module should use `logging.getLogger(__name__)` and never call `basicConfig` itself, to avoid clobbering the global config
- `ppt_worker_thread.py` no longer swallows exceptions silently — chengbao-term processing errors are now logged at WARNING level instead of `except: pass`
- **渲染引擎在两处**：`src/render_core/`（Windows/GUI 用）与 `android/app/src/main/python/render_core/`（安卓 APK 内副本，字体在包内 `resources/fonts/`）。改动引擎后需同步两份拷贝（同步脚本待做，见 openspec `add-render-core-android` tasks 4.3）
- **排版校准常数是字体绑定的**：`font_registry.py` 按物理字体提供 `line_height_ratio + 首行基线仿射 (baseline_a, baseline_g)×lnSpc`。换字体必须用 `experiments/calibrate_sourcehan.py`（需本机 PowerPoint COM + 字体同时装入用户字体）重校，勿手调
- **安卓构建版本矩阵（实测可用）**：AGP 8.2.2 + Gradle 8.6 + JDK 17 + Chaquopy 16.1.0 + 目标 Python 3.12。Chaquopy 15.x 在构建机 Python 3.13 上会因 `cgi` 模块移除而失败
- **安卓模拟器内存**：largeHeap 已开；背景图建议 JPEG（draft 解码降采样，峰值 ~143MB vs PNG ~384MB @4000×6000）
