# AGENTS.md — PosterAgain

## Quick Commands

```bash
python main.py                          # Run GUI app
python -m venv .venv && .venv\Scripts\activate  # Setup env (Windows)
pip install -r config/requirements.txt  # Install deps
python build_exe.py                     # Build exe via PyInstaller
打包.bat                                 # Build exe (Windows batch wrapper)
python tests\unit\test_xxx.py           # Run a single test (standalone script)
```

No pytest config, no `conftest.py`, no `pyproject.toml`. Tests are standalone Python scripts — run them directly with `python`.

## Architecture

**What it does:** Desktop GUI app (PySide6) that reads Excel/CSV/JSON data and batch-generates personalized .pptx files from a template.

**v2.0.0 uses pure Python** — `python-pptx` only. No Microsoft Office or WPS COM dependency. The `use_com_interface` config flag is `false` and old COM processor code is dead.

### Module boundaries

| Path | Role |
|---|---|
| `main.py` | Entry point — inits `MemoryDataManager` + `ConfigManager`, then launches `MainWindow` |
| `src/gui/main_window.py` | PySide6 main window (~1500 lines). All UI logic, file selection, matching, generation orchestration |
| `src/gui/ppt_worker_thread.py` | `QThread` for non-blocking PPT generation. Signals: progress, log, match_completed, error |
| `src/ppt_generator.py` | Coordinates `PPTXProcessor` + `DataReader` + `ExactMatcher` for batch generation |
| `src/core/processors/pptx_processor.py` | Template processor — reads `.pptx`, finds placeholders, replaces text |
| `src/data_reader.py` | Reads Excel/CSV/JSON into pandas DataFrames |
| `src/exact_matcher.py` | Matches PPT placeholders (`ph_字段名`) to Excel column names |
| `src/memory_management/` | In-memory data caching and optimization |
| `config_manager.py` | Root-level compat shim — delegates to `src/config/` |
| `config/config.yaml` | App config (paths, matching rules, UI state). Loaded by `ConfigManager` |

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

- ~49 test files in `tests/unit/`, all standalone scripts
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
- `config/config.yaml` stores **absolute paths** from the last user session — don't commit path changes
- The `office_manager` parameter on `PPTWorkerThread` is deprecated (always set to `None`) — kept for backward compat
- `logging.basicConfig` is NOT called globally; each module uses `logging.getLogger(__name__)`. If you need to configure logging, do it in `main()`
