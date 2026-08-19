## Why

渲染引擎的字体解析硬编码 `C:\Windows\Fonts`（`pillow_image_exporter._FONT_FILES`），是 macOS/Android 移植的最大阻塞项，也是 FONT_MISSING 误报的来源。安卓路线 A（Kotlin/Compose + Chaquopy 嵌引擎核心）已定（2026-08-18 会话批准），引擎核心必须携带可合法分发的中文字体。

## What Changes

- 新增字体注册表 `src/core/processors/font_registry.py`：typeface → 字体文件 + **每物理字体的排版校准常数**；策略 `auto`（Windows 系统优先=零回归，其他平台捆绑优先）/ `bundled_first` / `system_first`
- 捆绑思源黑体/思源宋体 SC 子集（SIL OFL 1.1，可合法内嵌分发；微软雅黑为专有字体不可内嵌）到 `resources/fonts/`；`PosterAgain.spec` 已整包 `resources`，无需改动
- `pillow_image_exporter._font_path_for` 改走注册表；`LINE_HEIGHT_RATIO`/`BASELINE0_FRAC`/`STRIKE_POS_EM` 从注册表按解析到的物理字体取值（雅黑沿用 COM 实测原值，思源待校准）
- `font_checker` 与渲染预检的字体可用性判定并入注册表，消除捆绑字体可覆盖时的 FONT_MISSING 误报
- 新增配置 `image_generation.font_policy`

## Impact

- Affected specs: image-generation
- Affected code: `src/core/processors/pillow_image_exporter.py`、新增 `src/core/processors/font_registry.py`、`src/core/utils/font_checker.py`、`config/config.yaml`、`main.py`（策略接线）、新增 `tests/unit/test_font_registry.py`
- 用户已批准（2026-08-18 会话：路线 A + Phase 0 启动）
