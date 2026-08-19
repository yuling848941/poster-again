## Why

安卓路线 A（Kotlin/Compose + Chaquopy 嵌 Python 渲染引擎）已批准（2026-08-18）。
Phase 0（字体打包化 + COM 校准）完成后，引擎需收拢为自包含包并完成内存适配，
才能真正嵌入安卓 App 端到端运行。本变更覆盖 Phase 1~3 全部工作。

## What Changes

- **Phase 1 引擎独立成包**：渲染五模块（pillow_image_exporter / format_inheritance /
  shape_geometry / render_precheck / font_registry）迁入 `src/render_core/`，
  新增一句话渲染 API `render(pptx|bytes, out_dir, ...) -> RenderResult(images, issues)`；
  `src/core/processors/` 留 sys.modules 转发 shim，全部旧导入路径不变（零回归）
- **Phase 2 内存受限渲染路径**：新增 `_load_rgba`（JPEG draft 解码期降采样、
  不透明图先缩放后补 alpha、大 blob 临时文件解码），消除旧路径的全分辨率 RGBA
  中间副本；实测 4000×6000：PNG 384MB（旧 ~480MB）、JPEG 143MB（draft 生效）
- **Phase 3 安卓壳**：`android/` 新工程（AGP 8.2.2 + Kotlin 1.9.22 + Compose +
  Chaquopy 16.1.0 / Python 3.12），嵌入 render_core 副本 + 捆绑字体（包内
  resources/fonts 同级布局），largeHeap，按钮触发 assets 模板 → PNG → Compose 展示

## Impact

- Affected specs: image-generation（渲染核心与内存路径）、新增 android-app 能力
- Affected code: `src/render_core/`（新包）、`src/core/processors/*.py`（5 个 shim）、
  `tests/unit/test_render_core.py`、`tests/unit/test_renderer_memory.py`（新增）、
  `android/`（新工程）
- 依赖闭包不变：{stdlib, python-pptx(→lxml), Pillow, numpy(仅渐变)}
