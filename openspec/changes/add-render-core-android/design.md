## Context

- 交接结论：安卓路线 A = 原生 UI + Chaquopy 嵌引擎；引擎依赖闭包干净（无 GUI/pandas）
- python-pptx 在 `Presentation()` 时全量加载包部件（急切加载），Pillow PNG 解码无降采样——这两项构成内存下限
- Chaquopy 15.0.1 的构建工具链不兼容 Python 3.13（cgi 移除），16.1.0 + 目标 Python 3.12 实测可用

## Goals / Non-Goals

- Goals：引擎可整体拷贝嵌入（自包含）；渲染内存路径适配移动端；安卓 App 端到端跑通（模板 → PNG → 屏幕显示）
- Non-Goals：安卓 UI 完整产品化（文件选择/批量/分享）；真实模板性能调优；iOS

## Decisions

- **物理迁移而非门面**：五模块移入 `src/render_core/`，`core/processors/` 用
  `sys.modules[__name__] = _real` 转发 shim——旧导入路径（GUI/测试/experiments）零改动，
  而 `src/render_core/` 目录本身可直接拷走
- **render() 支持字节输入**：安卓场景 assets 读出即传，内部落临时文件
- **内存模型（实测归因）**：峰值 ≈ 解释器+库(~45MB) + pptx 急切加载(≈1.1×文件) +
  全分辨率解码(W×H×3)。PNG 不可再压（体系下限）；JPEG 用 draft 跳过全分辨率解码；
  旧路径的全分辨率 RGBA 副本已消除。安卓模板建议背景用 JPEG
- **Chaquopy 16.1.0 + Python 3.12 + AGP 8.2.2 + Gradle 8.6 + JDK 17**（版本组合实测通过；
  15.0.1 因构建机 Python 3.13 无 cgi 失败）
- **ABI**：arm64-v8a + x86_64；字体包内同级 `render_core/resources/fonts`（注册表已有该布局候选）
- **首个 APK 不含 numpy**（渐变文字路径暂缺，测试模板用纯色）；后续按需加（+~40MB/ABI）

## Risks / Trade-offs

- render_core 在仓库与 android/ 工程各有一份副本（同步靠拷贝）→ 后续加打包脚本或 symlink 策略
- Chaquopy pip 依赖网络（预编译仓 + PyPI），离线构建需 prebuilt 缓存
- 86MB debug APK（含双 ABI Python 运行时 + 40MB 字体）——release + 单 ABI 可减半

## Migration Plan

Windows 端零迁移（shim 保旧路径）。回滚 = revert 本变更。

## Open Questions

- 真实保险贺报模板在模拟器/真机的渲染耗时与内存（Phase 4 实测）
- 加 numpy 后 APK 体积与启动耗时
