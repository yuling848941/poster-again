## ADDED Requirements

### Requirement: 安卓壳应用（Chaquopy 嵌渲染核心）
仓库 SHALL 包含 `android/` Gradle 工程（Kotlin + Jetpack Compose + Chaquopy），将 render_core 渲染核心与捆绑思源字体嵌入 APK，提供从应用 assets 读取模板并渲染为 PNG 展示的端到端路径。

#### Scenario: 端到端渲染
- **WHEN** 在安卓设备/模拟器上点击渲染按钮
- **THEN** 嵌入的 Python 渲染核心以捆绑字体（bundled_first）将 assets 内 .pptx 渲染为 PNG 并显示于界面

#### Scenario: 构建可复现
- **WHEN** 使用 AGP 8.2.2 + Gradle 8.6 + JDK 17 + Chaquopy 16.1.0（Python 3.12）执行 assembleDebug
- **THEN** 产出可安装的 debug APK（python-pptx/Pillow/lxml 经 Chaquopy 预编译仓解析）
