## ADDED Requirements

### Requirement: 自包含渲染核心包 render_core
渲染引擎 SHALL 以 `src/render_core/` 自包含包形式提供（pillow_image_exporter + render_precheck + format_inheritance + shape_geometry + font_registry），对外暴露 `render(pptx|bytes, output_dir, image_format, quality, precheck) -> RenderResult(images, issues)` 一句话 API；包依赖闭包 MUST 限于 {标准库, python-pptx(→lxml), Pillow, numpy}，不依赖 GUI/pandas/配置文件。旧导入路径（src.core.processors.*）MUST 经转发 shim 继续可用且行为不变。

#### Scenario: 安卓模拟自包含
- **WHEN** 将 `src/render_core/` + 字体目录拷贝到独立位置，在无仓库依赖的干净子进程中以 bundled_first 策略调用 render()
- **THEN** 渲染成功且输出与仓库内渲染逐像素一致

#### Scenario: 旧路径零回归
- **WHEN** GUI/测试/脚本经 `src.core.processors.*` 旧路径导入渲染引擎
- **THEN** 获得与迁移前相同的模块对象与行为，全部既有测试通过

### Requirement: 内存受限渲染路径
图片解码路径 SHALL 避免不必要的全分辨率中间副本：JPEG 在无需裁剪时 MUST 于解码期降采样（draft）；不透明图片 MUST 先缩放再补 alpha 且与"先 convert 后 resize"逐像素等价；srcRect 裁剪 MUST 在缩放前于源图坐标执行。

#### Scenario: 像素等价
- **WHEN** 以 RGB/RGBA/调色板透明/srcRect 各场景经新路径加载
- **THEN** 输出与旧路径逐像素一致

#### Scenario: 峰值内存受控
- **WHEN** 渲染含 4000×6000 背景图的模板
- **THEN** JPEG 背景峰值显著低于 PNG（draft 生效），PNG 峰值不含旧路径的全分辨率 RGBA 副本
