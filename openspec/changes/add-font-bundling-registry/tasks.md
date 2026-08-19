## 1. 字体资产
- [x] 1.1 下载思源黑体/思源宋体 SC 子集 Regular+Bold 到 `resources/fonts/`，验证 PIL 可加载、CJK 字形正常
## 2. 注册表与接线
- [x] 2.1 实现 `src/core/processors/font_registry.py`（来源解析 + 策略 + 每字体校准常数 + is_resolvable/available_typefaces）
- [x] 2.2 `pillow_image_exporter` 接线：`_font_path_for` 改走注册表；三个校准常数按物理字体取值；预检 installed 集并入注册表
- [x] 2.3 `font_checker.check_template_fonts` 并入注册表判定（可解析即不算缺失）
- [x] 2.4 `config/config.yaml` 增加 `image_generation.font_policy`，`main.py` 启动时注入注册表
## 3. 验证
- [x] 3.1 新增 `tests/unit/test_font_registry.py`（standalone 脚本，覆盖策略顺序/回退链/常数绑定/捆绑目录解析）
- [x] 3.2 6 套渲染测试 + `test_pptx_processor` + `test_complete_workflow` 全绿（Windows 零回归）
## 4. 后续（独立变更，不阻塞本变更）
- [x] 4.1 思源常数 COM 校准：字体安装到本机 → `experiments/calibrate_sourcehan.py` 重校 → 替换 METRICS_SOURCE_HAN 并置 calibrated=True
  （2026-08-18 完成：黑体/宋体 (1.1942, 0.1852, 0.7803)，仿射基线模型，残差 ≤2px@q2；
  过程中修复 resolve() 来源排序 bug——system_first 下雅黑默认链曾短路捆绑直接命中）
- [x] 4.2 校准完成后评估翻转默认策略为 bundled_first（三端渲染一致）
  （2026-08-19 评估结论：维持 auto 默认，不翻转。理由：Windows 桌面端
  auto=系统字体优先，与本机 PowerPoint 打开模板的视觉一致（用户日常对照
  基准）；翻转 bundled_first 会让桌面端与 PowerPoint 出现字体差异（安卓
  无雅黑被迫用思源已是特例）。三端一致性诉求场景（如 PC 与手机同模板必须
  逐像素一致）已有一次性方案：config 设 font_policy: bundled_first，
  一行配置可切，无需改默认值）
