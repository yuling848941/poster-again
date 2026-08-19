## Context

- 排版校准常数（LINE_HEIGHT_RATIO/BASELINE0_FRAC/STRIKE_POS_EM）是**字体绑定**的：现值仅对微软雅黑经 PowerPoint COM 逐像素校准（残余偏差 ±4.5 逻辑像素）。换字体必须用 `experiments/calibrate.py` 重校（只能在装有 PowerPoint 的 Windows 机上做）。
- 现有 6 套渲染测试共 61 项像素级断言，全部基于微软雅黑渲染校准。任何默认字体切换都会造成大面积像素回归。
- 微软雅黑不可合法内嵌分发；思源黑体（SIL OFL 1.1）可，且与 Android 系统 CJK 字体 Noto Sans CJK 同源，是三端一致性目标下的自然选择。

## Goals / Non-Goals

- Goals：字体解析跨平台化；每物理字体绑定校准常数；Windows 渲染零回归；安卓引擎核心（renderer + registry + resources/fonts）自洽可带走
- Non-Goals：思源常数的 COM 重校准（后续单独做，见 tasks 4.1）；非 Windows 平台的实际移植；Latin 字体渲染改进（维持现状：未知字体回退默认）

## Decisions

- **默认策略 auto = Windows 系统优先 / 其他平台捆绑优先**，而非全局 bundled_first。
  理由：61 项像素断言基于雅黑校准，默认 bundled_first 会在思源重校准完成前引入大面积回归。最终"三端一致"需在校准完成后翻转默认值（届时为独立变更）。
- **校准常数挂在物理字体（按文件路径）上**。同一 typeface（微软雅黑）在不同策略下解析到不同物理文件（msyh.ttc vs SourceHanSansCN.otf），度量随实际渲染字体走。
- **首行基线仿射模型**（2026-08-18 校准时发现）：PowerPoint 首行基线对 lnSpc 非纯比例——同一组 base0 无法同时拟合 L=1.0 单行与 L=1.2 多行块（两者需求差 ~0.05em）。改为 baseline = (a + g×lnSpc)×字号px，5 个测量点（3 种行距 × 3 种字号）最小二乘残差 ≤2px@q2。微软雅黑以 a=0/g=0.9 等价改写（0.72×1.25×L×F ≡ 0.9L×F），Windows 渲染零变化（回归测试含 COM 交叉验证确认）。
- **思源黑体/宋体共享垂直度量常数**：两字体 COM 对照校准结果逐项一致（Adobe 同族设计），字形差异已单独核实（像素级比对 4189 差异像素）。共用一组常数 (1.1942, 0.1852, 0.7803)。
- **校准方法论沉淀在 `experiments/calibrate_sourcehan.py`**（多行距分块 + 仿射最小二乘拟合 + 残差报告），替代旧 calibrate.py 的手工调参流程；前置条件：字体同时存在于 resources/fonts/（渲染端）与本机用户字体（COM 端）。
- 思源常数暂沿用雅黑值 + `calibrated=False` 标记 + 一次性日志告警（已在 4.1 完成校准后解除）。
- 捆绑字体放 `resources/fonts/`：PyInstaller spec 已 `datas=[('resources','resources')]`，frozen 路径 `sys._MEIPASS/resources/fonts` 自动生效。
- 引擎核心不读配置文件：策略由应用层（`main.py` 启动时）或环境变量 `POSTERAGAIN_FONT_POLICY` 注入，注册表自身保持无配置依赖（安卓带走时零改动）。

## Risks / Trade-offs

- 思源未校准期间若被使用（非 Windows 或显式 bundled_first），多行文本行高有细微偏差 → 已有告警；校准工具就绪（calibrate.py）
- 仿宋/楷体等无捆绑等价物的字体，在非 Windows 平台回退思源黑体，字形观感变化 → 预检 FONT_MISSING 提示保留；属可接受的替代渲染
- 仓库体积 +约 30MB（4 个 OTF）→ 私有工具仓库，可接受；后续若介意可转 Git LFS

## Migration Plan

Windows 默认行为不变，无迁移。回滚 = revert 本变更即可（单一变更，无数据格式改动）。

## Open Questions

- 思源 SC 子集 4 个字重（Regular/Bold × 黑体/宋体）是否够覆盖模板需求？—— 用真实模板跑预检后确认，不够再加 Light/Medium。
