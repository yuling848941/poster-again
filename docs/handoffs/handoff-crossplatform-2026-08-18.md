# 交接文档：PosterAgain 安卓端 / macOS 端开发

> 交接日期：2026-08-18
> 上一任：ZCode 会话（2026-08-17 完成纯 Python 渲染引擎并合并 master）
> 交接范围：macOS 移植与 Android 方案落地。Windows 版已稳定，不在交接范围内。

---

## 一、项目现状（30 秒版）

- **仓库**：`E:\poster code\poster-again`，master @ `9cbabb7`（2026-08-17 合并）
- **是什么**：PySide6 桌面 GUI。读 Excel/CSV/JSON → 按模板批量生成个性化 .pptx + PNG 海报（保险贺报场景）
- **关键背景**：图片导出**曾经依赖 Office/WPS 的 COM 接口**，2026-08 已重写为纯 Python 渲染引擎（Pillow），COM 彻底解耦（代码休眠保留）。**这是跨平台可行性的根基——先确认你理解这一点再动手**
- 测试：6 套独立脚本共 61 项像素级断言，全过。无 pytest，直接 `python tests\unit\test_xxx.py` 运行
- 仓库约定、模块边界、常用命令见仓库根 `AGENTS.md`；验收背景见 `验收指南.md`

## 二、核心资产地图（先读这些文件）

| 文件 | 作用 |
|---|---|
| `src/core/processors/pillow_image_exporter.py` | 渲染引擎主体（文字蒙版管线：阴影→描边→纯色/渐变/图片填充；三层渲染 母版→版式→幻灯片；`_Xform` 组合坐标变换） |
| `src/core/processors/format_inheritance.py` | OOXML 格式继承链解析（run→形状→版式占位符→母版→txStyles），含主题数据 ThemeData/bgRef |
| `src/core/processors/shape_geometry.py` | 18 种预设形状几何（`SUPPORTED_PRST` 白名单与预检同源） |
| `src/core/processors/render_precheck.py` | 白名单预检（不支持元素稳定问题码 + 日志告警） |
| `experiments/calibrate.py` | **排版校准工具**：合成测试 PPT → COM 导参考 → 逐行测差值（移植前必须用它校准字体） |
| `tests/unit/test_renderer_*.py` | 6 套测试（precheck/features/shapes/transforms/text_decorations/robustness） |

## 三、跨平台评估结论（上一任已完成，直接采用）

### macOS：可行性高，预估 3~5 天

技术栈（PySide6/python-pptx/Pillow/numpy/pandas）全部有官方 macOS 支持。要做三件事：

1. **字体系统跨平台化（最大项，1~2 天）**：当前硬编码 `C:\Windows\Fonts\msyh.ttc`，`font_checker.py` 用 winreg。**推荐方案：随应用打包思源黑体/宋体（SIL OFL 可合法内嵌；微软雅黑不可内嵌）**——注意 Android 系统 CJK 字体 Noto Sans CJK 与思源同源，打包思源可三端渲染一致，还顺带解决现有 FONT_MISSING 告警
2. 平台代码清理（winreg 分支、路径），约 1 天
3. PyInstaller 打 .app + 签名公证（需 Apple Developer $99/年 + 一台 Mac），约 0.5~1 天

**每个新字体必须用 `experiments/calibrate.py` 重新校准**（见第四节）。

### Android：三堵墙，先做路线决策

1. **PySide6 无 Android 生产支持** → 1500 行 `main_window.py` 必须重写（Kotlin/Compose 或跨端框架）
2. **Python 引擎嵌入**：Chaquopy 可嵌（numpy/pandas/Pillow 有预编译包）；CPython 3.13+ 官方 Android 支持仍是实验性
3. **内存约束**：每日贺报 47MB 背景图解码后 231MB RGBA + 全画布图层，中端手机会 OOM——需要 draft 降采样/分块渲染改造

**路线决策（交接后第一件事）**：
- A. 原生 UI + Chaquopy 嵌 Python 引擎（3~6 周，双代码库）
- B. 服务端 + H5/微信端（1~2 周，但要先过合规：保险客户数据能否出内网）
- C. 等 PySide6 移动端成熟（不建议）
- **前置产品问题**：移动端到底谁用、什么场合用？如果只是查看/转发海报，可能根本不需要客户端

## 四、必须带走的隐性知识（不在任何文档里）

1. **排版校准常数是字体绑定的**：`pillow_image_exporter.py` 顶部 `LINE_HEIGHT_RATIO=1.25`、`BASELINE0_FRAC=0.72` 只对**微软雅黑**经 COM 逐像素校准（残余偏差 ±4.5 逻辑像素）。`STRIKE_POS_EM=0.245` 同理。换字体（含思源）必须重新校准，否则多行文本行高漂移
2. **COM 交叉验证方法论**（只能在这台 Windows 机器做，移植前抓紧用）：合成测试 PPT → 休眠的 `COMImageExporter` 导出 PowerPoint 参考图 → `ink_bbox` 墨迹对比逐项校准。本机装有 PowerPoint，61 项测试里的"COM 交叉验证"部分在 mac/linux 上会自动跳过
3. **测试脚本的坑（真实踩过）**：测量区域会被相邻元素污染（区域重叠/画布裁剪）；测试函数内变量遮蔽（曾因 `gx` 被重赋值误判渲染不确定性）。写像素断言时区域要精确、变量作用域要干净
4. **引擎已支持/不支持的完整边界**：见 `render_precheck.py` 模块文档字符串（问题码一览 + 已退役清单）。表格/图表/SmartArt/竖排文字等仍不支持——Android/macOS 移植**不扩大范围**，预检告警机制会自动兜住
5. **未做的已知决策**：COM 兜底链（预检发现"内容缺失类"问题时自动回退 COM）用户已知悉但选择暂不做；表格支持讨论过数据形态设计（固定表 `{{占位符}}` 逐格 / 动态明细表+明细sheet）后搁置，设计结论见会话记录，未写入仓库文档

## 五、交接后建议顺序

1. （可选但强烈推荐，对 Windows 版也纯赚）**字体打包化改造**：引入思源字体内嵌 + 字体注册表抽象，在 Windows master 上完成并用 calibrate.py 校准
2. macOS 移植（按第三节三件事）
3. Android：先回答产品问题 → 选路线 → 做
4. 任何时候拿真实模板验证：跑 GUI 或 headless（`pillow_image_exporter` 可独立调用），看 `[预检]` 日志

## 六、环境速查

- 开发机：Windows，仓库在 `E:\poster code\poster-again`，虚拟环境 `.venv`（激活方式见 `AGENTS.md`）
- 打包：`.venv\Scripts\python.exe -m PyInstaller PosterAgain.spec --noconfirm`（已验证 73MB exe）
- 主仓库有一个 `git stash`（"merge前WIP备份"）和残留 `..\poster-again-purepy\dist\` 目录，确认稳定后可清理

## 七、建议技能（Suggested skills）

接手后按需调用：
- `android-emulator:android-dev` —— Android 侧起模拟器/构建验证
- `architecture` / `codebase-design` —— Android 路线 A 的混合架构设计
- `documentation` —— 沉淀移植过程中的架构决策
- `testing-strategy` —— macOS 上无法 COM 对照后的测试策略调整
- `debug` / `diagnosing-bugs` —— 渲染差异排查（配合 ink_bbox 方法论）
