# 交接文档：安卓端路线 A 落地（Phase 0~3 + Phase A 批量个性化完成）

> 日期：2026-08-18（接 handoff-crossplatform-2026-08-18.md）
> 状态：**批量个性化端到端已打通**——模板+CSV/XLSX 数据 → 批量替换渲染 → 画廊/相册/分享；
> 设备 vs PC 渲染差异 0.049% 像素（纯抗锯齿级）

---

## 一、已完成全景

| 阶段 | 内容 | 证据 |
|---|---|---|
| Phase 0 | 思源字体捆绑 + font_registry + COM 校准（仿射基线模型 1.1942/0.1852/0.7803，残差≤2px@q2） | openspec `add-font-bundling-registry`（全部勾选） |
| Phase 1 | 引擎五模块迁入 `src/render_core/` + 旧路径 shim + `render()` API | `test_render_core.py` 7/7（含独立目录"安卓模拟"测试） |
| Phase 2 | 内存路径优化（JPEG draft / 先缩放后补 alpha / 临时文件解码） | `test_renderer_memory.py` 9/9；PNG 384MB（旧~480）JPEG 143MB @4000×6000 |
| Phase 3 | android/ 工程（Compose + Chaquopy 16.1.0/Py3.12）→ 86MB APK → 模拟器端到端 | 截图 + logcat；设备/PC 渲染版式逐像素一致（0.4% 抗锯齿差异） |

openspec：`add-render-core-android`（任务 1~3 全勾，4.x 为后续）。

## 二、安卓工程速查

- 位置 `android/`；构建：`cd android && "$LOCALAPPDATA/gradle-8.6/bin/gradle.bat" assembleDebug`
- 版本矩阵（实测）：**AGP 8.2.2 + Gradle 8.6 + JDK 17 + Chaquopy 16.1.0 + 目标 Python 3.12**。Chaquopy 15.x 在构建机 Python 3.13 下挂（cgi 移除）
- SDK：`%LOCALAPPDATA%\Android\Sdk`（cmdline-tools 在 `latest/`，licenses 已写入）
- AVD：`poster34`（Pixel 6 / API 34 / x86_64）。注意 avdmanager 曾把 SDK 根认错，`config.ini` 的 `image.sysdir.1` 已手修，重建 AVD 时留意
- 起模拟器：`emulator -avd poster34 -no-window -gpu swiftshader_indirect`（需同时 export ANDROID_HOME 与 ANDROID_SDK_ROOT）
- adb 走 Git Bash 时路径要 `MSYS_NO_PATHCONV=1`，否则 `/sdcard` 被转成本地路径

## 三、本轮沉淀的关键知识（不在任何官方文档）

0. **行距是 typographic 模型（2026-08-18 承保贺报重叠 bug 修复时发现）**：
   PowerPoint 相邻行推进 = descent(上一行) + ascent(下一行)，其中
   ascent=(a+g×lnSpc)×字号px、descent=行高−ascent。同字号时 ≡ 行高（历史校准
   全部自洽）；**字号骤增的相邻行**（如 36pt 行下一行 66pt）按旧行高推进会少一
   个 ascent、大字上叠（承保贺报"廖云瑞叠朱金英家族"即此）。COM 对照实测带位
   对齐 ±5px@q2。空段落占一个整行盒（ascent=0/descent=全行高）。

1. **Chaquopy 16 的 app.imy 打包**：python 源码打进 `assets/chaquopy/app.imy`，`.py` 由导入器直接服务（`__file__` 是真实形态路径但 .py 不在磁盘），**数据文件解压到** `files/chaquopy/AssetFinder/app/…`。捆绑字体必须在包内同级 `render_core/resources/fonts`（`bundled_font_dir` 已有该候选；曾因改成"只向上探测"丢同级候选而在设备上找不到字体——repo 测试没暴露，因为仓库布局靠上两级命中）
2. **python-pptx 急切加载**：`Presentation()` 一次性读入全部部件——内存下限 ≈ 1.1×pptx 文件大小，再 + 全分辨率解码（W×H×3）。安卓模板背景**建议 JPEG**（draft 解码降采样，143MB vs PNG 384MB @4000×6000）
3. **跨平台一致性边界**：设备（Chaquopy Pillow/FreeType）与 PC（Pillow 12）渲染同模板：版式逐像素一致，0.4% 像素有抗锯齿级差异（平均 5.5/255）——属光栅化器差异，非版式漂移
4. **Chaquopy 模拟器实测性能**：720×960 海报（含预检）745ms @x86_64 模拟器；真机 ARM 需另测

## 四、真实模板 + 真机实测（2026-08-18 补充）

- **真机（用户手机）**：内置测试模板渲染 **623ms/张**（模拟器 745ms）
- **真机 + 真实模板（50MB 每日贺报，3584×1792）**：**3445ms/张**——业务可接受（现场生成单张场景）；
  PC 端 2.5s（含导入），真机仅慢 ~40%，Python on ARM 表现良好
- **承保贺报（566KB，含文字渐变）**：最初 APK 缺 numpy 报 ModuleNotFoundError；
  修复后设备渲染成功，与 PC 差异 0.31% 抗锯齿级，预检零问题
- numpy 1.26.2 已入包（APK 121→141MB）；引擎渐变渲染新增纯 Pillow 兜底
  （无 numpy 自动退化，与 numpy 路径差异 ≤1/255，四角度实测）——numpy 永远优先
- **代理环境下的 Chaquopy 构建套路（重要，极易踩）**：
  - pip 走 SOCKS 需要 PySocks：`no_proxy='*' ./app/build/python/env/debug/Scripts/python.exe -m pip install PySocks -i https://mirrors.aliyun.com/pypi/simple`（阿里云可直连）
  - 然后 gradle 带 `http_proxy=socks5h://127.0.0.1:10808 https_proxy=socks5h://...`（覆盖注册表里错误的 HTTP 代理指向）
  - `rm -rf app/build` 会删掉构建环境（连 PySocks 一起），重建后要重装
  - 注册表系统代理指向 10808（SOCKS 口）时，pip 按 HTTP 用它会握手超时——报 ProxyError 就是这个
- App 已具备现场闭环：**选择模板渲染（SAF 文件选择）+ 保存到相册（Pictures/PosterAgain）+ 分享（FileProvider）**
- 最新 APK：`dist/PosterAgain-android-debug.apk`（141MB，debug 签名，arm64-v8a + x86_64，含 numpy）

## 五、下一步（2026-08-18 晚更新：1/2/4 已完成，见第七节）

1. ~~真实模板验证~~ ✅（第四节）
2. ~~numpy 入包~~ ✅（第四节）
3. render_core 双副本同步脚本化（现手工 cp，见 tasks 4.3）
4. ~~安卓 UI 产品化：批量/保存分享~~ ✅（第七节 Phase A）
5. release 单 ABI 构建评估体积；Windows 端默认策略翻 bundled_first（4.2，需先重校测试基线）
6. 真机批量耗时实测 + 渲染失败单行重试 UI（Phase B）

## 六、桌面端回归基线（全部通过）

test_render_core 7 · test_renderer_memory 9 · test_font_registry 20 · test_render_precheck 14 · test_renderer_features 6 · test_renderer_shapes 14 · test_renderer_transforms 12 · test_renderer_text_decorations 12 · test_renderer_robustness 3 · test_pptx_processor ✅ · test_complete_workflow ✅ · test_personalize 12 · test_data_rows 11

## 七、Phase A 批量个性化（2026-08-18 完成）

**引擎侧（`src/render_core/`，改后需 cp 同步 `android/app/src/main/python/render_core/`）**：
- `personalize.py`：`Personalizer`——模板**一次加载**，`fields()` 识别 ph_（AltText/
  形状名）+ `{{}}` 三种占位符；`render_row()` = apply → `export_presentation` → restore。
  **关键坑**：restore 必须就地替换 spTree 子元素——整树替换会让 python-pptx 的
  `Slide.shapes` 惰性缓存指向游离旧树，后续 apply 改的是孤儿节点、静默失效
- `data_rows.py`：CSV（utf-8-sig/gbk/utf-8 探测）/JSON/XLSX（openpyxl read_only），
  无 pandas
- `match_fields(fields, headers)`：精确 → 剥 ph_ 前缀 → 忽略大小写

**安卓侧**：`poster_bridge.py` 会话 API（start_session/render_index/render_preview/
end_session，返回 JSON 字符串——Chaquopy `asMap()` 泛型推断不可靠）；MainActivity
重写为 BatchScreen（SAF 选模板+数据 → 匹配预览 → 批量进度 → LazyRow 画廊 →
存相册/分享）；openpyxl 入 Chaquopy pip。

**端到端验收（模拟器 `poster34`，`android/e2e_batch.py` 自动驱动）**：
- CSV 与 XLSX 两轮：5 占位符识别、5 行×5 列、5 字段全匹配、5 张全渲染
- 设备 row_0 vs PC 同数据：**0.049% 像素 @幅度 22.6/255**（纯抗锯齿，三带均匀）
- xlsx 轮产物与 CSV 轮**逐字节同尺寸**（同数据 → 同渲染）
- 模拟器单行渲染 <200ms（5 行同一秒内完成）；真机单页基准 623ms/3445ms 见第四节

**排查记录（都是测试脚手架 bug，引擎无差）**：①生成 CSV 时金额 `88,888` 没加引号
被拆列（设备解析成 `88` 是正确 CSV 行为）；②PC 对比字典漏 `ph_` 前缀——`{{}}` 行内
替换剥前缀匹配不受影响，但 ph_ 整框替换要求键完全等于 `ph_字段名`。

**真实业务文件验证（10月递交贺报(3).pptx + KA递交模板.xlsx，2026-08-18）**：
- 6 占位符识别、29 列 xlsx 4 行、5 字段精确匹配（ph_个险趸期判定无对应数据列，
  按设计保留模板原样）、4 张全渲染 ~7s（模拟器）
- **修了两个真实环境问题**：①安卓 SAF 偶发返回**截断显示名**（"KA递交模"丢扩展名）
  → `read_rows` 内容嗅探兜底（zip+xl/→xlsx；{/[→json；文本→csv；空文件/旧 .xls
  明确报错），MainActivity 改用 `OpenableColumns.DISPLAY_NAME` 取真名（双保险）；
  ②e2e 脚本：长文件名在 Downloads 列表截断显示→前缀匹配、列表超屏→滚屏查找、
  抽屉时序不稳→进入验证+重试
- **字体边界（重要业务语义）**：该模板主题东亚字体=微软雅黑。PC（system_first）
  用真雅黑，安卓无雅黑→兜底思源黑体，**物理字体不同→3.5% 内容级像素差异**；
  PC 强制 `POSTERAGAIN_FONT_POLICY=bundled_first` 后 4 行全部 0.011~0.014%
  @amp21（纯抗锯齿）——**引擎管线零差异，跨端差异全部来自字体可用性**。
  各平台选最优可用字体是当前设计（Windows 零回归 / 安卓无 Office 字体）；
  要逐字体一致需业务模板统一用可分发字体（雅黑不可捆绑）或 PC 翻 bundled_first
