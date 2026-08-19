## 1. Phase 1 引擎独立成包
- [x] 1.1 五模块迁入 `src/render_core/`，内部 import 改 render_core 路径
- [x] 1.2 `src/core/processors/` 留 5 个 sys.modules 转发 shim（旧路径零回归）
- [x] 1.3 `render()` 一句话 API（路径/bytes 输入，RenderResult 输出）
- [x] 1.4 `tests/unit/test_render_core.py`（含独立目录 + 干净子进程的"安卓模拟"自包含测试，逐像素一致）
## 2. Phase 2 内存受限渲染路径
- [x] 2.1 `_load_rgba`：JPEG draft / 先缩放后补 alpha / 大 blob 临时文件解码
- [x] 2.2 `tests/unit/test_renderer_memory.py`：像素等价（RGB/RGBA/P+透明/srcRect）+ 峰值内存（PNG 384MB<420 / JPEG 143MB<250）
## 3. Phase 3 安卓壳
- [x] 3.1 SDK/Gradle 环境从零搭建（cmdline-tools + licenses + 组件 + Gradle 8.6）
- [x] 3.2 `android/` 工程脚手架（AGP 8.2.2 + Compose + Chaquopy 16.1.0 / Py3.12）
- [x] 3.3 嵌入 render_core 副本 + 捆绑字体 + 测试模板 assets + 桥接层
- [x] 3.4 `assembleDebug` 构建通过（86MB APK）
- [x] 3.5 模拟器端到端验证（安装 → 渲染 → 截图）
  （2026-08-18 完成：Pixel 6 AVD / Android 14 x86_64，点击渲染 745ms 出图；
  干净模板预检零告警；设备输出与 PC bundled_first 渲染版式逐像素一致，
  仅 0.4% 抗锯齿级差异（两端 FreeType 版本不同）；过程修复安卓端
  bundled_font_dir 同级候选缺失问题）
## 4. 后续（独立变更）
- [x] 4.1 numpy 加入 Chaquopy pip（渐变文字支持）+ APK 体积评估
  （2026-08-18 完成：numpy 1.26.2 双 ABI 入包，APK 121→141MB；同时引擎新增
  纯 Pillow 渐变兜底路径（无 numpy 自动退化，与 numpy 路径差异 ≤1/255），
  numpy 永远优先；承保贺报（渐变模板）设备端渲染验证通过，与 PC 差异 0.31% 抗锯齿级）
- [x] 4.2 真实模板性能实测（Phase 4）；release 单 ABI 构建
  （2026-08-18 真机+真模板收官：测试模板 623ms、50MB 每日贺报 3445ms/张、
  峰值 PSS ~233MB、预检零问题、与 PC 渲染差异 0.002% 像素——业务验收通过；
  release 单 ABI 构建仍待做）
- [x] 4.3 render_core 双副本同步脚本化
  （2026-08-19 完成：tools/sync_render_core.py——src/render_core → 安卓副本单向
  同步 *.py，LF 归一，跳过安卓专属 resources/ 字体；--check 防漂移模式供 CI）

## 5. Phase A 批量个性化（2026-08-18 完成）
- [x] 5.1 `render_core/personalize.py`：Personalizer——模板一次加载，fields() 识别
  ph_（AltText/形状名）与 {{}} 三种占位符；apply/restore 就地换 spTree 子元素
  （整树替换会令 python-pptx 惰性缓存指向游离旧树，后续 apply 静默失效）；
  render_row = apply → export_presentation → restore；match_fields 精确→剥前缀→忽略大小写
- [x] 5.2 `render_core/data_rows.py`：CSV（utf-8-sig/gbk/utf-8 探测）/JSON/
  XLSX（openpyxl read_only）无 pandas 读取，统一返回 (headers, rows)
- [x] 5.3 安卓批量 UI（BatchScreen）：模板+数据 SAF 选择、字段匹配预览、
  批量渲染进度、LazyRow 结果画廊、存相册（MediaStore）/分享（FileProvider）
- [x] 5.4 `poster_bridge.py` 会话 API（start_session/render_index/render_preview/
  end_session），openpyxl 入 Chaquopy pip；桥接返回 JSON 字符串规避
  Chaquopy asMap() 泛型推断问题
- [x] 5.5 模拟器端到端（`android/e2e_batch.py` 自动驱动）：CSV 与 XLSX 两轮各
  5 行×5 字段全匹配全渲染；设备 row_0 与 PC 同数据渲染差异 0.049% 像素
  @幅度 22.6/255（纯抗锯齿级，三带均匀），xlsx 轮产物与 CSV 轮逐字节同尺寸
  （同数据同渲染）；模拟器单行渲染 <200ms。排查中修掉两处测试脚手架 bug：
  生成 CSV 未给含逗号金额加引号（设备解析为 '88' 属正确 CSV 行为）、
  PC 对比字典漏 ph_ 前缀导致 ph_ 整框替换被跳过（{{}} 行内剥前缀匹配不受影响）
- [x] 5.6 PC 侧测试：test_personalize.py 12/12、test_data_rows.py 11/11
- [x] 5.7 （顺延 Phase B）真机批量耗时实测、release 单 ABI、渲染失败单行重试 UI
  （2026-08-19：release 单 ABI 完成——正式签名 keys/posteragain-release.jks +
  keystore.properties（均 gitignored），assembleRelease 仅 arm64-v8a 77MB，
  apksigner 证书 CN=PosterAgain；单行失败重试 UI 完成——AppState per-row
  catch + retryFailedRows + RenderScreen 失败卡片。真机批量耗时实测仍待做：
  用 dist/PosterAgain-1.1.0-release.apk 在真机安装冒烟）
- [x] 5.8 真实业务文件端到端（10月递交贺报(3).pptx + KA递交模板.xlsx）
  （2026-08-18 完成：6 占位符识别、29 列 xlsx 4 行、5 字段精确匹配
  （ph_个险趸期判定在数据中无对应列，按设计保留原样）、4 张全渲染 ~7s；
  修复两个真实环境问题——①安卓 SAF 偶发返回截断显示名（"KA递交模"丢扩展名）
  → read_rows 增加内容嗅探兜底（zip+xl/→xlsx、{[/[→json、文本→csv、空文件/旧xls明确报错），
  MainActivity 改用 OpenableColumns.DISPLAY_NAME 取真名；②e2e 脚本
  长文件名截断匹配/列表滚屏/Downloads 进入验证重试。
  字体边界确认：模板主题东亚字体=微软雅黑，PC system_first 用真雅黑、安卓无雅黑
  兜底思源黑体 → 物理字体不同出现 3.5% 内容级差异；PC 强制 bundled_first 后
  4 行全部 0.011~0.014% @amp21 纯抗锯齿级——引擎管线零差异，
  差异全部来自字体可用性语义（各平台选最优可用字体，同字体时逐像素一致））
- [x] 5.9 分享/存相册真机化验证（2026-08-18）
  （修复 FileProvider 只声明了旧 render_out/ 而批量产物在 batch/、预览在 preview/
  ——点"分享选中"会 IllegalArgumentException 崩溃；file_paths.xml 补齐三目录
  + 分享按钮加 runCatching 防御。模拟器实测：全部存相册 4/4 入库 MediaStore
  （Pictures/PosterAgain）；分享选中弹出系统分享面板（Bluetooth/Messaging/Print），
  进程存活不崩。真机上分享目标即微信/钉钉等）
- [x] 5.10 输出图片命名与 PC 版对齐（2026-08-18）
  （render_core 新增 row_base_name/sanitize_filename——与 ppt_generator 同规则：
  f"{序号}_{姓名}"（姓名列=精确'姓名'或首个含'姓名'列，空/NaN 兜底 output；
  非法字符<>:"/\|?*→'_'，整体首尾去点空），与 PC _sanitize_filename 逐例对照一致；
  poster_bridge.render_index 渲染后重命名（单页 {base}.png / 多页 {base}_{i}.png，
  与 pptx_processor.export_to_images 重命名规则一致）；存相册 DISPLAY_NAME 改用
  文件原名。实测 KA 数据 4 行：1_陈春兰 / 2_郝杜娟 / 3_郝杜娟（同名行序号区分）/
  4_黄云彩，内部文件、分享附件名、相册名三处一致）
- [x] 5.11 文字效果与渐变颜色修复（争霸赛承保贺报1.pptx 实测，2026-08-18）
  （用户报障：PPT 正常但图片缺效果/金字发暗。根因四项——①文字 glow 发光未实现
  （模板大量 glow+outerShdw+gradFill 文字）→ 实现三层高斯外扩 lighter 合成 +
  screen 加色混合（PowerPoint 的 glow 是光晕会照亮背景，普通 alpha 混合偏暗；
  与 COM 真值对照调增益 1.4，标题区差异 1.77/255）；②prstClr 预设色（white 等）
  不支持 → 解析为黑；③lumMod/lumOff 亮度调制不支持；④渐变同 pos 多个 stop 未按
  PowerPoint 语义取后者（np.interp 对非单调 pos 行为未定义）。文字渐变/图片填充
  投影空间从字形 bbox 改为文本框 box（PowerPoint 语义）。COM 三角度探针验证
  渐变角度语义一致（0°左→右/90°上→下/315°左下→右上）；prstClr/lumMod/
  stop 覆盖均有单测与最小复现验证；预检 glow 从 TEXT_EFFECTS 退役。
  回归：render_core 7、features 9、shapes 14、transforms 12、decorations 12、
  memory 9、robustness 3、precheck 14、personalize 19、data_rows 15、
  font_registry 20、complete_workflow SUCCESS 全绿）
- [x] 5.12 两端重新打包 + 安卓真模板效果验证（2026-08-18）
  （Windows dist/PosterAgain.exe 108MB 重打包（numpy 随引擎入包，较旧版 +35MB）；
  安卓 dist/PosterAgain-android-debug.apk 141MB 重打包含 glow/颜色修复引擎。
  模拟器 e2e：争霸赛承保贺报1.pptx + 承保快报KA.xlsx——10 行全渲染，
  3 字段自动匹配（ph_姓名→营销员姓名 经唯一子串兜底，ph_险种 多候选不猜
  与 GUI 一致）；设备 vs PC(bundled_first) 0.017% @amp20 纯抗锯齿级，
  glow/渐变修复两端逐像素一致。过程中查明旧疑案：此前 SAF"截断显示名"
  实为 adb push 中文文件名被截断（设备文件真实名字就是"争霸赛承"无扩展名），
  内容嗅探兜底因此仍是必要防御；e2e 用 ASCII 文件名推设备）
- [x] 5.13 glow 五色统一校准（2026-08-18，回答"换颜色要不要重调"）
  （构造红/橙/黄/绿/蓝五色同参数 glow 探针 + COM 真值，网格扫描 σ组×增益×伽马。
  定稿 GLOW_SIGMA_KS=(0.08,0.18,0.3) / GLOW_GAIN=2.45 / GLOW_GAMMA=0.7，
  全色相强度比 ±4%、光晕面积比 ≤+19%——常数为色相无关，模板换任意发光色无需重调。
  过程发现：①增益在 alpha 封顶后失效，浓度差本质是高斯边缘半亮 → 伽马提亮曲线
  （LUT 实现）解决；②PowerPoint 确实对 glow 颜色应用 satMod（ED7D31 探针对照证实）；
  ③真实模板"意马当先"区按探针口径实测面积比 0.99/强度比 0.98——两端一致，
  早前"橙光差 2.4 倍"为 R>120 阈值口径的非线性放大，非真实差异。
  回归 8 项全绿；双端副本同步重打包 dist/PosterAgain.exe(108MB) +
  PosterAgain-android-debug.apk(141MB)）
- [x] 5.14 PPTX 内嵌字体支持 + COM 分层导出（2026-08-18）
  （A. 嵌入字体：font_registry.load_embedded_fonts 解析 embeddedFontLst/fntdata——
  未压缩/ODTTF 混淆字体提取为临时文件注册；resolve 顺序=系统/捆绑 → 嵌入
  → 默认链（PowerPoint 同语义，Windows 零回归）；Personalizer 与
  export_to_images 入口自动加载。单测 10/10（含真实 PowerPoint 产物验证）。
  边界：PowerPoint 子集化压缩字体（微软私有格式，LibreOffice 亦不支持）明确
  告警跳过、回退系统字体——该场景由 COM 分层导出覆盖。
  B. COM 分层导出：processing.image_export_engine=auto/office/python
  （环境变量 POSTERAGAIN_IMAGE_ENGINE 可覆盖；旧 use_com_interface=true→office）。
  auto=检测到 Office 用 COM（100% 保真）否则纯 Python；COM 失败自动降级保持出图。
  实测本机 auto→COM 1080x1920 导出成功；强制 python 与 COM 对比平均差 7.27/255
  （光晕曲线近似差异，符合预期）。test_com_exporter_no_kill 3/3。
  回归全套 12 项绿；双端重打包 dist/PosterAgain.exe(112MB)+APK(141MB)）
- [x] 5.15 字体一致性收口：本机扩展字体静默增强（2026-08-18 晚）
  （用户发现 WPS 与 MS Office 打开同模板显示不同——破案：模板字体名
  "字魂疯藏山海经体"为笔误（真名"字魂馆藏山海经体"），系统未装、模板未嵌入。
  WPS = 云字体自动下载 + 近名容错；MS PowerPoint = 严格按名替换。引擎原 fallback
  雅黑，故与 WPS 显示不同。方案（按用户决策：隐藏能力，不作为显性功能——云字体
  是会员权益且授权仅限 WPS 内使用）：font_registry 增加本机扩展字体目录静默扫描
  （%APPDATA%/kingsoft/office6/docerFonts 及 hiddenfonts，POSTERAGAIN_WPS_FONTS_DIR
  可覆盖；fontTools 读本地化 name 表中文名 + OS/2.fsSelection 判粗体，PIL 英文名
  兜底）+ 近名容错（编辑距离<=1 且唯一候选，复刻 WPS 行为）。日志全部 debug 级，
  无文件不回退无感知；机器上有缓存即与 WPS 显示一致。实测：字魂疯藏山海经体
  →近名命中字魂馆藏山海经体.ttf，渲染字形与 WPS 导出一致（IoU 0.55 对齐后，
  修复前 0.1x）。test_font_registry 23/23；全套回归 12 项绿。版权提醒：业务上
  要求跨机器一致时，正式方案应为模板改用可分发字体或购买商用授权，不得将
  WPS 缓存字体拷入产品分发）
- [x] 5.16 移动端字段配置（手动绑定+固定值+前后缀+记住配置，2026-08-19）
  （需求：ph_险种 类多候选列自动匹配不猜、ph_个险趸期判定 类数据无对应列时，
  安卓端人工兜底。字段配置统一模型 {source: auto|column|fixed, value, prefix,
  suffix}——与 PC 端 matching_rules + text_addition_rules 语义对齐。
  bridge：start_session 返回 headers；apply_field_configs(sid, json) 设置会话
  手动配置；render_index 按配置取值并拼前后缀。Kotlin：字段绑定列表每项可点
  → FieldConfigDialog（源三选/列列表/固定值输入/前后缀）；持久化
  SharedPreferences，键=模板名+字段列表指纹（模板字段变更自动作废），
  openSession 自动套用；保存即时+渲染前推送。
  验证：bridge 单测 14/14（绑列/固定值/前后缀/恢复自动/命名/restore 安全）；
  模拟器 e2e：基本流程零回归（10 行全渲染）+ UI 实测固定值配置保存生效 +
  杀进程重启后自动套用。范围（用户拍板）：字段级不做按行改值）
- [x] 5.17 溢出文字变黑修复（2026-08-19）
  （用户报障：文本框不够大时溢出文字变黑。根因：5.11 把文字渐变投影改为
  按文本框 box 后，渐变图只贴框内，框外 RGB=(0,0,0) 而文字蒙版 alpha 仍在
  ——溢出文字成黑（混底后灰）。PowerPoint 真值实测：溢出文字渐变为端点色
  延伸。修复：_make_linear_gradient 增加 ref_w/ref_h（t 归一化基准=框尺寸）
  + off_x/off_y（投影偏移）；文字渐变生成范围=框∪字形实际范围，框外 t 钳制
  到端点色；纯 Pillow 路径同步；blip 图片填充溢出区按边缘延伸。
  验证：A/B 改前改后争霸赛渲染【差异像素 0、最大差 0】（框内逐位等价，
  已校准金字渐变零影响）；溢出探针框下区 黑(133,133,133)→(254,243,219)
  ≈真值(254,246,233)；blip 溢出冒烟非黑；features 测试加溢出用例 12/12；
  回归 12 项全绿；双端重打包）
