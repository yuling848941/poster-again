# pptx-render-core

纯 Python 的 **PPTX → 图片** 渲染库。不依赖 Microsoft Office、WPS 或 LibreOffice，
Windows / macOS / Linux / Android 都能跑。

适合这类场景：模板固定、只换文字的**海报 / 贺报 / 证书 / 邀请函**批量出图。

- **文字效果**：线性渐变、图片填充、发光、外阴影、描边、删除线、下划线、合成斜体
  （观感对照 PowerPoint 校准；实现见底层库 [fontfx](../fontfx)）
- **版式还原**：母版 / 版式 / 占位符格式继承、行距 / 字距 / 对齐 / 锚定、normAutofit
- **图形**：背景、图片（含裁剪）、组合、旋转 / 翻转、18 种预设形状（纯色 / 渐变 / 图片填充）
- **字体**：自带思源黑体 / 思源宋体（SIL OFL），跨平台结果一致；也支持文档内嵌字体
- **渲染前预检**：模板里有不支持的元素会明确列出来，而不是悄悄画错
- **批量个性化**：模板只解析一次，逐行「替换 → 渲染 → 恢复」，50MB 的模板也不必每行重新解压

## 安装

```bash
pip install pptx-render-core            # 依赖 python-pptx / lxml / Pillow / fontfx
pip install "pptx-render-core[numpy]"   # 渐变走向量化实现，更快（不装也能跑）
pip install "pptx-render-core[fontinfo]"  # 用 fontTools 读字体内部名，粗体判断更可靠
```

在本仓库里开发时用可编辑安装：`pip install -e libs/fontfx -e libs/render_core`
（`pip install -r requirements.txt` 已经包含这一步）。

## 快速上手

渲染一个文件：

```python
from render_core import render

result = render("template.pptx", "out/", image_format="PNG", quality=2.0)
print(result.images)       # ['out/slide_1.png', ...]
for issue in result.issues:
    print(f"第{issue.slide}页 {issue.shape!r} [{issue.code}] {issue.message}")
```

`quality` 是缩放倍数：1.0 = 96dpi，2.0 / 3.0 更清晰也更慢、更占内存。
`issues` 非空时图片仍会生成，只是这些元素的渲染是近似的（见下面「预检」）。

用数据表批量出图：

```python
from render_core import Personalizer, read_rows, match_fields

pz = Personalizer("template.pptx")                  # 模板只解析一次
headers, rows = read_rows(open("data.xlsx", "rb").read(), "data.xlsx")
mapping = match_fields(pz.fields(), headers)        # 占位符 -> 列名

for i, row in enumerate(rows):
    data = {field: row[column] for field, column in mapping.items()}
    images = pz.render_row(data, "out/")            # 替换 -> 渲染 -> 自动恢复模板
```

跑 `python libs/render_core/examples/render_example.py` 会生成一个带渐变标题、
发光副标题和描边文字的示例 PPTX 并渲染成 PNG。

## 占位符约定

模板里用下面三种方式之一标记「这个位置要替换」：

1. 形状的**替代文本**（Alt Text）以 `ph_` 开头，如 `ph_姓名` —— 推荐，用户在 PowerPoint 里看不见
2. 形状**名称**以 `ph_` 开头
3. 正文里的 `{{字段名}}` —— 行内替换，一行里可以放多个

前两种是整框替换：保留第一段第一个 run 的格式，其余内容清空。
`match_fields` 按列名精确匹配 → 去掉 `ph_` 前缀 → 忽略大小写 / 空白 → 唯一子串兜底；
有歧义时不猜，留在结果外面由调用方决定。

## 接口

| 入口 | 用途 |
|---|---|
| `render(pptx, output_dir, image_format='PNG', quality=1.0, precheck=True)` | 渲染一个 PPTX（路径或 bytes），返回 `RenderResult(images, issues)` |
| `Personalizer(pptx)` | 批量会话：`fields()` 列占位符，`render_row(data, out_dir)` 渲染一行，`apply()` / `restore()` 手动控制 |
| `read_rows(data_bytes, filename)` | 读 csv / json / xlsx 成 `(表头, [{列: 值}])`，不需要 pandas |
| `match_fields(fields, headers)` | 占位符 ↔ 表头 匹配 |
| `row_base_name(headers, row, index)` | 输出文件基础名：`序号_姓名` |
| `PillowImageExporter` | 细粒度控制：`precheck(path)`、`export_to_images(...)`、`export_presentation(prs, ...)` |
| `RenderPrecheck(installed_fonts)` | 只做预检 |
| `font_registry` | 字体解析与策略（见下） |

## 预检

引擎对 PPTX 的支持是**白名单**式的。预检把不支持或只能近似的元素列成 `PrecheckIssue`
（字段：`slide` / `shape` / `code` / `message`），问题码是稳定的，可以按码做程序化决策：

| 不会渲染 | 说明 |
|---|---|
| `TABLE` `CHART` `SMARTART` `OLE` `CONNECTOR` | 表格、图表、SmartArt、嵌入对象、连接线 |
| `SHAPE_GEOM` | 自定义几何或白名单外的预设形状（其中的文字仍会渲染） |
| `WORDART` | 艺术字变形 |
| `VERT_TEXT` | 竖排文字 |

| 近似渲染 | 说明 |
|---|---|
| `FONT_MISSING` | 字体未安装，回退默认字体，与 PowerPoint 的替换结果可能不同 |
| `TEXT_EFFECTS` | 文字倒影 / 3D 等效果（发光、外阴影、描边已支持） |
| `TEXT_STROKE` | 文字的非纯色描边按无描边渲染 |
| `TEXT_DECOR` | 波浪线 / 加粗下划线等按单实线渲染 |
| `GRAD_PATH` | 路径渐变（径向等），仅支持线性 |
| `SHAPE_FILL` | 形状的路径渐变 / 图案填充 |
| `SHAPE_EFFECTS` | 外阴影以外的形状效果 |
| `SHAPE_LINE` | 形状的虚线 / 非纯色描边 |
| `PIC_EFFECTS` | 图片的透明度 / 双色调等效果 |
| `GRP_TEXT_SCALE` | 组合非等比缩放，其中的文字按平均比例近似 |
| `NO_WRAP` `BULLETS` `LNSPC_PTS` `ALIGN_JUST` | 禁止换行、项目符号、固定磅值行距、两端对齐 |
| `INHERIT_UNRESOLVED` `BG_REF` `BG_PATTERN` | 继承链无法解析、主题背景样式无法解析、背景图案 |

「近似渲染」类问题不会阻止出图，只是这些地方和 PowerPoint 打开的效果会有差异。
对还原度要求高的模板，建议先看一遍 `issues`。

## 行为开关

| 环境变量 | 作用 |
|---|---|
| `POSTERAGAIN_FONT_POLICY` | 字体解析策略：`auto`（默认，Windows 系统字体优先 / 其他平台捆绑字体优先）、`system_first`、`bundled_first` |
| `POSTERAGAIN_FONTS_DIR` | 额外的字体目录，优先级最高 |
| `POSTERAGAIN_PIC_CACHE_MB` | 图片缓存上限，默认 256，`0` 关闭。手机上建议 64 |

图片缓存存的是**缩放后**的位图（宽 × 高 × 4 字节），用来让海报背景图在逐行渲染时复用；
命中与否像素完全一致。

## 没法用 pip 的环境

把下面三样东西放到一起（同一个目录，或在 `sys.path` 上）即可：

```
render_core/            # libs/render_core/src/render_core/，含 resources/fonts/
fontfx/                 # libs/fontfx/src/fontfx/
```

安卓工程（Chaquopy）就是这样做的，`tools/sync_render_core.py` 负责同步并检查漂移。

## 测试

```bash
python libs/render_core/tests/run_tests.py
```

每个测试文件也都是独立的脚本。改动引擎后建议再跑一次像素回归，确认输出没变：

```bash
python tools/render_regression.py corpus --out before.jsonl "你的模板.pptx"
# ... 改代码 ...
python tools/render_regression.py corpus --out after.jsonl "你的模板.pptx"
python tools/render_regression.py verify before.jsonl after.jsonl
```
