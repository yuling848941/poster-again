# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

PPT 批量生成工具 - 基于 PySide6 和 python-pptx 的桌面应用，根据 Excel 数据批量生成 PowerPoint 演示文稿。

**v2.0.0 重要更新**: 使用纯 Python 的 python-pptx 库，无需 Microsoft Office 或 WPS。

## 技术栈

- **GUI**: PySide6 (Qt for Python)
- **PPT 处理**: python-pptx
- **数据处理**: pandas, openpyxl
- **配置管理**: PyYAML, JSON

## 常用命令

```bash
# 启动应用程序
python main.py

# 安装虚拟环境
python -m venv .venv && .venv\Scripts\activate  # Windows

# 安装依赖 (无 requirements.txt，手动安装)
pip install PySide6 python-pptx pandas openpyxl PyYAML psutil

# 运行测试 (单个测试文件，无 pytest 框架)
python tests/unit/test_pptx_processor.py
python tests/unit/test_chengbao_term_data.py
python tests/unit/test_complete_workflow.py

# 打包为 exe
python build_exe.py
打包.bat                                 # Windows 批处理方式
```

**测试说明**: `tests/unit/` 下约 49 个测试文件都是独立脚本，没有 pytest 配置或 conftest.py。没有"运行全部测试"的目标 — 始终运行与更改相关的具体测试文件。

## 核心架构

### 模块结构

```
src/
├── main.py                      # 应用入口
├── ppt_generator.py             # PPT 生成器 (协调层)
├── data_reader.py               # 数据读取 (支持 Excel/CSV/JSON)
├── exact_matcher.py             # 占位符匹配算法
├── excel_detector.py            # Excel 文件检测
├── gui/
│   ├── main_window.py           # 主窗口 (~1500 行，所有 UI 逻辑)
│   ├── ppt_worker_thread.py     # 后台工作线程 (QThread)
│   ├── settings_manager.py      # 设置管理
│   ├── path_manager.py          # 路径记忆
│   ├── match_table_manager.py   # 匹配表格管理
│   ├── config_dialog.py         # 配置对话框
│   ├── simple_config_dialog.py  # 简化配置对话框
│   ├── chengbao_term_input_dialog.py  # 承保趸期输入对话框
│   └── widgets/
│       └── add_text_dialog.py   # 文本增加对话框
├── core/
│   ├── processors/
│   │   └── pptx_processor.py    # PPTX 模板处理器 (纯 Python)
│   ├── interfaces/
│   │   └── template_processor.py # 模板处理器接口
│   ├── detectors/
│   │   └── office_suite_detector.py  # Office 套件检测
│   ├── factory/
│   │   └── office_suite_factory.py # 工厂类
│   └── utils/
│       └── font_checker.py      # 字体检查工具
├── memory_management/
│   ├── memory_data_manager.py   # 内存数据管理
│   ├── memory_optimizer.py      # 内存优化器
│   └── data_formatter.py        # 数据格式化 (千位分隔符/期趸数据)
├── config/
│   ├── __init__.py              # ConfigManager (多重继承 Mixin 模式)
│   ├── path_config.py           # 路径配置 Mixin
│   ├── image_config.py          # 图片生成配置 Mixin
│   ├── placeholder_config.py    # 占位符配置 Mixin
│   └── gui_config.py            # GUI 配置 Mixin
└── term_data_generator.py       # 期趸数据生成器
```

### 数据流

1. 用户选择 PPT 模板 → `PPTXProcessor.load_template()`
2. 用户选择 Excel 数据 → `DataReader.load_excel()` → `MemoryDataManager` 内存处理
3. 自动匹配占位符 → `ExactMatcher.auto_match_placeholders()`
4. 批量生成 → `PPTGenerator.batch_generate()`

### 占位符格式

PPT 模板中的占位符支持三种方式：

| 方式 | 设置方法 | 示例 |
|------|----------|------|
| Alt Text (推荐) | 右键形状 → 编辑替代文本 | `ph_姓名` |
| 形状名称 | 选择窗格重命名 | `ph_公司名称` |
| 文本标记 | 直接输入文本 | `{{字段名}}` |

### 特殊功能

**期趸数据处理**:
- 自动检测 `SFYP2(不含短险续保)` 和 `首年保费` 列
- 计算承保趸期数据，支持用户输入对话框

**千位分隔符** (`data_formatter.py:add_thousands_separator`):
- 自动为数字列添加千位分隔符

## 导入模式

大多数模块使用双重 try/except 导入以支持 `src.` 前缀和扁平导入：

```python
try:
    from src.ppt_generator import PPTGenerator
except ImportError:
    from ppt_generator import PPTGenerator
```

这支持从根目录和 `src/` 内部运行。添加模块时请保持此模式。

## 配置管理

- **配置路径**: `config/config.yaml` (默认)，也支持 `PosterAgain_config.yaml`
- **用户设置**: 通过 `SettingsManager` 保存 (图片生成设置、贺语模板、路径记忆)
- **匹配规则**: 通过 `ConfigManager` 导出/导入 JSON 配置
- **ConfigManager** 使用 Mixin 多重继承组合各配置功能

## 测试

测试文件位于 `tests/unit/`，主要覆盖：
- PPTX 处理器功能
- 承保趸期数据计算
- 配置保存/加载
- GUI 工作流

## OpenSpec 工作流程

项目使用 OpenSpec 进行规范驱动开发：

```bash
# 查看现有规范
openspec list --specs

# 查看活跃变更
openspec list

# 创建新提案
openspec:proposal

# 归档已完成变更
openspec archive <change-id> --yes
```

所有架构变更、新功能开发应遵循 OpenSpec 流程。详见 `openspec/AGENTS.md`。

## 关键设计决策

1. **内存管理**: 使用 `MemoryDataManager` 进行数据缓存，避免临时文件
2. **异步处理**: GUI 使用 `PPTWorkerThread` 避免界面冻结
3. **占位符匹配**: 支持精确匹配、忽略大小写、特殊字符清理三种策略
4. **文本增加**: 支持为匹配结果添加前缀/后缀 (`exact_matcher.py`)
5. **文件占用处理**: PPT 文件使用临时文件 + 复制模式，避免锁定问题

## 打包分发

- PyInstaller 规范文件: `PosterAgain.spec`
- 打包脚本: `build_exe.py` (清理旧 dist/build，运行 pyinstaller)
- 输出: `dist/PosterAgain.exe`
- 图标: `resources/icon.ico`
- 隐藏导入包括: `yaml`, `pandas`, `pptx`, `win32com`, `pythoncom`, `pywintypes`

## 注意事项

- `main.py` 在导入时创建模块级单例 (`memory_manager`, `config_manager`) — 导入 `main` 有副作用
- `ConfigManager()` 无参数调用时默认使用 `config/config.yaml`
- `config/config.yaml` 存储上次用户的**绝对路径** — 不要提交路径更改
- `PPTWorkerThread` 上的 `office_manager` 参数已弃用 (始终设为 `None`) — 保留用于向后兼容
- `logging.basicConfig` 未全局调用；每个模块使用 `logging.getLogger(__name__)`。如需配置日志，在 `main()` 中进行
- `use_com_interface` 配置标志为 `false`，旧 COM 处理器代码已失效
