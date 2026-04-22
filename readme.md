# PPT批量生成工具

## 🎉 重要更新

**v2.0.0 版本使用纯 Python 实现，无需 Microsoft Office 或 WPS！**

## 📋 功能简介

PPT批量生成工具是一个基于 Python 的桌面应用程序，可以根据 Excel 数据批量生成 PowerPoint 演示文稿。

### 主要特性

- ✅ **无需 Office/WPS**：使用 python-pptx 库直接处理 PPT 文件
- ✅ **多种占位符方式**：支持 Alt Text、形状名称、文本标记
- ✅ **智能匹配**：自动匹配 Excel 列和 PPT 占位符
- ✅ **批量生成**：一键生成数百份个性化 PPT
- ✅ **格式保留**：完整保留模板样式、字体、颜色
- ✅ **跨平台**：支持 Windows、macOS、Linux

## 🚀 快速开始

### 环境要求

- Python 3.8+
- **无需 Microsoft Office**
- **无需 WPS Office**

### 安装

```bash
# 克隆仓库
git clone <repository-url>
cd poster-again

# 创建虚拟环境（推荐）
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux

# 安装依赖
pip install -r requirements.txt
```

### 运行

```bash
python main.py
```

## 📖 使用指南

### 1. 准备模板

创建 `.pptx` 文件，使用以下方式添加占位符：

**推荐方式：Alt Text（替代文本）**
1. 右键形状 → "编辑替代文本"
2. 输入：`ph_字段名`
3. 例如：`ph_姓名`、`ph_公司名称`

**其他方式：**
- 形状名称：`ph_字段名`
- 文本标记：`{{字段名}}`

### 2. 准备数据

创建 Excel 文件，列名与占位符对应：

| 姓名 | 公司名称 | 金额 |
|------|----------|------|
| 张三 | ABC公司 | 10000 |
| 李四 | XYZ公司 | 20000 |

### 3. 生成PPT

1. 打开程序
2. 选择模板文件
3. 选择数据文件
4. 点击"自动匹配"
5. 点击"批量生成"

详细使用说明请查看 [docs/usage_guide_no_office.md](docs/usage_guide_no_office.md)

## 🏗️ 项目结构

```
poster-again/
├── main.py                      # 程序入口
├── requirements.txt             # 依赖列表
├── config/                      # 配置文件
│   └── config.yaml
├── src/                         # 源代码
│   ├── core/                    # 核心模块
│   │   ├── processors/          # 处理器
│   │   │   └── pptx_processor.py    # PPTX处理器（纯Python）
│   │   ├── interfaces/          # 接口定义
│   │   └── factory/             # 工厂类
│   ├── gui/                     # GUI界面
│   │   ├── main_window.py       # 主窗口
│   │   └── ppt_worker_thread.py # 工作线程
│   ├── ppt_generator.py         # PPT生成器
│   ├── data_reader.py           # 数据读取
│   └── exact_matcher.py         # 匹配器
├── docs/                        # 文档
│   ├── usage_guide_no_office.md # 使用指南
│   └── refactor_plan_no_office.md # 重构计划
├── output/                      # 输出目录
└── tests/                       # 测试文件
```

## 🛠️ 技术栈

- **GUI**: PySide6 (Qt for Python)
- **PPT处理**: python-pptx
- **数据处理**: pandas, openpyxl
- **配置管理**: PyYAML

## 🔄 版本历史

### v2.0.0 (2026-03-24)
- ✅ 移除 Microsoft Office 依赖
- ✅ 移除 WPS Office 依赖
- ✅ 使用 python-pptx 库处理PPT文件
- ✅ 支持 Alt Text 占位符
- ✅ 简化安装流程

### v1.x
- 需要安装 Microsoft Office 或 WPS
- 使用 COM 接口操作PPT

## 📝 许可证

MIT License

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📧 联系方式

如有问题，请提交 Issue 或联系开发者。
