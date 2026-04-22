# Project Context

## Purpose
基于PySide6和Python的PPT批量生成工具，支持从Excel、CSV和JSON文件读取数据，并使用PowerPoint模板自动生成个性化PPT文档。项目通过COM接口与PowerPoint集成，保持PPT中的复杂特效，提供高效的批量文档生成能力。

## Tech Stack
- **核心语言**: Python 3.x
- **GUI框架**: PySide6
- **Office集成**: Python WIN32COM (pywin32)
- **数据处理**: pandas (Excel/CSV), json
- **配置管理**: 原生JSON格式
- **异步处理**: Python threading + PySide6信号槽
- **内存管理**: 自研MemoryDataManager + MemoryOptimizer

## Project Conventions

### Code Style
- **命名约定**:
  - 类名: PascalCase (如 `TemplateProcessor`)
  - 函数名: snake_case (如 `process_template`)
  - 变量名: snake_case (如 `data_manager`)
  - 常量: UPPER_SNAKE_CASE
- **文档字符串**: 所有公共函数和类使用中文docstring
- **类型注解**: 核心模块使用类型提示 (PEP 484)
- **格式化**: 遵循PEP 8，缩进4个空格

### Architecture Patterns
- **模块化设计**: 单一职责原则，每个模块功能明确
- **COM集成模式**: 通过win32com.client与PowerPoint应用程序交互
- **内存管理优化**: 使用MemoryDataManager和MemoryOptimizer优化大数据集处理
- **异步处理**: GUI使用PPTWorkerThread进行后台任务处理，避免界面冻结
- **占位符约定**: PPT元素命名格式为`ph_字段名`，通过PowerPoint选择窗格设置

### Testing Strategy
- **单元测试**: 核心模块(`src/core/`, `src/data_reader.py`, `src/ppt_generator.py`)必须有对应测试文件
- **集成测试**: 测试完整的数据处理流程，包括数据读取→模板处理→PPT生成
- **GUI测试**: 使用专门测试文件验证界面功能
- **内存测试**: 验证大数据集处理的内存使用优化
- **测试文件命名**: `test_[模块名].py`

### Git Workflow
- **分支策略**: 使用Git Flow，主分支(main) + 开发分支(develop) + 功能分支(feature/)
- **提交规范**: 使用约定式提交格式 `type(scope): subject`
  - type: feat(新功能), fix(修复), docs(文档), refactor(重构), test(测试), chore(构建)
- **变更记录**: 重要变更在CHANGELOG.md中记录
- **OpenSpec工作流**: 所有架构变更、新功能开发或重大优化都应遵循OpenSpec流程

## Domain Context
- **PPT模板处理**: PPT中的文本框、形状等元素需通过选择窗格命名为`ph_字段名`格式
- **数据匹配算法**: 使用`ExactMatcher`类进行精确匹配，支持数字格式化和千位分隔符
- **内存优化**: 大数据集处理时使用`MemoryDataManager`进行内存管理，支持数据缓存和释放
- **COM对象生命周期**: PowerPoint COM对象必须在使用后正确释放，避免应用程序实例残留
- **字段匹配规则**:
  - 优先精确匹配字段名
  - 支持数字类型自动格式化
  - 支持千位分隔符(,)和小数点(.)处理

## Important Constraints
- **平台限制**: 仅支持Windows平台 (依赖PowerPoint COM接口)
- **PowerPoint依赖**: 需要安装Microsoft PowerPoint
- **内存限制**: 大数据集(>1000条)处理时必须使用内存优化器
- **占位符命名**: PPT元素命名必须严格遵循`ph_字段名`格式，否则无法匹配
- **数据格式**: Excel/CSV/JSON文件必须包含有效的字段映射关系

## External Dependencies
- **pywin32**: Windows COM接口访问PowerPoint (pythoncom, win32com.client)
- **PySide6**: GUI框架
- **pandas**: Excel/CSV文件读取和处理
- **openpyxl**: Excel文件写入支持
- **Python logging**: 日志记录模块
- **Microsoft PowerPoint**: 必须安装，用于PPT生成
