# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

Always open `@/openspec/AGENTS.md` when the request:
- Mentions planning or proposals (words like proposal, spec, change, plan)
- Introduces new capabilities, breaking changes, architecture shifts, or big performance/security work
- Sounds ambiguous and you need the authoritative spec before coding

Use `@/openspec/AGENTS.md` to learn:
- How to create and apply change proposals
- Spec format and conventions
- Project structure and guidelines

Keep this managed block so 'openspec update' can refresh the instructions.

<!-- OPENSPEC:END -->

## 项目概述

这是一个基于PySide6和Python的PPT批量生成工具，支持从Excel、CSV和JSON文件读取数据，并使用PowerPoint模板自动生成个性化PPT文档。项目通过COM接口与PowerPoint集成，保持PPT中的复杂特效。

## 核心架构

### 主要模块结构

- **`src/core/template_processor.py`**: PPT模板处理器，通过COM接口连接PowerPoint，处理占位符提取和内容替换
- **`src/data_reader.py`**: 数据读取器，支持Excel、CSV、JSON格式，集成了内存数据管理器
- **`src/ppt_generator.py`**: PPT生成器，协调模板处理和数据读取进行批量生成
- **`src/gui/main_window.py`**: GUI主界面，基于PySide6
- **`src/memory_management/`**: 内存管理系统，包含内存优化和数据格式化
- **`src/config_manager.py`**: 配置管理器
- **`src/exact_matcher.py`**: 精确匹配算法，处理PPT占位符与数据字段的匹配

### 关键设计模式

1. **COM集成模式**: 通过`win32com.client`与PowerPoint应用程序交互
2. **内存管理优化**: 使用`MemoryDataManager`和`MemoryOptimizer`优化大数据集处理
3. **异步处理**: GUI使用`PPTWorkerThread`进行后台任务处理
4. **占位符命名约定**: PPT元素通过选择窗格命名为`ph_字段名`格式

## 常用开发命令

### 应用程序运行

```bash
# GUI模式启动
python main.py

# CLI模式（如果存在）
python cli.py --template template.pptx --data data.xlsx --output ./output
```

### 测试和验证

```bash
# 内存管理相关测试
python test_memory_data_processing.py
python test_complete_memory_system.py

# 功能测试
python test_data_reader_memory.py
python test_ppt_generator_memory.py

# 模板和占位符测试
python extract_ppt_placeholders.py
python test_placeholder_fix.py
```

### 开发辅助工具

```bash
# 检查PPT占位符
python check_ppt_placeholders.py

# 调试数据匹配
python debug_placeholder_mismatch.py

# 内存清理
python cleanup_tempfiles.py
```

## 开发约定

### 代码组织

- **模块职责**: 每个模块保持单一职责，避免过度复杂化
- **错误处理**: 使用Python标准logging模块，提供详细的错误日志
- **类型注解**: 核心模块使用类型提示提高代码可读性
- **文档字符串**: 所有公共函数和类都应有中文文档字符串

### 数据处理规范

1. **占位符格式**: PPT元素必须命名为`ph_字段名`格式，通过PowerPoint选择窗格设置
2. **字段匹配**: 使用`ExactMatcher`类进行精确匹配，支持数字格式化和千位分隔符
3. **内存优化**: 大数据集处理时使用`MemoryDataManager`进行内存管理
4. **数据验证**: 所有输入数据都应进行格式验证

### COM接口使用

- PowerPoint应用程序通过`TemplateProcessor`类管理
- COM连接应在使用后正确释放，避免应用程序实例残留
- 支持批量处理时的COM对象生命周期管理

### 测试策略

- **单元测试**: 核心模块都有对应的测试文件
- **集成测试**: 测试完整的数据处理流程
- **GUI测试**: 使用专门的GUI测试文件验证界面功能
- **内存测试**: 验证大数据集处理的内存使用优化

## 配置管理

- 使用`ConfigManager`类进行统一配置管理
- 支持JSON格式配置文件
- 配置项包括：PPT处理参数、数据读取选项、匹配算法参数等

## 性能考虑

- **批量处理**: 使用内存优化器处理大数据集
- **异步操作**: GUI操作使用工作线程避免界面冻结
- **COM资源管理**: 及时释放PowerPoint COM对象
- **缓存机制**: 内存数据管理器提供数据缓存功能

## 错误处理和日志

- 使用Python标准logging模块，日志级别设置为INFO
- GUI模式通过`QMessageBox`显示用户友好的错误信息
- 所有异常都应记录详细的堆栈跟踪信息
- 提供详细的操作日志用于问题排查

## OpenSpec工作流程

项目使用OpenSpec进行规范驱动的开发：

1. **提案创建**: 使用`/openspec:proposal`命令创建新的功能提案
2. **设计阶段**: 提案通过后进行技术设计和任务分解
3. **实现阶段**: 按照任务清单进行开发实现
4. **归档阶段**: 完成后使用`/openspec:archive`归档变更

所有架构变更、新功能开发或重大优化都应遵循OpenSpec流程。