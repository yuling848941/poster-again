## Context

当前系统支持PPT占位符与Excel数据的精确匹配，以及文本增加功能（前缀/后缀），但缺乏配置持久化机制。用户每次更换模板或数据都需要重新配置匹配规则，影响工作效率。

## Goals / Non-Goals

### Goals:
- 提供配置文件的保存和加载功能
- 实现智能配置验证，确保模板和数据的兼容性
- 保持配置文件的向后兼容性
- 提供用户友好的配置管理界面

### Non-Goals:
- 不修改现有的匹配算法逻辑
- 不改变"递交趸期数据"功能的配置处理
- 不实现配置文件的云同步功能

## Decisions

- **配置文件格式**: 采用JSON格式，便于阅读和调试
- **文件扩展名**: 使用.pptcfg作为专有配置文件扩展名
- **验证策略**: 双重验证 - 先验证配置文件格式，再验证当前模板/数据兼容性
- **错误处理**: 部分加载策略，只应用兼容的配置项，不兼容的项给出提示

## Architecture

### 核心组件
1. **ConfigManager扩展**: 增加占位符配置管理能力
2. **ConfigValidator**: 配置文件验证器
3. **ConfigSerializer**: 配置序列化/反序列化器

### 配置文件结构
```json
{
  "version": "1.0.0",
  "metadata": {
    "created_at": "2024-01-01T00:00:00Z",
    "description": "配置描述",
    "template_info": {
      "file_name": "template.pptx",
      "placeholders": ["ph_name", "ph_title"]
    },
    "data_info": {
      "file_name": "data.xlsx",
      "columns": ["姓名", "职位"]
    }
  },
  "config": {
    "matching_rules": [
      {
        "placeholder": "ph_name",
        "column": "姓名",
        "text_addition": {
          "prefix": "",
          "suffix": "先生"
        }
      }
    ]
  }
}
```

## Risks / Trade-offs

### Risks:
- 配置文件格式变更可能导致兼容性问题
- 智能匹配可能产生误判
- 用户期望管理不当

### Trade-offs:
- **灵活性 vs 简单性**: 选择JSON格式而非二进制格式，提高了可读性但增加了文件大小
- **严格验证 vs 宽容加载**: 采用部分加载策略，提高了容错性但可能隐藏一些问题

## Migration Plan

1. 扩展现有ConfigManager类，增加配置保存/加载方法
2. 创建配置验证器，确保配置文件的正确性
3. 在GUI中添加保存/加载配置的界面元素
4. 集成配置功能到现有的工作流程中

## Open Questions

- 配置文件是否需要加密处理以保护敏感数据？
- 是否需要支持配置文件的导入/导出批量操作？
- 如何处理配置文件中的模板文件路径变更？