## Context
项目是一个PPT批量生成工具，已经实现了基本的PPT生成功能。用户要求添加直接生成图片的功能，将生成的PPT转换为指定格式的图片文件。

## Goals / Non-Goals
- Goals:
  - 支持PPT幻灯片转换为多种图片格式
  - 集成到现有的批量生成流程
  - 保持与现有功能的兼容性
- Non-Goals:
  - 不修改PPT模板处理逻辑
  - 不改变现有的文件命名规则

## Decisions
- Decision: 使用COM接口进行PPT到图片的转换，因为项目已经在使用COM接口保持PPT特效
- Decision: 扩展现有的PPTGenerator类而不是创建新的图片生成器
- Decision: 支持常见图片格式(PNG, JPG)

## Risks / Trade-offs
- COM接口依赖: 需要系统安装Microsoft PowerPoint
- 性能考虑: 图片生成可能比纯PPT生成耗时更长
- 内存使用: 批量生成大量图片可能消耗更多内存

## Migration Plan
- 先实现基础的PNG和JPG格式支持
- 添加图片质量选择功能（原始大小、1.5倍、2倍、3倍、4倍）
- 保持向后兼容，默认行为不变

## Open Questions
- 是否需要支持自定义图片分辨率？
**已解决**: 不需要，按PPT分辨率
- 是否需要支持单页PPT生成单张图片，还是整个PPT生成多张图片？
**已解决**: 按程序现在的逻辑，无需改变