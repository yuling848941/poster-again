# 图片生成规格

## Purpose
支持将PPT演示文稿转换为图片文件，包括PNG和JPG格式，并提供质量选择选项。

## Requirements

### Requirement: 图片格式支持
系统 SHALL 支持PNG和JPG图片格式的输出

#### Scenario: PNG格式生成
- **WHEN** 用户选择PNG格式进行图片生成
- **THEN** 系统应生成高质量的PNG图片文件

#### Scenario: JPG格式生成
- **WHEN** 用户选择JPG格式进行图片生成
- **THEN** 系统应生成压缩的JPG图片文件

### Requirement: 图片质量选择
系统 SHALL 支持多种图片质量选择

#### Scenario: 原始大小图片生成
- **WHEN** 用户选择"原始大小"质量进行图片生成
- **THEN** 系统应使用PPT原始分辨率生成图片

#### Scenario: 放大图片生成
- **WHEN** 用户选择放大倍数（1.5倍、2倍、3倍、4倍）进行图片生成
- **THEN** 系统应使用相应倍数的分辨率生成图片

### Requirement: 图片文件命名
系统 SHALL 为生成的图片文件提供清晰的命名规则

#### Scenario: 图片文件命名
- **WHEN** 系统生成图片文件
- **THEN** 文件名格式应为"output_{index}.{format}"

### Requirement: 图片生成进度跟踪
系统 SHALL 在生成图片过程中提供实时进度反馈

#### Scenario: 图片生成进度显示
- **WHEN** 系统正在生成图片文件
- **THEN** 进度条应显示当前生成进度和状态信息

### Requirement: 直接生成到指定目录
系统 SHALL 将图片文件直接生成到用户指定的输出目录

#### Scenario: 无子目录生成
- **WHEN** 系统生成图片文件
- **THEN** 图片文件应直接出现在指定目录中，不创建额外的子目录

### Requirement: GUI界面集成
系统 SHALL 在GUI界面中提供图片生成选项

#### Scenario: 图片生成选项显示
- **WHEN** 用户勾选"直接生成图片"复选框
- **THEN** 系统应显示图片格式和质量选择选项

#### Scenario: 质量选择同步
- **WHEN** 用户启用或禁用图片生成选项
- **THEN** 相关的格式和质量选择菜单应相应显示或隐藏

## Technical Notes

### COM接口依赖
- 使用Microsoft PowerPoint COM接口进行图片转换
- 需要系统安装PowerPoint应用程序
- 通过临时修改幻灯片尺寸实现质量调整

### 质量实现机制
- 通过修改`PageSetup.SlideWidth`和`PageSetup.SlideHeight`属性调整分辨率
- 生成完成后自动恢复原始尺寸
- 支持1.0x（原始）、1.5x、2.0x、3.0x、4.0x缩放比例

### 文件处理
- 使用临时目录处理PowerPoint自动创建的子目录结构
- 递归查找所有生成的图片文件
- 移动文件到用户指定的最终位置

## Implementation Status

### 已实现功能
- ✅ PNG和JPG格式支持
- ✅ 质量选择（原始大小、1.5倍、2倍、3倍、4倍）
- ✅ GUI界面集成
- ✅ 进度跟踪和错误处理
- ✅ 直接输出到指定目录
- ✅ 与现有PPT生成功能完全兼容

### 测试验证
- ✅ 基本图片生成功能测试通过
- ✅ 不同质量设置效果验证通过
- ✅ 文件大小随质量设置正确变化
- ✅ GUI界面功能正常工作