## MODIFIED Requirements
### Requirement: PPT批量生成
系统 SHALL 支持批量生成PPT文件，并为每行数据生成一个PPT文件

#### Scenario: 直接生成图片模式
- **WHEN** 用户勾选"直接生成图片"复选框并选择图片格式
- **THEN** 系统应将生成的PPT转换为指定格式的图片文件

## ADDED Requirements
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

### Requirement: 图片生成进度跟踪
系统 SHALL 在生成图片过程中提供实时进度反馈

#### Scenario: 图片生成进度显示
- **WHEN** 系统正在生成图片文件
- **THEN** 进度条应显示当前生成进度和状态信息

### Requirement: 图片文件命名
系统 SHALL 为生成的图片文件提供清晰的命名规则

#### Scenario: 图片文件命名
- **WHEN** 系统生成图片文件
- **THEN** 文件名格式应为"output_{index}.{format}"