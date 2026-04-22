# 记住用户路径和设置选择功能 - 设计文档

## 架构分析

### 当前系统架构

```
MainWindow (GUI界面)
    ↓ 文件选择器 + 图片生成设置
ConfigManager (配置管理)
    ↓ 配置保存/加载
config.yaml (配置文件)
```

### 设计原则

1. **最小侵入性**：基于现有ConfigManager扩展，不破坏现有功能
2. **用户友好**：路径失效时优雅降级，不影响程序使用
3. **数据安全**：路径文件存在性验证，避免无效路径
4. **设置持久化**：图片生成设置的状态保存和恢复

### 组件设计

#### ConfigManager增强

**新增方法：**
- `load_last_paths()` - 加载上次保存的路径
- `get_start_dir_for_file_type(file_type)` - 获取文件类型对应的起始目录
- `validate_path(path)` - 验证路径有效性
- `load_image_generation_settings()` - 加载图片生成设置
- `update_image_generation_settings()` - 保存图片生成设置

**修改现有方法：**
- `update_last_paths()` - 增强参数验证和错误处理

#### MainWindow增强

**新增方法：**
- `load_last_paths()` - 界面初始化时加载路径
- `restore_path_to_edit(path, edit_widget, setter)` - 恢复路径到界面控件
- `load_image_generation_settings()` - 加载图片生成设置
- `save_image_generation_settings()` - 保存图片生成设置

**修改现有方法：**
- `browse_template_file()` - 增加路径保存和智能起始目录
- `browse_data_file()` - 增加路径保存和智能起始目录
- `browse_output_dir()` - 增加路径保存和智能起始目录
- 图片生成相关控件的信号处理 - 添加设置自动保存

### 数据流设计

#### 程序启动时
```
程序启动 → MainWindow初始化 → load_last_paths() →
验证路径有效性 → 恢复有效路径到界面 →
load_image_generation_settings() → 恢复图片生成设置
```

#### 用户选择文件时
```
用户选择文件 → 验证路径 → 更新界面 →
保存路径到ConfigManager → 自动保存到config.yaml
```

#### 用户修改图片设置时
```
用户修改设置 → 信号触发 → save_image_generation_settings() →
ConfigManager保存到config.yaml
```

#### 下次程序启动时
```
程序启动 → 从config.yaml加载路径 →
设置文件选择器起始目录为上次路径 →
从config.yaml加载图片生成设置 → 恢复界面状态
```

### 错误处理策略

1. **路径不存在**：跳过恢复，使用默认值
2. **配置文件损坏**：回退到默认配置
3. **权限问题**：记录警告，继续使用默认路径
4. **图片设置无效**：使用默认设置（PNG格式，原始大小）

### 配置结构

扩展config.yaml结构，添加图片生成设置：
```yaml
paths:
  last_template_dir: "E:/path/to/template.pptx"
  last_data_dir: "E:/path/to/data.xlsx"
  last_output_dir: "E:/path/to/output"
  default_output_dir: "./output"

image_generation:
  enabled: true
  format: "PNG"  # 或 "JPG"
  quality: 1.0   # 1.0, 1.5, 2.0, 3.0, 4.0
```

### 测试策略

1. **单元测试**：ConfigManager新增方法的测试
2. **集成测试**：MainWindow路径和设置加载/保存流程测试
3. **用户场景测试**：
   - 正常路径恢复
   - 路径不存在时的处理
   - 配置文件损坏时的处理
   - 图片生成设置的保存和恢复
   - 图片设置无效时的处理

### 图片生成控件分析

根据MainWindow代码分析，图片生成相关控件：
- `generate_images_checkbox` - QCheckBox，勾选启用图片生成
- `image_format_combo` - QComboBox，选择格式（PNG/JPG）
- `image_quality_combo` - QComboBox，选择质量（原始大小, 1.5倍, 2倍, 3倍, 4倍）

这些控件的状态需要在程序启动时恢复，并在用户修改时自动保存。