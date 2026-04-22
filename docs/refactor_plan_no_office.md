# PPT生成工具重构计划：移除 Office/WPS 依赖

## 目标
将项目从依赖 Microsoft Office/WPS COM 接口重构为仅使用 `python-pptx` 库，实现纯 Python 方案。

## 现状分析

### 当前架构
```
PPTGenerator
├── OfficeSuiteManager (检测 Office/WPS)
├── OfficeSuiteFactory (创建 COM 处理器)
│   ├── MicrosoftProcessor (COM)
│   └── WPSProcessor (COM)
├── PPTXProcessor (python-pptx) - 已存在但未使用
└── TemplateProcessor (旧版 COM)
```

### 需要处理的核心问题
1. **占位符识别**：当前使用 COM 接口读取"选择窗格"中的形状名称
2. **模板处理**：COM 处理器复杂且依赖外部软件
3. **代码冗余**：大量代码用于处理 COM 连接、错误恢复、应用生命周期

## 新架构设计

### 简化后的架构
```
PPTGenerator
├── PPTXTemplateProcessor (基于 python-pptx)
│   ├── 使用 alt_text 作为占位符标识
│   ├── 支持形状名称作为备选
│   └── 纯 Python，无需外部依赖
├── DataReader (pandas - 已纯Python)
└── ExactMatcher (纯Python - 无需修改)
```

## 占位符方案

### 主要方案：Alt Text（替代文本）
- **用户操作**：右键形状 → "编辑替代文本" → 输入 `ph_姓名`
- **代码读取**：`shape.alt_text`
- **优点**：完全隐藏，不影响预览，python-pptx原生支持

### 备选方案：形状名称
- **用户操作**：选择窗格中重命名形状
- **代码读取**：`shape.name`
- **限制**：在PowerPoint中编辑需要选择窗格（但只需设置一次）

### 混合策略
1. 优先检查 `alt_text`（以 `ph_` 开头）
2. 其次检查 `shape.name`（以 `ph_` 开头）
3. 最后检查文本内容（如 `{{姓名}}` 格式）

## 代码重构策略

### 策略：保留并重构（不是删除）

**原因**：
1. 保留向后兼容性（万一用户需要COM功能）
2. 保留代码历史价值
3. 避免激进重构带来的风险

**具体做法**：
1. **保留** COM 处理器代码，但标记为 `@deprecated`
2. **新增** `PPTXTemplateProcessor` 作为主处理器
3. **修改** `PPTGenerator` 默认使用 `PPTXTemplateProcessor`
4. **简化** GUI，移除 Office 检测相关 UI

## 文件变更计划

### 1. 核心处理器重构

#### 修改 `src/core/template_processor.py`
- **保留** `TemplateProcessor` (COM版本) - 添加 `@deprecated` 标记
- **保留** `PPTXProcessor` - 大幅增强功能
  - 添加 `find_placeholders_by_alt_text()` 方法
  - 添加 `find_placeholders_by_name()` 方法
  - 添加 `replace_placeholders()` 方法
  - 支持保持文本格式（只替换文本，不丢失样式）

#### 新增 `src/core/processors/pptx_processor.py`
- 将增强后的 `PPTXProcessor` 独立成文件
- 实现完整的 `ITemplateProcessor` 接口

### 2. 工厂和接口调整

#### 修改 `src/core/factory/office_suite_factory.py`
- **保留** 原有工厂逻辑
- **新增** `create_pptx_processor()` 方法
- **修改** `create_auto_processor()` 默认返回 PPTX 处理器

#### 修改 `src/core/interfaces/template_processor.py`
- **保留** 接口定义（无需修改）

### 3. PPT生成器调整

#### 修改 `src/ppt_generator.py`
- **简化** `_initialize_template_processor()`
- **移除** `use_com_interface` 参数（或设为 False 默认）
- **移除** `office_manager` 依赖
- **保留** 但标记为可选的 COM 相关方法

### 4. GUI 简化

#### 修改 `src/gui/main_window.py`
- **移除** Office 套件选择下拉框
- **移除** Office 检测相关代码
- **简化** 初始化流程

#### 修改 `src/gui/ppt_worker_thread.py`
- **简化** 处理器初始化

### 5. 配置管理调整

#### 修改 `src/config/office_cache.py`
- **保留** 但标记为可选/遗留
- **简化** 配置项，移除 Office 相关配置

#### 修改 `src/config/gui_config.py`
- **移除** Office 套件选择配置

### 6. 保留但不主动使用的文件

以下文件**保留但不再主动调用**：
- `src/core/processors/microsoft_processor.py` - 标记 `@deprecated`
- `src/core/processors/wps_processor.py` - 标记 `@deprecated`
- `src/gui/office_suite_detector.py` - 标记 `@deprecated`
- `src/config/office_cache.py` - 标记为遗留功能

## 实施步骤

### Phase 1: 备份与准备
1. 备份所有关键文件
2. 创建测试用例验证当前功能
3. 准备示例模板（使用 alt_text）

### Phase 2: 核心重构
1. 增强 `PPTXProcessor` 类
2. 修改 `PPTGenerator` 默认使用 PPTX 处理器
3. 更新占位符查找逻辑（alt_text + name）

### Phase 3: GUI 简化
1. 移除 Office 检测 UI
2. 简化初始化流程
3. 更新用户提示信息

### Phase 4: 测试与验证
1. 测试模板加载
2. 测试占位符识别（alt_text）
3. 测试文本替换
4. 测试批量生成

### Phase 5: 文档更新
1. 更新 README
2. 添加模板制作指南
3. 添加迁移指南（从 COM 到 python-pptx）

## 风险与应对

| 风险 | 影响 | 应对措施 |
|------|------|----------|
| 某些复杂PPT格式不支持 | 中 | 保留COM处理器作为备选 |
| 用户已有模板使用选择窗格命名 | 高 | 提供迁移工具或同时支持name和alt_text |
| 性能下降 | 低 | python-pptx通常比COM更快 |
| 功能缺失 | 中 | 分阶段测试，确保核心功能完整 |

## 预期收益

1. **无需安装 Office/WPS** - 降低用户使用门槛
2. **启动速度提升** - 无需检测和连接外部应用
3. **稳定性提升** - 消除 COM 接口的不稳定性
4. **部署简化** - 纯 Python 依赖，打包更简单
5. **跨平台潜力** - 未来可支持 Linux/Mac

## 回滚计划

如果重构出现问题：
1. 从备份恢复原始文件
2. 或通过配置切换回 COM 模式（保留的代码）
