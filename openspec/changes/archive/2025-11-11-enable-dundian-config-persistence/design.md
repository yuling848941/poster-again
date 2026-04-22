# 递交趸期配置持久化设计文档

## 问题分析

### 当前状况
1. **递交趸期功能存在**：用户可以右键点击占位符表格，选择"递交趸期数据"功能
2. **配置不支持**：这个功能的设置信息无法保存到配置文件中
3. **需要重新操作**：每次使用应用程序都需要重新设置递交趸期映射

### 用户痛点
- 无法保存哪个占位符与"期趸数据"列的映射关系
- 每次加载模板和数据后都需要重新操作递交趸期数据功能
- 配置管理功能不完整

## 解决方案设计

### 1. 简单的配置扩展

#### 记录映射关系
```python
# 在主窗口中记录递交趸期映射
if not hasattr(main_window, 'dundian_mappings'):
    main_window.dundian_mappings = {}

# 记录哪个占位符映射到期趸数据
main_window.dundian_mappings[placeholder_name] = "期趸数据"
```

#### 配置文件格式
```json
{
  "gui_settings": {
    "text_additions": {
      "ph_保险种类": {"suffix": "1件"}
    },
    "dundian_mappings": {
      "ph_保费金额": "期趸数据"
    }
  }
}
```

### 2. 配置管理集成

#### 扩展SimpleConfigManager
```python
def _collect_gui_state(self, main_window):
    """收集主窗口的GUI状态"""
    gui_config = {
        "text_additions": {},
        "dundian_mappings": {}  # 新增递交趸期映射
    }

    # 收集递交趸期映射
    if hasattr(main_window, 'dundian_mappings'):
        gui_config["dundian_mappings"] = main_window.dundian_mappings.copy()

    return gui_config

def _restore_gui_state(self, main_window, gui_settings):
    """恢复GUI状态"""
    # 恢复递交趸期映射
    dundian_mappings = gui_settings.get('dundian_mappings', {})
    if dundian_mappings:
        if not hasattr(main_window, 'dundian_mappings'):
            main_window.dundian_mappings = {}

        for placeholder, data_column in dundian_mappings.items():
            # 只有在占位符名称匹配时才恢复
            if self._placeholder_exists(main_window, placeholder):
                main_window.dundian_mappings[placeholder] = data_column
                # 更新匹配表格显示
                self._update_match_table_for_dundian(main_window, placeholder, data_column)
```

### 3. 数据流设计

```
用户右键选择递交趸期数据 → 记录映射关系 → 保存到配置文件
                      ↑                                      ↓
配置文件 ←─── SimpleConfigManager ←─── 应用配置恢复映射关系
```

#### 核心逻辑
- **保存时**：记录用户设置的占位符 → 期趸数据映射
- **加载时**：检查配置中的占位符是否存在，存在则恢复映射
- **匹配机制**：只有占位符名称完全匹配时才应用设置

### 4. 实现策略

#### 阶段性实施
1. **阶段1**：添加基础GUI控件
2. **阶段2**：集成配置管理功能
3. **阶段3**：用户体验优化

#### 风险控制
- 新增功能不影响现有功能
- 保持配置文件的向后兼容性
- 提供清晰的错误处理和用户提示

### 5. 技术实现细节

#### 控件状态管理
```python
def on_dundian_checkbox_changed(self, state):
    """处理复选框状态变化"""
    enabled = state == Qt.Checked
    self.dundian_type_combo.setEnabled(enabled)
    # 可以添加状态变化日志

def on_dundian_type_changed(self, type_text):
    """处理类型选择变化"""
    # 可以添加选择变化日志
```

#### 配置同步逻辑
```python
# 在配置管理器中确保正确的状态同步
def _collect_dundian_settings(self, main_window):
    """收集递交趸期设置"""
    return {
        "enabled": main_window.dundian_checkbox.isChecked(),
        "default_type": main_window.dundian_type_combo.currentText()
    }
```

### 6. 用户体验考虑

#### 可见性
- 控件应该放置在用户容易发现的位置
- 提供清晰的标签和说明文字
- 使用工具提示提供额外信息

#### 一致性
- 与现有的图片生成设置保持相同的视觉风格
- 使用一致的交互模式和响应方式
- 保持配置保存/加载流程的一致性

#### 反馈机制
- 设置变化时提供即时反馈
- 配置保存/加载时显示状态信息
- 错误情况下提供清晰的错误提示

### 7. 测试策略

#### 单元测试
- GUI控件的功能测试
- 配置管理器的集成测试
- 数据序列化/反序列化测试

#### 集成测试
- 完整的保存/加载流程测试
- 配置文件兼容性测试
- 用户操作场景测试

#### 用户体验测试
- 界面易用性测试
- 功能发现性测试
- 错误处理测试

## 成功标准

### 功能完整性
- ✅ 用户可以通过GUI控件设置递交趸期选项
- ✅ 设置能够正确保存到配置文件
- ✅ 配置能够正确加载并恢复界面状态
- ✅ 配置文件保持向后兼容性

### 用户体验
- ✅ 操作流程直观易懂
- ✅ 设置变化提供即时反馈
- ✅ 与现有配置管理功能保持一致
- ✅ 错误情况下有清晰的提示信息

### 技术质量
- ✅ 代码结构清晰，易于维护
- ✅ 配置管理逻辑正确可靠
- ✅ 不影响现有功能的性能
- ✅ 异常处理完善