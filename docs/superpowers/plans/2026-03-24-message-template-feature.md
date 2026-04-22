# 贺语模板功能实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为PPT批量生成工具添加贺语模板功能，允许用户输入包含占位符的贺语模板，在生成贺报时同步生成同名的txt贺语文件

**Architecture:** 
1. 在GUI右侧面板替换"日志信息"区域为"贺语模板"输入框
2. 使用ConfigManager保存/加载贺语模板
3. 在批量生成流程中集成贺语生成逻辑，将`{{表头}}`格式的占位符替换为实际数据
4. 生成与贺报同目录、同名的txt文件

**Tech Stack:** Python, PySide6, PyYAML

---

## 文件结构

### 需要修改的文件：

1. **`src/gui/main_window.py`** - 修改UI布局，替换日志区域为贺语模板输入框，添加保存/加载逻辑
2. **`src/gui/ppt_worker_thread.py`** - 添加贺语模板传递和生成触发
3. **`src/ppt_generator.py`** - 实现贺语生成核心逻辑
4. **`src/config_manager.py`** - 添加贺语模板的配置保存/加载方法

---

## Task 1: 修改 ConfigManager 添加贺语模板配置支持

**Files:**
- Modify: `src/config_manager.py`

- [ ] **Step 1: 添加贺语模板配置键名常量**

在 `ConfigManager` 类中添加：
```python
MESSAGE_TEMPLATE_KEY = "message_template"
```

- [ ] **Step 2: 添加获取贺语模板方法**

```python
def get_message_template(self) -> str:
    """获取保存的贺语模板"""
    return self.config.get(self.MESSAGE_TEMPLATE_KEY, "")
```

- [ ] **Step 3: 添加保存贺语模板方法**

```python
def save_message_template(self, template: str) -> bool:
    """保存贺语模板到配置"""
    try:
        self.config[self.MESSAGE_TEMPLATE_KEY] = template
        return self.save_config()
    except Exception as e:
        logger.error(f"保存贺语模板失败: {str(e)}")
        return False
```

- [ ] **Step 4: 测试配置方法**

创建一个简单的测试脚本验证配置读写：
```python
from src.config_manager import ConfigManager

config = ConfigManager()
config.save_message_template("恭喜{{姓名}}完成业绩！")
result = config.get_message_template()
assert result == "恭喜{{姓名}}完成业绩！"
print("测试通过")
```

---

## Task 2: 修改 MainWindow 替换日志区域为贺语模板输入框

**Files:**
- Modify: `src/gui/main_window.py`

- [ ] **Step 1: 修改 create_right_panel 方法**

找到 `create_right_panel` 方法，将日志区域替换为贺语模板区域：

**原代码（约300-330行）：**
```python
# 错误日志组
log_group = QGroupBox("日志信息")
log_group.setStyleSheet(PanelStyles.group_box_style())
log_layout = QVBoxLayout()

self.log_text = QTextEdit()
self.log_text.setReadOnly(True)
self.log_text.setMaximumHeight(200)
self.log_text.setStyleSheet(LogStyles.text_edit_style())
log_layout.addWidget(self.log_text)

log_group.setLayout(log_layout)
layout.addWidget(log_group)
```

**替换为：**
```python
# 贺语模板组
message_group = QGroupBox("贺语模板")
message_group.setStyleSheet(PanelStyles.group_box_style())
message_layout = QVBoxLayout()

# 提示标签
hint_label = QLabel("支持 {{表头}} 格式的占位符，例如：{{姓名}}、{{保费}}")
hint_label.setFont(Fonts.small_font())
hint_label.setStyleSheet("color: #666; margin-bottom: 5px;")
message_layout.addWidget(hint_label)

# 贺语模板输入框
self.message_template_edit = QTextEdit()
self.message_template_edit.setPlaceholderText("在此输入贺语模板...\n\n例如：\n┏━🥇团险新单捷报🥇━┓\n  ㊗㊗ 热烈祝贺 ㊗㊗\n  {{机构}} {{姓名}}\n  承保 企业优选 新单\n  保费 {{保费}}元")
self.message_template_edit.setMaximumHeight(200)
self.message_template_edit.setStyleSheet(LogStyles.text_edit_style())
self.message_template_edit.textChanged.connect(self.on_message_template_changed)
message_layout.addWidget(self.message_template_edit)

message_group.setLayout(message_layout)
layout.addWidget(message_group)
```

- [ ] **Step 2: 添加贺语模板变化处理方法**

在 `MainWindow` 类中添加：

```python
def on_message_template_changed(self):
    """贺语模板内容变化时自动保存"""
    # 使用防抖机制保存
    self._trigger_settings_save_debounced()
```

- [ ] **Step 3: 修改防抖保存方法以包含贺语模板**

找到 `_save_settings_debounced` 方法，添加贺语模板保存：

```python
def _save_settings_debounced(self):
    """防抖保存设置（包括图片生成设置和贺语模板）"""
    try:
        # 保存图片生成设置
        generate_images = self.direct_image_checkbox.isChecked()
        image_format = self.image_format_combo.currentText() if generate_images else "PNG"
        image_quality = self.image_quality_combo.currentText() if generate_images else "原始大小"
        
        self.config_manager.save_image_generation_settings(
            generate_images, image_format, image_quality
        )
        
        # 保存贺语模板
        if hasattr(self, 'message_template_edit'):
            template = self.message_template_edit.toPlainText()
            self.config_manager.save_message_template(template)
            
    except Exception as e:
        logger.error(f"防抖保存设置失败: {str(e)}")
```

- [ ] **Step 4: 添加贺语模板加载方法**

在 `init_ui` 方法的最后（在 `load_image_generation_settings()` 之后）添加：

```python
# 加载保存的贺语模板
self.load_message_template()
```

然后添加加载方法：

```python
def load_message_template(self):
    """加载保存的贺语模板"""
    try:
        template = self.config_manager.get_message_template()
        if template and hasattr(self, 'message_template_edit'):
            self.message_template_edit.setPlainText(template)
            logger.info("已加载保存的贺语模板")
    except Exception as e:
        logger.error(f"加载贺语模板失败: {str(e)}")
```

- [ ] **Step 5: 保留日志功能但改为内部使用**

由于移除了日志显示区域，需要将日志输出改为打印到控制台或文件。找到 `append_log` 方法修改为：

```python
def append_log(self, message):
    """添加日志消息（输出到控制台）"""
    print(f"[LOG] {message}")
    logger.info(message)
```

---

## Task 3: 修改 PPTWorkerThread 添加贺语模板传递

**Files:**
- Modify: `src/gui/ppt_worker_thread.py`

- [ ] **Step 1: 添加贺语模板属性**

在 `PPTWorkerThread` 类的 `__init__` 方法中添加：

```python
self.message_template = ""  # 贺语模板
```

- [ ] **Step 2: 添加设置贺语模板方法**

```python
def set_message_template(self, template: str):
    """设置贺语模板"""
    self.message_template = template
```

- [ ] **Step 3: 修改批量生成方法传递贺语模板**

找到 `start_batch_generate` 或相应的生成方法，确保贺语模板被传递给 `PPTGenerator`。

在调用 `ppt_generator.batch_generate` 之前添加：

```python
# 设置贺语模板
if self.message_template:
    self.ppt_generator.set_message_template(self.message_template)
```

---

## Task 4: 修改 PPTGenerator 实现贺语生成核心逻辑

**Files:**
- Modify: `src/ppt_generator.py`

- [ ] **Step 1: 添加贺语模板属性**

在 `PPTGenerator` 类的 `__init__` 方法中添加：

```python
self.message_template = ""  # 贺语模板
```

- [ ] **Step 2: 添加设置贺语模板方法**

```python
def set_message_template(self, template: str):
    """
    设置贺语模板
    
    Args:
        template: 贺语模板文本，支持 {{表头}} 格式的占位符
    """
    self.message_template = template
    logger.info(f"已设置贺语模板，长度: {len(template)} 字符")
```

- [ ] **Step 3: 添加贺语生成方法**

```python
def generate_message(self, row_data: pd.Series) -> str:
    """
    根据模板和数据生成贺语
    
    Args:
        row_data: 一行数据（pandas Series）
        
    Returns:
        str: 生成的贺语文本
    """
    if not self.message_template:
        return ""
    
    message = self.message_template
    
    # 使用正则表达式查找所有 {{表头}} 格式的占位符
    import re
    placeholders = re.findall(r'\{\{(.*?)\}\}', message)
    
    # 替换每个占位符
    for placeholder in placeholders:
        column_name = placeholder.strip()
        if column_name in row_data:
            value = str(row_data[column_name])
            message = message.replace(f'{{{{{column_name}}}}}', value)
        else:
            logger.warning(f"贺语模板中的占位符 '{{{{{column_name}}}}}' 未在数据中找到对应列")
            # 保留原占位符，或替换为空字符串
            message = message.replace(f'{{{{{column_name}}}}}', '')
    
    return message
```

- [ ] **Step 4: 添加保存贺语文件方法**

```python
def save_message_file(self, message: str, output_path: str) -> bool:
    """
    保存贺语到txt文件
    
    Args:
        message: 贺语文本
        output_path: 输出文件路径（与贺报同名，扩展名为.txt）
        
    Returns:
        bool: 保存成功返回True
    """
    try:
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        # 写入文件（UTF-8编码）
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(message)
        
        logger.info(f"贺语文件已保存: {output_path}")
        return True
        
    except Exception as e:
        logger.error(f"保存贺语文件失败 {output_path}: {str(e)}")
        return False
```

- [ ] **Step 5: 修改 batch_generate 方法集成贺语生成**

找到 `batch_generate` 方法，在生成每个文件的循环中添加贺语生成逻辑。

在循环中，生成文件名后添加：

```python
# 生成贺语文件（如果设置了贺语模板）
if self.message_template:
    try:
        message = self.generate_message(row)
        if message:
            # 构建贺语文件路径（与输出文件同名，扩展名为.txt）
            message_filename = f"{filename}.txt"
            message_output_path = os.path.join(output_dir, message_filename)
            self.save_message_file(message, message_output_path)
    except Exception as e:
        logger.error(f"生成贺语文件时出错: {str(e)}")
```

**注意：** 需要根据 `generate_images` 标志调整输出文件名逻辑：
- 如果生成图片：`filename` 已经包含扩展名（如 `output1.jpg`）
- 需要去掉扩展名再添加 `.txt`

修正后的文件名处理：

```python
# 生成贺语文件（如果设置了贺语模板）
if self.message_template:
    try:
        message = self.generate_message(row)
        if message:
            # 构建贺语文件路径（去掉原有扩展名，添加.txt）
            base_filename = os.path.splitext(filename)[0]
            message_filename = f"{base_filename}.txt"
            message_output_path = os.path.join(output_dir, message_filename)
            self.save_message_file(message, message_output_path)
    except Exception as e:
        logger.error(f"生成贺语文件时出错: {str(e)}")
```

---

## Task 5: 修改 MainWindow 批量生成时传递贺语模板

**Files:**
- Modify: `src/gui/main_window.py`

- [ ] **Step 1: 在 batch_generate 方法中传递贺语模板**

找到 `batch_generate` 方法，在启动工作线程之前添加：

```python
# 设置贺语模板
if hasattr(self, 'message_template_edit'):
    message_template = self.message_template_edit.toPlainText().strip()
    self.worker_thread.set_message_template(message_template)
    if message_template:
        self.append_log("已启用贺语生成功能")
```

---

## Task 6: 测试功能

**Files:**
- Test: 手动测试

- [ ] **Step 1: 运行程序验证UI**

运行程序，检查：
1. 右侧面板显示"贺语模板"区域而不是"日志信息"
2. 文本输入框可以正常输入
3. 提示标签显示正确

- [ ] **Step 2: 测试贺语模板保存/加载**

1. 输入贺语模板内容
2. 关闭程序
3. 重新打开程序，检查内容是否自动加载

- [ ] **Step 3: 测试贺语生成功能**

1. 准备测试数据（Excel文件，包含"姓名"、"保费"等列）
2. 输入贺语模板：
   ```
   恭喜{{姓名}}完成{{保费}}元业绩！
   ```
3. 执行批量生成
4. 检查输出目录是否生成了与贺报同名的txt文件
5. 检查txt文件内容是否正确替换了占位符

- [ ] **Step 4: 测试边界情况**

1. 空贺语模板 - 应该不生成txt文件
2. 贺语模板中包含不存在的数据列 - 应该替换为空字符串并记录警告
3. 特殊字符和表情符号 - 应该正确保存到txt文件

---

## 提交清单

所有任务完成后，提交代码：

```bash
git add src/gui/main_window.py src/gui/ppt_worker_thread.py src/ppt_generator.py src/config_manager.py
git commit -m "feat: 添加贺语模板功能

- 替换日志区域为贺语模板输入框
- 支持 {{表头}} 格式的占位符
- 批量生成时同步生成同名txt贺语文件
- 自动保存/加载贺语模板配置"
```
