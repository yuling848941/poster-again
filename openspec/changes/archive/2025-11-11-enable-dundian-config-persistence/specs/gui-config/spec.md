## ADDED Requirements

### Requirement: 递交趸期映射配置保存
系统SHALL支持保存用户设置的递交趸期数据映射关系，即记录哪个占位符与"期趸数据"列的关联。

#### Scenario: 用户设置递交趸期数据映射
- **WHEN** 用户右键点击占位符表格中的某一行
- **AND** 选择"递交趸期数据"选项
- **THEN** 系统记录该占位符与"期趸数据"的映射关系
- **AND** 系统更新匹配表格中的显示
- **AND** 系统保存映射关系到主窗口状态

#### Scenario: 配置保存包含递交趸期映射
- **WHEN** 用户点击"保存配置"按钮
- **AND** 用户已设置递交趸期数据映射
- **THEN** 系统将递交趸期映射关系保存到配置文件的dundian_mappings部分
- **AND** 配置文件格式包含占位符名称和对应的数据列
- **AND** 系统确认配置保存成功

#### Scenario: 配置文件缺少递交趸期映射
- **WHEN** 用户加载的配置文件不包含dundian_mappings
- **THEN** 系统使用空的dundian_mappings对象
- **AND** 系统正常加载其他配置项
- **AND** 不影响应用程序的正常功能

### Requirement: 递交趸期映射配置加载
系统SHALL支持从配置文件加载递交趸期数据映射关系，并根据占位符名称匹配决定是否应用。

#### Scenario: 配置加载包含递交趸期映射
- **WHEN** 用户加载包含dundian_mappings的配置文件
- **AND** 配置中的占位符在当前模板中存在
- **THEN** 系统恢复该占位符与"期趸数据"的映射关系
- **AND** 系统更新匹配表格显示该映射关系
- **AND** 系统记录加载成功的映射关系

#### Scenario: 占位符名称不匹配
- **WHEN** 用户加载的配置文件包含dundian_mappings
- **AND** 配置中的占位符在当前模板中不存在
- **THEN** 系统跳过该占位符的映射恢复
- **AND** 系统继续处理其他配置项
- **AND** 系统在日志中记录跳过的占位符

#### Scenario: 部分占位符匹配
- **WHEN** 用户加载的配置文件包含多个递交趸期映射
- **AND** 部分占位符在当前模板中存在
- **THEN** 系统只恢复匹配成功的占位符映射
- **AND** 系统在日志中显示匹配成功的数量
- **AND** 系统在日志中显示跳过的占位符

## MODIFIED Requirements

### Requirement: 配置有效性检查
现有的_has_valid_settings方法SHALL扩展以支持递交趸期映射配置的有效性检查。

#### Scenario: 配置有效性检查
- **WHEN** 系统检查是否有有效的设置可以保存
- **THEN** 系统检查text_additions是否有内容
- **AND** 系统检查dundian_mappings是否有内容
- **AND** 任意一个有内容就认为配置有效
- **AND** 两个都为空时提示没有有效设置

#### Scenario: 配置文件结构扩展
- **WHEN** 系统保存配置文件
- **THEN** 配置文件包含text_additions和dundian_mappings两个部分
- **AND** dundian_mappings结构为 {"占位符名称": "期趸数据"}
- **AND** 配置文件保持向后兼容性
- **AND** 旧版本配置文件能正常加载新格式

#### Scenario: SimpleConfigManager功能扩展
- **WHEN** SimpleConfigManager收集GUI状态时
- **THEN** _collect_gui_state方法收集dundian_mappings
- **AND** _restore_gui_state方法恢复dundian_mappings
- **AND** 恢复时检查占位符是否存在
- **AND** 只恢复匹配成功的占位符映射