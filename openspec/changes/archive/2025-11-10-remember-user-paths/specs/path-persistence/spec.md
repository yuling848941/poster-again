# Path Persistence Spec

## ADDED Requirements

### Requirement: Automatic Path Saving
- **系统 SHALL 在用户选择文件/目录后自动保存路径到配置文件**
- **系统 SHALL 确保保存操作不阻塞用户界面**
- **系统 SHALL 只保存有效的路径**

#### Scenario: 用户选择模板文件
```
Given 用户点击"浏览..."按钮选择模板文件
When 用户选择了有效的PPT模板文件
Then 系统应自动将文件路径保存到config.yaml的paths.last_template_dir
And 文件选择器的起始目录应更新为该文件所在目录
```

#### Scenario: 用户选择数据源文件
```
Given 用户点击"浏览..."按钮选择数据源文件
When 用户选择了有效的Excel/CSV/JSON数据文件
Then 系统应自动将文件路径保存到config.yaml的paths.last_data_dir
And 文件选择器的起始目录应更新为该文件所在目录
```

#### Scenario: 用户选择输出目录
```
Given 用户点击"浏览..."按钮选择输出目录
When 用户选择了有效的输出目录
Then 系统应自动将目录路径保存到config.yaml的paths.last_output_dir
And 目录选择器的起始目录应更新为该目录
```

### Requirement: Startup Path Restoration
- **程序 SHALL 在启动时自动恢复上次保存的有效路径**
- **恢复的路径 SHALL 显示在对应的输入框中**
- **系统 SHALL 只恢复存在的路径到界面**

#### Scenario: 程序启动时恢复有效路径
```
Given 程序启动
And config.yaml中包含上次保存的有效路径
When MainWindow初始化完成
Then 系统应将模板文件路径恢复到模板文件输入框
And 系统应将数据源文件路径恢复到数据源文件输入框
And 系统应将输出目录路径恢复到输出目录输入框
```

#### Scenario: 路径不存在时的处理
```
Given 程序启动
And config.yaml中包含上次保存的路径
When 路径对应的文件/目录不存在
Then 系统应跳过该路径的恢复
And 对应的输入框保持为空
And 系统应在日志中记录警告信息
```

### Requirement: Smart Default Directory
- **文件选择器 SHALL 从上次使用的相关目录开始**
- **系统 SHALL 为不同类型的文件使用不同的起始目录**
- **系统 SHALL 在没有保存目录时使用用户主目录作为默认**

#### Scenario: 模板文件选择器的智能起始目录
```
Given 用户点击模板文件"浏览..."按钮
When config.yaml中保存了last_template_dir
Then 文件选择器应从last_template_dir对应的目录开始
```

#### Scenario: 数据文件选择器的智能起始目录
```
Given 用户点击数据文件"浏览..."按钮
When config.yaml中保存了last_data_dir
Then 文件选择器应从last_data_dir对应的目录开始
```

#### Scenario: 输出目录选择器的智能起始目录
```
Given 用户点击输出目录"浏览..."按钮
When config.yaml中保存了last_output_dir
Then 目录选择器应从last_output_dir开始
```

### Requirement: Path Validation
- **系统 SHALL 验证保存的路径有效性**
- **路径验证 SHALL 包括文件存在性检查**
- **系统 SHALL 在路径验证失败时优雅降级**

#### Scenario: 路径有效性验证
```
Given 系统准备恢复保存的路径
When 路径指向的文件/目录存在
Then 系统应将该路径恢复到界面
When 路径指向的文件/目录不存在
Then 系统应跳过该路径
And 在日志中记录路径无效的警告
```

### Requirement: Configuration Integration
- **路径持久化功能 SHALL 与现有ConfigManager无缝集成**
- **系统 SHALL 不破坏现有的配置结构和功能**
- **系统 SHALL 支持配置的向后兼容性**

#### Scenario: 配置文件兼容性
```
Given 现有的config.yaml配置文件
When 启用路径持久化功能
Then 系统应能正确读取现有的paths配置
And 不应影响其他配置项的使用
```

#### Scenario: 新配置项的默认值
```
Given 新安装的程序或重置的配置
When 用户首次使用程序
Then paths相关的配置项应为空字符串
And 系统应正常工作，不显示错误信息
```