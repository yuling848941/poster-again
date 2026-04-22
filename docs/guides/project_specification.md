# PPT批量生成工具 - 项目规范文档

## 1. 项目概述

### 1.1 项目名称
PPT批量生成工具 (Poster Generator)

### 1.2 项目目标
开发一个基于PySide6的自动化工具，可以根据Excel数据批量生成个性化的PPT，同时保持PPT的复杂特效。


## 2. 功能需求

### 2.1 核心功能

#### 2.1.1 PPT模板处理
- **功能描述**: 支持保持复杂的PPT特效（3D、辉光、映像、阴影等）
- **输入**: PPT模板文件(.pptx)
- **处理**: 
  - 通过PowerPoint COM集成，从选择窗格命名中获取占位符（命名约定：ph_字段名）
  - 实现占位符与Excel表格数据字段的精确匹配
  - 保持PPT中的所有特效和格式
- **输出**: 处理后的PPT模板对象

#### 2.1.2 精确匹配算法
- **功能描述**: 将PPT占位符与Excel表头进行精确匹配
- **匹配策略**:
  - 精确匹配：占位符后缀与表头完全一致（如"ph_123"匹配"123"）
- **匹配流程**:
  1. 提取PPT中所有占位符（从选择窗格命名）
  2. 读取Excel表头字段
  3. 执行精确匹配算法
  4. 应用匹配结果

#### 2.1.2 多数据源支持
- **功能描述**: 支持多种数据源格式
- **支持的格式**:
  - Excel (.xlsx)
  - CSV (.csv)
  - JSON (.json)


#### 2.1.3 批量生成
- **功能描述**: 根据数据源和模板批量生成PPT
- **处理流程**:
  1. 读取数据源
  2. 对于每一条数据记录:
     - 加载PPT模板
     - 替换占位符
     - 保存生成的PPT
- **输出**: 批量生成的PPT文件

### 2.2 用户界面

#### 2.2.1 GUI模式
- **功能描述**: 提供简洁美观的图形界面，有文件选择器、占位符匹配界面、进度监控和错误日志显示,有自动匹配按钮、批量生成按钮
- **主要组件**:
  - 文件选择器（模板文件、数据源文件）
  - 占位符匹配界面（显示ppt模板和excel表头自动匹配结果）
  - 进度监控
  - 错误日志显示

#### 2.2.2 CLI模式
- **功能描述**: 命令行模式，支持自动化脚本和批处理
- **主要命令**:
  - 模板验证
  - 数据源检查
  - 批量生成
  - 配置管理

## 3. 技术规范

### 3.1 开发环境
- **编程语言**: Python 3.9+
- **GUI框架**: PySide6
- **数据处理**: pandas, openpyxl
- **PPT处理**: python-pptx, pywin32 (PowerPoint COM集成)
- **配置管理**: PyYAML

### 3.2 项目结构
```
poster-pyside6/
├── src/
│   ├── core/
│   │   ├── template_processor.py    # PPT模板处理
│   │   ├── data_reader.py           # 数据源读取
│   │   ├── ppt_generator.py         # PPT生成器
│   │   └── config_manager.py        # 配置管理
│   ├── gui/
│   │   ├── main_window.py           # 主窗口
│   │   ├── widgets/                 # 自定义控件
│   │   └── resources/               # 资源文件
│   ├── cli/
│   │   ├── commands.py              # CLI命令
│   │   └── main.py                  # CLI入口
│   └── utils/
│       ├── logger.py                # 日志工具
│       └── validators.py            # 验证工具
├── tests/                           # 测试文件
├── docs/                            # 文档
├── examples/                        # 示例文件
├── requirements.txt                 # 依赖列表
└── README.md                        # 项目说明
```

### 3.3 核心模块设计

#### 3.3.1 模板处理模块 (template_processor.py)
```python
class TemplateProcessor:
    def __init__(self):
        self.powerpoint_app = None  # PowerPoint COM应用程序实例
    
    def connect_to_powerpoint(self) -> bool:
        """连接到PowerPoint COM应用程序"""
        
    def load_template(self, template_path: str) -> Presentation:
        """通过COM接口加载PPT模板"""
        
    def find_placeholders(self, template: Presentation) -> List[str]:
        """从选择窗格命名中获取所有占位符"""
        
    def extract_shape_names(self, slide) -> Dict[str, Shape]:
        """提取幻灯片中所有形状的名称（从选择窗格）"""
        
    def replace_placeholders(self, template: Presentation, data: Dict[str, str]) -> Presentation:
        """根据数据替换占位符内容"""
        
    def save_presentation(self, presentation: Presentation, output_path: str) -> bool:
        """保存演示文稿"""
        
    def close_powerpoint(self):
        """关闭PowerPoint COM连接"""
```

#### 3.3.2 精确匹配模块 (exact_matcher.py)
```python
class ExactMatcher:
    def extract_placeholders(self, placeholders: List[str]) -> List[str]:
        """从占位符列表中提取字段名（去除ph_前缀）"""
        
    def find_exact_match(self, placeholder: str, headers: List[str]) -> str:
        """为占位符找到精确匹配的表头"""
        
    def match_placeholders_to_headers(self, placeholders: List[str], headers: List[str]) -> Dict[str, str]:
        """执行占位符与表头的精确匹配"""
```

#### 3.3.2 数据读取模块 (data_reader.py)
```python
class DataReader:
    def read_excel(self, file_path: str) -> pd.DataFrame
    def read_csv(self, file_path: str) -> pd.DataFrame
    def read_json(self, file_path: str) -> pd.DataFrame
    def validate_data(self, data: pd.DataFrame, required_fields: List[str]) -> bool
```

#### 3.3.3 PPT生成器模块 (ppt_generator.py)
```python
class PPTGenerator:
    def __init__(self, template_processor: TemplateProcessor, data_reader: DataReader)
    def generate_batch(self, template_path: str, data_source: str, output_dir: str) -> List[str]
    def generate_single(self, template_path: str, data: Dict[str, str], output_path: str) -> bool
```

### 3.4 数据规范

#### 3.4.1 模板占位符规范
- **命名约定**: `ph_字段名`
- **示例**: `ph_姓名`, `ph_部门`, `ph_工号`
- **支持类型**: 文本、数字、日期

#### 3.4.2 数据源字段规范
- **字段命名**: 与模板占位符去掉"ph_"前缀后的字段名对应
- **数据类型**: 
  - 文本: 字符串
  - 数字: 整数或浮点数
  - 日期: 日期格式或字符串

### 3.5 错误处理
- **模板错误**: 无效的PPT文件、缺少占位符
- **数据错误**: 缺少必要字段、数据格式不正确
- **生成错误**: 输出路径无效、权限不足
- **处理方式**: 
  - GUI模式: 显示错误对话框
  - CLI模式: 输出错误信息并返回非零退出码

### 3.6 性能考虑
- **批量处理**: 支持大数据集的分批处理
- **内存管理**: 及时释放PPT对象，避免内存泄漏
- **进度反馈**: 提供处理进度信息

## 4. 开发计划

### 4.1 第一阶段 (核心功能)
1. 实现模板处理模块
2. 实现数据读取模块
3. 实现PPT生成器模块
4. 基本的CLI接口

### 4.2 第二阶段 (用户界面)
1. 设计并实现GUI界面
2. 完善CLI功能
3. 添加配置管理功能

### 4.3 第三阶段 (优化和测试)
1. 性能优化
2. 单元测试和集成测试
3. 文档编写
4. 示例和教程

## 5. 质量保证

### 5.1 测试策略
- **单元测试**: 覆盖所有核心模块
- **集成测试**: 测试模块间的交互
- **端到端测试**: 测试完整的工作流程

### 5.2 代码规范
- **PEP 8**: 遵循Python编码规范
- **类型注解**: 使用类型提示提高代码可读性
- **文档字符串**: 为所有公共函数和类提供文档

### 5.3 版本控制
- **Git**: 使用Git进行版本控制
- **分支策略**: 使用功能分支进行开发
- **提交规范**: 遵循清晰的提交信息规范