# 承保趸期数据功能设计文档

## 架构概述

承保趸期数据功能基于现有的数据处理管道设计，延续"递交趸期数据"功能的设计模式，通过扩展现有的数据格式化模块实现。

## 核心组件

### 1. 数据处理层

#### 1.1 DataReader扩展（src/data_reader.py）
- **位置**：在现有加载逻辑中添加承保趸期数据检查
- **职责**：
  - 检测"SFYP2(不含短险续保)"和"首年保费"列是否存在
  - 调用数据格式化器计算承保趸期数据
  - 管理计算结果的存储

#### 1.2 DataFormatter扩展（src/memory_management/data_formatter.py）
- **新增方法**：`calculate_chengbao_term_data()`
- **职责**：
  - 实现比例计算逻辑
  - 处理三种计算场景
  - 管理用户输入收集

### 2. 用户界面层

#### 2.1 批量输入对话框（src/gui/chengbao_term_input_dialog.py）
- **职责**：提供批量输入界面
- **组件**：
  - 行号和保单号显示标签
  - 数值输入文本框（带验证）
  - 确定/取消按钮
- **界面布局**：
  ```
  第1行 保单号: ABC123 缴费年期: [ 20 ] 年交SFYP
  第3行 保单号: DEF456 缴费年期: [ 30 ] 年交SFYP
  第7行 保单号: GHI789 缴费年期: [ 10 ] 年交SFYP
  ```
- **验证规则**：只接受≥2的正整数

## 数据流

```
加载Excel文件
    ↓
读取DataFrame
    ↓
检查列存在性
    ├─ SFYP2列存在？
    └─ 首年保费列存在？
    ↓ YES
调用承保趸期数据计算
    ↓
循环处理每一行
    ├─ R = SFYP2 / 首年保费
    ├─ 分类处理（三种情况）
    └─ 记录需要用户输入的行
    ↓
如果有R=1的行
    ├─ 弹出批量输入对话框
    ├─ 用户填写数值
    └─ 验证输入
    ↓
生成承保趸期数据列
    ↓
存储结果并返回
```

## 算法实现

### 比例计算算法

```python
def calculate_chengbao_term_data(data):
    sfyp2_col = "SFYP2(不含短险续保)"
    premium_col = "首年保费"
    result_col = "承保趸期数据"

    for index, row in data.iterrows():
        sfyp2 = row[sfyp2_col]
        premium = row[premium_col]

        if pd.isna(sfyp2) or pd.isna(premium) or premium == 0:
            result = ""
        else:
            ratio = sfyp2 / premium

            if ratio == 0.1:
                result = "趸交FYP"
            elif ratio == 1:
                # 收集行号，等待用户输入
                pending_rows.append(index)
            else:
                # 0.2 <= ratio <= 0.9
                years = int(round(ratio * 10))
                result = f"{years}年交SFYP"

    return result_data, pending_rows
```

### 批量输入对话框设计

```python
class ChengbaoTermInputDialog(QDialog):
    def __init__(self, pending_rows):
        # 创建动态表单布局
        for row_index in pending_rows:
            # 创建标签：第X行 缴费年期:
            label = QLabel(f"第{row_index}行 缴费年期:")
            # 创建输入框（带验证）
            input_box = QLineEdit()
            input_box.setValidator(QIntValidator(2, 999))
            # 添加到布局

    def validate_inputs(self):
        # 检查所有输入是否有效（≥2的正整数）
        # 返回验证结果和输入值字典
```

## 模块依赖

```
data_reader.py
    ↓
data_formatter.py
    ↓
chengbao_term_input_dialog.py
    ↓
GUI显示
```

## 错误处理

1. **列不存在**：
   - 记录警告日志
   - 跳过承保趸期数据计算
   - 继续加载其他数据

2. **除零错误**：
   - 检查分母为0的情况
   - 跳过该行计算
   - 记录错误信息

3. **用户输入无效**：
   - 输入框直接拒绝无效字符
   - 用户必须输入有效值才能提交
   - 提供清除按钮

4. **数据格式错误**：
   - 转换前验证数据类型
   - 转换失败时使用空值
   - 记录详细错误日志

## 性能考虑

1. **计算优化**：
   - 只在必要时触发计算
   - 使用向量化操作处理DataFrame
   - 避免循环中的昂贵操作

2. **内存管理**：
   - 批量处理减少中间对象
   - 及时释放临时数据
   - 使用copy()避免修改原始数据

3. **UI响应性**：
   - 对话框显示前预计算所有数据
   - 使用异步方式处理大数据集
   - 提供进度指示器（如果需要）

## 测试策略

1. **单元测试**：
   - 测试比例计算逻辑
   - 测试三种场景的输出
   - 测试边界条件

2. **集成测试**：
   - 测试完整的数据加载流程
   - 测试与现有功能的兼容性
   - 测试数据存储结果

3. **UI测试**：
   - 测试对话框显示
   - 测试输入验证
   - 测试用户交互流程

## 扩展性

- 未来可以轻松添加其他趸期计算模式
- 可以扩展输入验证规则
- 可以添加更多用户自定义选项
- 支持不同的输出格式