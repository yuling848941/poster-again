# 代码质量修复 PRD (v2.0 - 全面修订版)

| 文档属性 | 说明 |
|----------|------|
| 项目名称 | PPT 批量生成工具 - 代码质量修复 |
| 文档版本 | v2.1 (修订版) |
| 创建日期 | 2026-04-08 |
| 修订日期 | 2026-04-08 |
| 优先级 | P0/P1/P2 |
| 预计工作量 | 10-15 小时 (已根据审视结果调整) |

---

## 1. 背景与目标

### 1.1 背景

在全面的代码质量检查中发现项目存在以下问题：

| 问题类型 | v1.0 评估 | 实际数量 | 偏差 |
|---------|----------|---------|------|
| 调试 print 语句 | 30+ 处 | **53 处** | -43% |
| 裸 except 捕获 | 2 处 | **16 处** | -87.5% |
| 宽泛 Exception 捕获 | 未评估 | **100+ 处** | N/A |
| 类型注解缺失 | 3 处 | **1 处 (已确认)** | -67% |
| 重复 logger 导入 | 1 处 | **5 处** | -400% |

### 1.2 目标

**P0 级别（必须修复）**:
- 移除所有调试 print 语句，统一使用 logger
- 修复所有裸 `except:` 捕获

**P1 级别（重要修复）**:
- 审查并修复关键的宽泛 `except Exception:` 捕获
- 补全核心函数的类型注解
- 移除重复的 logger 导入

**P2 级别（优化改进）**:
- 提取公共清理方法，提升代码复用
- 统一日志级别使用规范

### 1.3 范围

| 模块 | 修复内容 | 优先级 |
|------|----------|--------|
| `src/gui/main_window.py` | 移除 print+logger 重复输出、修复裸 except | P0 |
| `src/gui/config_dialog.py` | 移除调试 print | P0 |
| `src/gui/widgets/add_text_dialog.py` | 移除调试 print | P0 |
| `src/gui/simple_config_dialog.py` | 修复裸 except | P1 |
| `src/data_reader.py` | 补全类型注解 | P1 |
| `src/memory_management/data_formatter.py` | 移除重复 logger 导入、修复类型、裸 except | P0/P1 |
| `src/ppt_generator.py` | 修复裸 except | P1 |
| `src/excel_detector.py` | 审查宽泛 Exception 捕获 | P2 |
| `src/core/processors/pptx_processor.py` | 修复裸 except (2 处) | P1 |
| `src/core/processors/com_image_exporter.py` | 修复裸 except (4 处) | P1 |
| `src/core/detectors/office_suite_detector.py` | 修复裸 except (3 处) | P1 |
| `src/memory_management/memory_data_manager.py` | 修复裸 except | P1 |
| `src/memory_management/memory_optimizer.py` | 修复裸 except | P1 |

---

## 2. 详细问题清单

### 2.1 P0 级别 - 调试输出清理 (53 处)

#### DBG-001: main_window.py (34 处)

| 行号 | 问题描述 | 修复方案 |
|------|----------|----------|
| 312, 318 | print("[LOG] 已启用/禁用直接生成图片选项") | 移除 print，保留 logger |
| 327, 335 | print 图片格式更改 | 移除 print，保留 logger |
| 340, 348 | print 图片质量更改 | 移除 print，保留 logger |
| 473, 481 | print 选择模板文件 | 移除 print，保留 logger |
| 496, 504 | print 选择数据文件 | 移除 print，保留 logger |
| 516, 524 | print 选择输出目录 | 移除 print，保留 logger |
| 567, 591, 594 | print 创建输出文件夹/生成格式 | 移除 print，保留 logger |
| 648, 669 | print 启用贺语/日志消息 | 移除 print，保留 logger |
| 745, 756 | print+logger 重复输出 | 移除 print，保留 logger |
| 829, 832 | print+logger 重复输出 | 移除 print，保留 logger |
| 875, 877, 884, 887 | print+logger 重复输出 | 移除 print，保留 logger |
| 927, 987, 992, 995 | print+logger 重复输出 | 移除 print，保留 logger |
| 1042, 1054, 1057, 1060 | print+logger 重复输出 | 移除 print，保留 logger |
| 1241 | print 从配置加载 | 移除 print，保留 logger |

#### DBG-002: config_dialog.py (5 处)

| 行号 | 问题描述 | 修复方案 |
|------|----------|----------|
| 187 | print(f"[DEBUG] 配置管理器创建...") | 移除或改为 logger.debug() |
| 349, 363, 368, 371 | print 调试输出 | 移除或改为 logger.debug() |

#### DBG-003: add_text_dialog.py (2 处)

| 行号 | 问题描述 | 修复方案 |
|------|----------|----------|
| 121, 122 | print 前后文本 | 移除或改为 logger.debug() |

#### DBG-004: 其他文件 (12+ 处)

| 文件 | 预估数量 | 修复方案 |
|------|---------|----------|
| tests/unit/*.py | 10+ | 保留（测试文件允许 print） |

### 2.2 P0 级别 - 裸 except 捕获 (16 处)

| ID | 文件 | 行号 | 问题描述 | 修复方案 | 风险等级 |
|----|------|------|----------|----------|----------|
| EXC-001 | ppt_generator.py | 382 | 裸 except 捕获文件名生成错误 | 改为 `except (KeyError, AttributeError, TypeError)` | 中 |
| EXC-002 | data_formatter.py | 76 | 裸 except 在列类型转换中 | 改为 `except (ValueError, TypeError)` | 中 |
| EXC-003 | memory_data_manager.py | 211 | 裸 except 在析构函数中 | 保留但添加注释说明 | 低 |
| EXC-004 | memory_optimizer.py | 221 | 裸 except 在类型转换中 | 改为 `except (ValueError, TypeError)` | 低 |
| EXC-005 | office_suite_detector.py | 148, 182 | 裸 except 在注册表检测中 | 改为 `except (OSError, PermissionError)` | 中 |
| EXC-006 | office_suite_detector.py | 200 | 裸 except 在检测方法中 | 改为具体异常或添加日志 | 中 |
| EXC-007 | main_window.py | 982 | 裸 except 在用户输入解析中 | 改为 `except (ValueError, IndexError)` | 中 |
| EXC-008 | main_window.py | 1317 | 裸 except 在析构函数中 | 保留但添加注释说明 | 低 |
| EXC-009 | simple_config_dialog.py | 160 | 裸 except 在配置获取中 | 改为 `except Exception as e:` 并记录日志 | 中 |
| EXC-010 | com_image_exporter.py | 98, 162, 165, 168 | 裸 except 在 COM 对象获取中 | 改为 `except pywintypes.com_error` | 高 |
| EXC-011 | pptx_processor.py | 127, 544 | 裸 except 在形状处理和文件清理中 | 改为 `except Exception as e:` 并记录日志 | 中 |

### 2.3 P1 级别 - 宽泛 Exception 捕获 (关键 10 处)

**注意**: 100+ 处 `except Exception` 中大部分是合理的边界错误处理，以下列出需要特别审查的：

| ID | 文件 | 行号 | 问题描述 | 建议修复 |
|----|------|------|----------|----------|
| WEXC-001 | data_reader.py | 82, 116, 149, 162, 173, 182, 213, 298, 465 | 多处宽泛 Exception 捕获 | 审查每处是否应记录日志或重新抛出 |
| WEXC-002 | ppt_generator.py | 104, 133, 162, 200, 322, 398, 442, 459, 469, 483, 524, 541, 551 | 多处宽泛 Exception 捕获 | 审查关键路径上的异常处理 |
| WEXC-003 | excel_detector.py | 59, 63, 87, 114, 158, 226, 329, 338, 368, 422 | 多处宽泛 Exception 捕获 | 添加详细日志记录 |

### 2.4 P1 级别 - 类型注解补全

| ID | 文件 | 行号 | 问题描述 | 修复方案 |
|----|------|------|----------|----------|
| TYPE-001 | data_formatter.py | 299 | 返回类型使用 `tuple[...]` 而非 `Tuple[...]` | 改为 `Tuple[pd.DataFrame, List[Dict[str, Any]]]` |
| TYPE-002 | data_reader.py | 378 | `_process_chengbao_term_data` 缺少完整类型注解 | 添加 `parent_widget: Optional[QWidget] = None` 和 `-> None` |

### 2.5 P2 级别 - 代码复用和日志规范

| ID | 文件 | 行号 | 问题描述 | 修复方案 |
|----|------|------|----------|----------|
| REFAC-001 | main_window.py | 1290-1318 | closeEvent 和 `__del__` 重复逻辑 | 提取 `_cleanup_resources()` 私有方法 |
| REFAC-002 | data_formatter.py | 326, 332, 363, 372, 407 | 方法内重复导入 logger | 移除，使用模块级 logger |
| REFAC-003 | 全局 | 多处 | logger 级别使用不一致 | 制定日志级别使用规范 |

---

## 3. 技术设计

### 3.1 日志统一规范

```python
# ❌ 修复前
print(f"[LOG] 已启用直接生成图片选项")
logger.info("已启用直接生成图片选项")

# ✅ 修复后
logger.info("已启用直接生成图片选项")
```

### 3.2 日志级别使用规范

| 级别 | 使用场景 | 示例 |
|------|----------|------|
| DEBUG | 调试信息，默认不输出 | `logger.debug(f"处理第{i}行数据")` |
| INFO | 正常业务流程 | `logger.info("PPT 模板加载成功")` |
| WARNING | 非预期但不影响流程 | `logger.warning("未找到可选列，使用默认值")` |
| ERROR | 错误但已处理 | `logger.error(f"文件加载失败：{e}")` |
| CRITICAL | 严重错误，流程中断 | `logger.critical("关键组件初始化失败")` |

### 3.3 类型注解规范

```python
# ❌ 修复前
def _process_chengbao_term_data(self, parent_widget=None):

# ✅ 修复后
from PySide6.QtWidgets import QWidget
from typing import Optional

def _process_chengbao_term_data(self, parent_widget: Optional[QWidget] = None) -> None:
```

```python
# ❌ 修复前
def calculate_chengbao_term_data(...) -> tuple[pd.DataFrame, List[Dict]]:

# ✅ 修复后
from typing import Tuple, List, Dict, Any
import pandas as pd

def calculate_chengbao_term_data(...) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
```

### 3.4 异常处理规范

```python
# ❌ 修复前
try:
    filename = f"{index+1}_{row.get('姓名', 'output')}"
except:
    filename = f"{index+1}_output"

# ✅ 修复后
try:
    filename = f"{index+1}_{row.get('姓名', 'output')}"
except (KeyError, AttributeError, TypeError) as e:
    logger.debug(f"获取行数据失败：{str(e)}，使用默认文件名")
    filename = f"{index+1}_output"
```

```python
# ✅ 特殊场景：析构函数中可保留裸 except
def __del__(self):
    try:
        self._cleanup_resources()
    except:
        # 析构函数中忽略所有异常，避免干扰 Python 垃圾回收
        pass
```

### 3.5 COM 异常处理规范

```python
# ❌ 修复前
try:
    powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
except:
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")

# ✅ 修复后
import pywintypes

try:
    powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
except pywintypes.com_error:
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
```

---

## 4. 验收标准

### 4.1 功能验收

- [ ] 应用程序正常启动 (`python main.py`)
- [ ] 模板加载、数据加载功能正常
- [ ] 自动匹配功能正常
- [ ] 批量生成功能正常
- [ ] 关闭应用时无资源残留
- [ ] 图片导出功能正常（如有 Office/WPS）

### 4.2 代码质量验收

- [ ] 无 `print(f"[DEBUG]` 或 `print(f"[LOG]` 语句（测试文件除外）
- [ ] 核心函数类型注解完整
- [ ] 无裸 `except:` 语句（析构函数除外）
- [ ] 运行 `python -m py_compile src/*.py` 无语法错误
- [ ] 所有 logger 导入统一为模块级

### 4.3 测试验收

- [ ] 运行 `python tests/unit/test_pptx_processor.py` 通过
- [ ] 运行 `python tests/unit/test_chengbao_term_data.py` 通过

**注意**: `test_gui_workflow.py` 依赖外部 Excel 文件 `KA 承保快报.xlsx`，不纳入自动化测试范围。

---

## 5. 实施计划

### 5.1 阶段划分

| 阶段 | 任务 | 文件数 | 问题数 | 预计时间 |
|------|------|--------|--------|----------|
| 阶段 1 | 移除调试 print 输出 (DBG-001~004) | 3 | 53 处 | 3-4 小时 |
| 阶段 2 | 修复裸 except 捕获 (EXC-001~011) | 9 | 16 处 | 3-4 小时 |
| 阶段 3 | 审查宽泛 Exception 捕获 (WEXC) | 3 | 10+ 处 | 2-3 小时 |
| 阶段 4 | 补全类型注解 (TYPE-001~004) | 3 | 4 处 | 1 小时 |
| 阶段 5 | 代码重构 (REFAC-001~003) | 2 | 3 处 | 1-2 小时 |
| 阶段 6 | 测试验证 | - | - | 1.5 小时 |

**总计**: 10-15 小时

### 5.2 变更文件列表

| 文件 | 变更类型 | 优先级 | 问题数 |
|------|----------|--------|--------|
| src/gui/main_window.py | 修改 | P0 | 34+ print, 2 except |
| src/gui/config_dialog.py | 修改 | P0 | 5 print |
| src/gui/widgets/add_text_dialog.py | 修改 | P0 | 2 print |
| src/gui/simple_config_dialog.py | 修改 | P1 | 1 except |
| src/data_reader.py | 修改 | P1 | 类型注解，9 except |
| src/memory_management/data_formatter.py | 修改 | P0/P1 | 1 except, 5 logger 重复，类型注解 |
| src/ppt_generator.py | 修改 | P1 | 1 except, 13+ Exception |
| src/excel_detector.py | 修改 | P2 | 10+ Exception 审查 |
| src/core/processors/pptx_processor.py | 修改 | P1 | 2 except |
| src/core/processors/com_image_exporter.py | 修改 | P1 | 4 except |
| src/core/detectors/office_suite_detector.py | 修改 | P1 | 3 except |
| src/memory_management/memory_data_manager.py | 修改 | P1 | 1 except |
| src/memory_management/memory_optimizer.py | 修改 | P1 | 1 except |

---

## 6. 风险评估

| 风险 | 可能性 | 影响 | 缓解措施 |
|------|--------|------|----------|
| 移除 print 可能影响调试 | 低 | 低 | 保留 logger，必要时可提高日志级别 |
| 类型注解可能引入新 bug | 低 | 中 | 充分测试，确保运行时兼容性 |
| 异常处理修改可能掩盖问题 | 中 | 中 | 添加详细日志记录，保留关键宽泛捕获 |
| COM 异常处理修改可能导致导出失败 | 中 | 高 | **已修订**: 捕获 `(pywintypes.com_error, OSError)`，保留降级逻辑 |
| GUI 行为变更（关闭事件） | 低 | 中 | 充分测试窗口关闭流程 |

---

## 7. 修订记录 (v2.1)

### 7.1 批判性审视后发现的问题

| 问题 ID | 类型 | 描述 | 状态 |
|--------|------|------|------|
| REV-001 | 类型注解 | PRD 声称只有 2 处缺失，实际项目中混用 `Tuple[...]` 和 `tuple[...]` 语法 | 已修复 |
| REV-002 | 异常审查标准 | WEXC 审查标准模糊，部分已记录日志的 `except Exception as e:` 被列为问题 | 已明确规则 |
| REV-003 | COM 异常处理 | 原修复方案只捕获 `pywintypes.com_error`，可能遗漏 `OSError` 等其他异常 | 已修订方案 |
| REV-004 | 析构函数例外 | `closeEvent` 中的 `except:` 是否适用例外规则不明确 | 已明确边界 |
| REV-005 | 工作量估算 | 8-12 小时可能不足，未考虑 COM 环境测试成本 | 已调整为 10-15 小时 |
| REV-006 | 遗漏问题 | `ppt_generator.py` 等文件缺少类型注解未列入清单 | 已补充 |

### 7.2 修订内容详情

#### 7.2.1 COM 异常处理方案修订 (EXC-010)

**原方案**:
```python
except pywintypes.com_error:
```

**修订后**:
```python
except (pywintypes.com_error, OSError) as e:
    logger.debug(f"获取活动 PowerPoint 实例失败：{e}，尝试创建新实例")
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
```

**原因**: `GetActiveObject` 可能抛出 `OSError`（如权限不足）或其他 COM 相关异常，需保留降级逻辑。

#### 7.2.2 宽泛 Exception 捕获审查标准

**明确规则**:
- ✅ **可接受**: `except Exception as e:` 已使用 `logger.error()` 记录日志并包含详细错误信息
- ⚠️ **需审查**: 关键业务路径上的宽泛捕获，需添加注释说明为何使用宽泛捕获
- ❌ **需修复**: 静默吞掉异常且无日志记录的 `except Exception:` 或 `except:`

#### 7.2.3 析构函数例外规则边界

**明确规则**:
- ✅ `__del__()` 析构函数中可保留裸 `except:`，但必须添加注释说明原因
- ⚠️ `closeEvent()` 等用户交互回调中的异常处理应遵循正常规范（记录日志 + 用户提示）
- ✅ 资源清理方法（如 `_cleanup_resources()`）中的异常可忽略，但需记录调试日志

#### 7.2.4 类型注解补全清单扩展

| ID | 文件 | 行号 | 问题描述 | 修复方案 |
|----|------|------|----------|----------|
| TYPE-001 | data_formatter.py | 299 | 返回类型混用 `tuple[...]` 和 `Tuple[...]` | 统一使用 `Tuple[pd.DataFrame, List[Dict[str, Any]]]` |
| TYPE-002 | data_reader.py | 378 | `_process_chengbao_term_data` 缺少完整类型注解 | 添加 `parent_widget: Optional[QWidget] = None` 和 `-> None` |
| TYPE-003 | ppt_generator.py | 61 | `set_progress_callback` 的 callback 参数缺少类型注解 | 添加 `callback: Callable[[int, str], None]` |
| TYPE-004 | ppt_generator.py | 70 | `_update_progress` 参数缺少类型注解 | 添加 `progress: int, message: str` |

#### 7.2.5 工作量估算调整

| 阶段 | 原估算 | 修订后 | 原因 |
|------|--------|--------|------|
| 阶段 1 | 2-3 小时 | 3-4 小时 | 53 处 print 需逐处核实上下文 |
| 阶段 2 | 2-3 小时 | 3-4 小时 | COM 异常处理需谨慎测试 |
| 阶段 3 | 1-2 小时 | 2-3 小时 | 需明确每处 Exception 的审查标准 |
| 阶段 4 | 30 分钟 | 1 小时 | 扩展类型注解清单 |
| 阶段 5 | 1-2 小时 | 1-2 小时 | 不变 |
| 阶段 6 | 1 小时 | 1.5 小时 | 增加 COM 环境回归测试 |
| **总计** | **8-12 小时** | **10-15 小时** | |

---

## 8. 附录

### 8.1 相关文件

- [代码质量检查报告](./docs/reports/CODE_QUALITY_REPORT.md)
- [CLAUDE.md](./CLAUDE.md)

### 8.2 参考资料

- [PEP 484 - Type Hints](https://peps.python.org/pep-0484/)
- [PEP 8 - Programming Recommendations on Exception Handling](https://peps.python.org/pep-0008/#programming-recommendations)
- [Python Logging HOWTO](https://docs.python.org/3/howto/logging.html)
- [EAFP vs LBYL](https://docs.python.org/3/glossary.html#term-eafp)

### 8.3 GUI 功能模块完整清单

#### 8.3.1 核心窗口组件

| 文件 | 功能描述 | 主要组件 |
|------|----------|----------|
| `src/gui/main_window.py` | 主窗口，应用入口 | 文件选择、匹配表格、进度显示、日志输出、批量生成控制 |
| `src/gui/ppt_worker_thread.py` | 后台工作线程 | 自动匹配任务、批量生成任务、进度信号发射 |

#### 8.3.2 配置管理组件

| 文件 | 功能描述 | 主要组件 |
|------|----------|----------|
| `src/gui/config_dialog.py` | 配置保存/加载对话框 (完整版) | 配置导出导入、兼容性检查、最近配置列表 |
| `src/gui/simple_config_dialog.py` | 简化配置管理对话框 | 当前状态显示、配置文件加载 |
| `src/gui/simple_config_manager.py` | 配置管理器 (非 GUI) | GUI 状态收集、恢复、趸期映射管理 |

#### 8.3.3 功能组件

| 文件 | 功能描述 | 主要组件 |
|------|----------|----------|
| `src/gui/widgets/add_text_dialog.py` | 文本增加对话框 | 前缀/后缀设置、实时预览 |
| `src/gui/chengbao_term_input_dialog.py` | 承保趸期数据批量输入对话框 | 多行输入、保单号显示、数值验证 |
| `src/gui/match_table_manager.py` | 匹配表格管理器 | 表格填充、状态同步 |
| `src/gui/settings_manager.py` | 设置管理器 | 图片生成设置、贺语模板管理 |
| `src/gui/path_manager.py` | 路径管理器 | 路径记忆、恢复、验证 |

#### 8.3.4 UI 样式组件

| 文件 | 功能描述 |
|------|----------|
| `src/gui/ui_constants.py` | 统一 UI 样式常量 (颜色、字体、尺寸、按钮样式、表格样式等) |

#### 8.3.5 GUI 详细功能清单

**主窗口 (main_window.py) 功能模块:**

| 功能模块 | 描述 | 优先级 |
|----------|------|--------|
| 模板文件选择 | 浏览选择 PPTX 模板文件，支持路径记忆 | P0 |
| 数据源文件选择 | 浏览选择 Excel/CSV/JSON 数据文件，支持路径记忆 | P0 |
| 输出目录选择 | 选择输出目录，自动创建时间戳子文件夹 | P0 |
| 自动匹配 | 启动后台线程进行占位符自动匹配 | P0 |
| 批量生成 | 启动后台线程批量生成 PPT 或图片 | P0 |
| 进度显示 | 实时显示处理进度和状态消息 | P0 |
| 直接生成图片 | 复选框启用图片导出模式 | P1 |
| 图片格式选择 | PNG/JPG格式选择 (启用图片生成时显示) | P1 |
| 图片质量选择 | 原始大小/1.5 倍/2 倍/3 倍/4 倍 (启用图片生成时显示) | P1 |
| 贺语模板 | 支持 {{表头}} 占位符的贺语模板编辑 | P1 |
| 匹配结果表格 | 显示占位符与数据列的匹配结果 | P0 |
| 自定义右键菜单 | 增加文本、递交趸期数据、承保趸期数据 | P1 |
| 配置管理 | 打开配置保存/加载对话框 | P1 |
| 路径记忆恢复 | 启动时自动恢复上次使用的路径 | P2 |
| 设置防抖保存 | 图片设置和贺语模板变化后 500ms 自动保存 | P2 |
| 窗口关闭资源释放 | 释放 PPT 生成器资源，等待工作线程完成 | P0 |

**承保趸期数据输入对话框 (chengbao_term_input_dialog.py) 功能:**

| 功能 | 描述 |
|------|------|
| 批量输入界面 | 滚动区域支持多行输入 |
| 保单号显示 | 只读但可复制的保单号显示框 |
| 数值验证 | 只接受≥2 的正整数输入 |
| 一键确定/取消 | 验证通过后批量提交或取消操作 |
| 窗口自适应大小 | 根据输入行数调整窗口尺寸 |

**配置管理对话框 (config_dialog.py/simple_config_dialog.py) 功能:**

| 功能 | 描述 |
|------|------|
| 保存当前配置 | 导出 GUI 状态到.pptcfg 文件 |
| 加载配置文件 | 导入配置文件并恢复 GUI 状态 |
| 兼容性检查 | 检查配置与当前模板/数据的兼容性 |
| 最近配置列表 | 显示最近使用的配置文件 |
| 配置预览 | 显示配置中的规则数量和详情 |

### 8.4 GUI 异常处理场景补充

| 场景 | 当前处理方式 | 改进建议 |
|------|-------------|----------|
| 文件被占用 | 临时文件 + 复制模式 | 添加用户提示对话框 |
| 生成过程取消 | 未实现取消按钮 | 建议添加取消按钮和中断处理 |
| 无效数据输入 | 部分验证 (如承保趸期≥2) | 统一输入验证框架 |
| Office/WPS未安装 | 无需 Office/WPS(纯Python) | N/A |
| 承保趸期数据缺失 | 弹出批量输入对话框 | 已实现 |
| 配置文件格式错误 | QMessageBox 错误提示 | 已实现 |
| 占位符不匹配 | 跳过并记录日志 | 建议添加用户提示 |

### 8.5 GUI 性能验收标准补充

| 指标 | 验收标准 | 测试方法 |
|------|----------|----------|
| 大量数据加载 (100+ 行) | 界面响应时间 < 2 秒 | 使用大型 Excel 文件测试 |
| 后台任务 UI 冻结 | 进度条实时更新，无卡顿 | 观察进度条和按钮状态 |
| 内存使用上限 | 峰值内存 < 500MB | 使用任务管理器监控 |
| 配置文件加载 | < 1 秒 | 使用含 50+ 规则的配置文件 |
| 承保趸期对话框 | 10 行输入窗口高度自适应 | 手动测试不同行数场景 |

### 8.6 GUI 测试计划补充

#### 8.6.1 手动测试用例

| ID | 测试场景 | 预期结果 |
|----|----------|----------|
| GUI-001 | 启动应用，选择模板/数据/输出目录 | 路径正确显示，路径记忆生效 |
| GUI-002 | 点击自动匹配 | 匹配结果正确填充表格 |
| GUI-003 | 右键菜单 - 增加文本 | 对话框弹出，前后缀设置生效 |
| GUI-004 | 右键菜单 - 承保趸期数据 | 对话框弹出，输入验证生效 |
| GUI-005 | 启用直接生成图片 | 图片格式/质量下拉菜单显示 |
| GUI-006 | 贺语模板编辑 | {{表头}} 格式正确解析 |
| GUI-007 | 配置保存/加载 | 配置正确导出导入，兼容性检查生效 |
| GUI-008 | 窗口关闭 | 无资源残留，无异常抛出 |

#### 8.6.2 自动化测试范围

| 测试文件 | 覆盖范围 | 状态 |
|----------|----------|------|
| `tests/unit/test_pptx_processor.py` | PPTX 处理器功能 | 已实现 |
| `tests/unit/test_chengbao_term_data.py` | 承保趸期数据计算 | 已实现 |
| `tests/unit/test_gui_workflow.py` | GUI 工作流 | 依赖外部文件，手动测试 |

### 8.7 v2.0 修订说明

相较于 v1.0，v2.0 主要修订：

1. **问题数量修正**:
   - print 语句：30+ → 53 处
   - 裸 except: 2 → 16 处
   - 新增宽泛 Exception 捕获审查：100+ 处

2. **新增文件**:
   - `src/core/processors/pptx_processor.py`
   - `src/core/processors/com_image_exporter.py`
   - `src/core/detectors/office_suite_detector.py`
   - `src/memory_management/memory_data_manager.py`
   - `src/memory_management/memory_optimizer.py`
   - `src/gui/simple_config_dialog.py`

3. **工作量调整**: 4-6 小时 → 8-12 小时

4. **新增规范**:
   - 日志级别使用规范
   - COM 异常处理规范
   - 析构函数例外说明

### 8.8 v2.1 修订说明 (批判性审视后)

相较于 v2.0，v2.1 主要修订：

1. **COM 异常处理方案修订**: 捕获 `(pywintypes.com_error, OSError)`，保留降级逻辑

2. **宽泛 Exception 捕获审查标准明确**:
   - ✅ 已记录日志的 `except Exception as e:` 可接受
   - ⚠️ 关键路径上的宽泛捕获需添加注释
   - ❌ 静默吞掉异常的需修复

3. **析构函数例外规则边界明确**:
   - ✅ `__del__()` 可保留裸 `except:`（需注释）
   - ⚠️ `closeEvent()` 等用户交互回调遵循正常规范

4. **类型注解补全清单扩展**: 新增 TYPE-003, TYPE-004

5. **工作量估算调整**: 8-12 小时 → 10-15 小时

---

**审批记录**

| 角色 | 姓名 | 日期 | 状态 |
|------|------|------|------|
| 产品负责人 | - | - | 待审批 |
| 技术负责人 | - | - | 待审批 |
