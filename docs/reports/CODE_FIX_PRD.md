# 代码质量修复 PRD

| 文档属性 | 说明 |
|----------|------|
| 项目名称 | PPT 批量生成工具 - 代码质量修复 |
| 文档版本 | v1.0 |
| 创建日期 | 2026-04-08 |
| 优先级 | P1/P2 |
| 预计工作量 | 4-6 小时 |

---

## 1. 背景与目标

### 1.1 背景
在代码质量检查中发现项目存在以下问题：
- 30+ 处调试 print 语句混入生产代码
- 部分函数缺少类型注解
- 存在宽泛的异常捕获
- 代码复用不足

### 1.2 目标
- 移除所有调试 print 语句，统一使用 logger
- 补全核心函数的类型注解
- 修复宽泛的 except 捕获
- 提取公共清理方法，提升代码复用

### 1.3 范围
| 模块 | 修复内容 |
|------|----------|
| `src/gui/main_window.py` | 移除 print+logger 重复输出 |
| `src/gui/config_dialog.py` | 移除调试 print |
| `src/gui/add_text_dialog.py` | 移除调试 print |
| `src/data_reader.py` | 补全类型注解 |
| `src/data_formatter.py` | 移除重复 logger 导入、修复返回类型 |
| `src/ppt_generator.py` | 修复裸 except |
| `src/excel_detector.py` | 修复裸 except |

---

## 2. 问题清单

### 2.1 P0 级别 - 调试输出清理

| ID | 文件 | 行号 | 问题描述 | 修复方案 |
|----|------|------|----------|----------|
| DBG-001 | main_window.py | 312,318,327,335,340,348,473,481,496,504,516,524,567,591,594,648,669,745,756,829,832,875,877,884,887,927,987,992,995,1042,1054,1057,1060,1241 | print+logger 重复输出 | 移除 print()，保留 logger |
| DBG-002 | config_dialog.py | 187,349,363,368,371 | 调试 print 输出 | 移除或改为 logger.debug() |
| DBG-003 | add_text_dialog.py | 121-122 | 调试 print 输出 | 移除 |

### 2.2 P1 级别 - 类型注解补全

| ID | 文件 | 行号 | 问题描述 | 修复方案 |
|----|------|------|----------|----------|
| TYPE-001 | data_reader.py | 378 | `parent_widget` 参数无类型注解 | 添加 `parent_widget: Optional[QWidget] = None` |
| TYPE-002 | data_formatter.py | 295-299 | 返回类型为 `tuple[...]` | 改为 `Tuple[pd.DataFrame, List[Dict[str, Any]]]` |
| TYPE-003 | data_formatter.py | 318 | 缺少返回值类型注解 | 添加 `-> Tuple[pd.DataFrame, List[Dict[str, Any]]]` |

### 2.3 P1 级别 - 异常处理修复

| ID | 文件 | 行号 | 问题描述 | 修复方案 |
|----|------|------|----------|----------|
| EXC-001 | ppt_generator.py | 382 | 裸 `except:` 捕获 | 改为 `except (KeyError, AttributeError, TypeError) as e:` 并记录日志 |
| EXC-002 | excel_detector.py | 188 | 裸 `except Exception:` 未记录日志 | 添加 `logger.warning()` 或改为具体异常 |

### 2.4 P2 级别 - 代码复用

| ID | 文件 | 行号 | 问题描述 | 修复方案 |
|----|------|------|----------|----------|
| REFAC-001 | main_window.py | 1290-1318 | closeEvent 和 __del__ 重复逻辑 | 提取 `_cleanup_resources()` 私有方法 |
| REFAC-002 | data_formatter.py | 326-407 | 方法内重复导入 logger | 使用模块级 logger，移除方法内导入 |

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

### 3.2 类型注解规范

```python
# ❌ 修复前
def _process_chengbao_term_data(self, parent_widget=None):

# ✅ 修复后
from PySide6.QtWidgets import QWidget
# ...
def _process_chengbao_term_data(self, parent_widget: Optional[QWidget] = None):
```

```python
# ❌ 修复前
def calculate_chengbao_term_data(...) -> tuple[pd.DataFrame, List[Dict]]:

# ✅ 修复后
from typing import Tuple
# ...
def calculate_chengbao_term_data(...) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
```

### 3.3 异常处理规范

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
    logger.warning(f"获取行数据失败：{str(e)}，使用默认文件名")
    filename = f"{index+1}_output"
```

### 3.4 代码复用设计

```python
# ✅ 新增方法
def _cleanup_resources(self):
    """清理 PPT 生成器和工作线程资源"""
    try:
        if hasattr(self.worker_thread, 'get_ppt_generator'):
            ppt_generator = self.worker_thread.get_ppt_generator()
            if ppt_generator:
                ppt_generator.close()
                logger.info("已释放 PPTX 处理器资源")
    except Exception as e:
        logger.debug(f"清理 PPT 生成器资源时出错：{e}")
    
    try:
        if hasattr(self.worker_thread) and self.worker_thread.isRunning():
            self.worker_thread.quit()
            self.worker_thread.wait(3000)
    except Exception as e:
        logger.debug(f"等待工作线程结束时出错：{e}")

# 修改 closeEvent
def closeEvent(self, event):
    self._cleanup_resources()
    event.accept()

# 修改 __del__
def __del__(self):
    self._cleanup_resources()
```

---

## 4. 验收标准

### 4.1 功能验收
- [ ] 应用程序正常启动 (`python main.py`)
- [ ] 模板加载、数据加载功能正常
- [ ] 自动匹配功能正常
- [ ] 批量生成功能正常
- [ ] 关闭应用时无资源残留

### 4.2 代码质量验收
- [ ] 无 `print(f"[DEBUG]` 或 `print(f"[LOG]` 语句
- [ ] 核心函数类型注解完整
- [ ] 无裸 `except:` 语句
- [ ] 运行 `python -m py_compile` 无语法错误

### 4.3 测试验收
- [ ] 运行 `tests/unit/test_pptx_processor.py` 通过
- [ ] 运行 `tests/unit/test_gui_workflow.py` 通过

---

## 5. 实施计划

### 5.1 阶段划分

| 阶段 | 任务 | 预计时间 |
|------|------|----------|
| 阶段 1 | 移除调试 print 输出 (DBG-001~003) | 1 小时 |
| 阶段 2 | 补全类型注解 (TYPE-001~003) | 1 小时 |
| 阶段 3 | 修复异常处理 (EXC-001~002) | 1 小时 |
| 阶段 4 | 代码重构 (REFAC-001~002) | 1-2 小时 |
| 阶段 5 | 测试验证 | 1 小时 |

### 5.2 变更文件列表

| 文件 | 变更类型 | 优先级 |
|------|----------|--------|
| src/gui/main_window.py | 修改 | P0 |
| src/gui/config_dialog.py | 修改 | P0 |
| src/gui/widgets/add_text_dialog.py | 修改 | P0 |
| src/data_reader.py | 修改 | P1 |
| src/memory_management/data_formatter.py | 修改 | P1 |
| src/ppt_generator.py | 修改 | P1 |
| src/excel_detector.py | 修改 | P1 |

---

## 6. 风险评估

| 风险 | 可能性 | 影响 | 缓解措施 |
|------|--------|------|----------|
| 移除 print 可能影响调试 | 低 | 低 | 保留 logger，必要时可提高日志级别 |
| 类型注解可能引入新 bug | 低 | 中 | 充分测试，确保运行时兼容 |
| 异常处理修改可能掩盖问题 | 中 | 中 | 添加详细日志记录 |

---

## 7. 附录

### 7.1 相关文件
- [代码质量检查报告](./docs/reports/CODE_QUALITY_REPORT.md)
- [CLAUDE.md](./CLAUDE.md)

### 7.2 参考资料
- [PEP 484 - Type Hints](https://peps.python.org/pep-0484/)
- [Python Logging HOWTO](https://docs.python.org/3/howto/logging.html)
- [EAFP vs LBYL](https://docs.python.org/3/glossary.html#term-eafp)

---

**审批记录**

| 角色 | 姓名 | 日期 | 状态 |
|------|------|------|------|
| 产品负责人 | - | - | 待审批 |
| 技术负责人 | - | - | 待审批 |
