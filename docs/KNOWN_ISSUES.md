# 已知问题与待办事项

本文档记录需要在未来处理但暂不紧急的事项。

## 架构层面

### God Class: main_window.py (1500+ 行)

`src/gui/main_window.py` 承担过多职责：UI 构建、文件选择、匹配、生成编排、Office 检测、设置管理、期趸数据处理等。

**建议拆分**:
- `FileSelectionPanel` — 模板/数据/输出路径选择
- `MatchResultsPanel` — 匹配表格展示与编辑
- `SettingsPanel` — 图片生成设置、消息模板
- `OfficeStatusPanel` — Office 可用性检测与状态显示
- `ChengbaoDataHandler` — 承保趸期数据输入与计算

### 测试框架迁移

`tests/unit/` 下约 49 个测试文件均为独立脚本，无 pytest 配置、无 conftest.py、无 fixtures、无覆盖率统计。

**建议**:
1. 创建 `conftest.py` 提供通用 fixtures
2. 逐步迁移测试到 pytest 风格
3. 添加 `pytest-cov` 实现 80% 覆盖率目标

### MD5 用于缓存键

`src/memory_management/memory_data_manager.py` 使用 MD5 计算文件哈希作为缓存键。非安全场景可接受，但理论上存在碰撞风险。

**建议**: 在后续优化中替换为 SHA-256。

## 技术债务

### 双目录结构 (`src/` vs 根目录)

根目录存在 `config_manager.py`、`gui/`、`core/` 等旧版文件，通过 `try/except ImportError` 回退兼容。

**建议**: 迁移到 `pyproject.toml` + `pip install -e .` 的标准包结构，彻底移除根目录副本。

### Worker 线程与 MainWindow 紧耦合

`PPTWorkerThread` 直接访问 `main_window.chengbao_user_inputs`、`main_window.data_path_edit` 等内部属性，无法独立测试。

**建议**: 通过依赖注入传递所需数据，而非持有 MainWindow 引用。
