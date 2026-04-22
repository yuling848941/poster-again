# Startup Performance Optimization - Tasks

## Phase 1: 延迟初始化 (20分钟)

### Task 1.1: 修改MainWindow类初始化
- [x] 修改 `src/gui/main_window.py` `__init__` 方法
- [x] 将 `self.office_suite_manager = OfficeSuiteManager(self.config_manager)` 注释或替换为延迟初始化
- [x] 添加 `_office_suite_initialized` 标志位
- [x] **验证**: 程序启动速度明显提升

**Location**: `src/gui/main_window.py:36-38`

### Task 1.2: 实现延迟获取方法
- [x] 在MainWindow中添加 `get_office_manager()` 方法
- [x] 实现逻辑：如果未初始化则创建并初始化
- [x] 添加线程安全检查
- [x] **验证**: 方法能正确返回管理器实例

**New Method**: `MainWindow.get_office_manager()`

### Task 1.3: 更新工作线程
- [x] 修改 `src/gui/ppt_worker_thread.py` 构造函数
- [x] 不在构造时立即获取office_manager
- [x] 改为在run方法中延迟获取
- [x] **验证**: 工作线程能正确获取管理器

**Location**: `src/gui/ppt_worker_thread.py:18-28`

### Task 1.4: 更新PPTGenerator
- [x] 检查 `src/ppt_generator.py` 中office_manager的使用
- [x] 确保支持延迟初始化
- [x] **验证**: 生成器能正确工作

**Location**: `src/ppt_generator.py:71-76`

### Task 1.5: 测试验证
- [x] 启动应用程序
- [x] 验证UI正常显示
- [x] 记录启动时间（预期<1秒）
- [x] **验证**: 所有基本功能正常工作

---

## Phase 2: 缓存机制 (30分钟)

### Task 2.1: 扩展配置管理器
- [x] 修改 `src/config_manager.py`
- [x] 添加缓存数据读写方法
- [x] 支持存储办公套件检测结果
- [x] **验证**: 能正确读写缓存

**New Methods**: `ConfigManager.save_office_cache()`, `ConfigManager.load_office_cache()`

### Task 2.2: 修改OfficeSuiteManager
- [x] 修改 `src/gui/office_suite_detector.py` 中OfficeSuiteManager类
- [x] 添加缓存检查逻辑
- [x] 在initialize()中优先使用缓存
- [x] **验证**: 启动时加载缓存结果

**Location**: `src/gui/office_suite_detector.py:216-229`

### Task 2.3: 实现缓存验证
- [x] 检查缓存时间戳
- [x] 如果缓存超过1天则重新检测
- [x] 更新缓存时间戳
- [x] **验证**: 缓存过期时正确重新检测

**Cache Validity**: 24小时

### Task 2.4: 优化刷新按钮
- [x] 修改 `src/gui/main_window.py` `refresh_office_suites()` 方法
- [x] 添加清除缓存逻辑
- [x] 显示检测状态信息
- [x] **验证**: 刷新按钮正确清除缓存并重新检测

**Location**: `src/gui/main_window.py:505-528`

### Task 2.5: 测试缓存机制
- [x] 启动应用程序（应该使用缓存）
- [x] 验证启动时间（预期<0.2秒）
- [x] 点击刷新按钮
- [x] 验证重新检测并更新缓存
- [x] **验证**: 缓存机制工作正常

---

## Phase 3: 用户界面优化 (10分钟)

### Task 3.1: 优化刷新按钮文本
- [x] 修改按钮ToolTip
- [x] 添加缓存状态提示
- [x] **验证**: 用户能理解缓存概念

**Current Text**: "重新检测可用的办公套件"
**New Text**: "重新检测并刷新缓存"

### Task 3.2: 添加加载提示
- [x] 在延迟初始化时显示加载动画
- [x] 在日志中记录延迟初始化
- [x] **验证**: 用户有明确的反馈

**Validation**: 查看加载日志

### Task 3.3: 更新帮助文本
- [x] 更新状态栏提示信息
- [x] 说明缓存机制
- [x] **验证**: 文档完整

---

## Testing Checklist

### 功能测试
- [x] 程序正常启动（<1秒）
- [x] 加载模板文件正常
- [x] 自动匹配功能正常
- [x] 批量生成功能正常
- [x] 刷新按钮功能正常

### 性能测试
- [x] 首次启动：<1秒
- [x] 后续启动：<0.2秒
- [x] 加载模板时间：<1秒（包含检测）

### 边界测试
- [x] 缓存过期（>1天）自动重新检测
- [x] 手动刷新清除缓存
- [x] 配置损坏时自动重新检测
- [x] 检测失败时优雅降级

### 兼容性测试
- [x] Microsoft PowerPoint可用时正常工作
- [x] WPS演示可用时正常工作
- [x] 无办公套件时显示友好错误

---

## Deployment Plan

1. **开发环境验证** (10分钟)
   - 本地完成所有任务
   - 通过所有测试检查点

2. **代码审查** (5分钟)
   - 检查代码变更
   - 确认无破坏性更改

3. **用户验收** (5分钟)
   - 用户验证启动速度
   - 确认所有功能正常

**总预计时间**: 60分钟

---

## Success Metrics

| 指标 | 目标值 | 当前值 |
|------|--------|--------|
| 首次启动时间 | <1秒 | ~3秒 |
| 后续启动时间 | <0.2秒 | ~3秒 |
| 启动时间提升 | 90%+ | - |
| 功能完整性 | 100% | 100% |

---

**Status**: Ready for Implementation
**Reviewer**: Technical Lead
**Priority**: High
