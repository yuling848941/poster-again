# Office Suite Detection Caching

## Purpose
通过缓存办公套件检测结果，将程序后续启动时间从0.5秒优化至0.1秒（再提升80%+），结合延迟初始化实现90%+总优化。

## ADDED Requirements

### Requirement: 配置文件缓存机制
OfficeSuiteManager SHALL在配置文件中缓存检测结果，避免重复检测。

#### Scenario: 缓存检测结果
- **GIVEN** 执行办公套件检测
- **WHEN** 检测完成并获得结果
- **THEN** 系统应将结果保存到配置文件
- **AND** 缓存数据应包含可用套件列表
- **AND** 缓存数据应包含当前选择套件
- **AND** 缓存数据应包含时间戳

#### Scenario: 缓存数据格式
- **GIVEN** 需要保存缓存
- **WHEN** 保存操作执行
- **THEN** 应使用以下JSON结构:
  ```json
  {
    "available_suites": ["MICROSOFT", "WPS"],
    "effective_suite": "MICROSOFT",
    "current_suite": "AUTO",
    "timestamp": 1731398400,
    "version": "1.0"
  }
  ```

#### Scenario: 缓存存储位置
- **GIVEN** 保存缓存数据
- **WHEN** 系统执行保存
- **THEN** 应存储在config.yaml文件中
- **AND** 使用键名"office_suite_cache"

### Requirement: 启动时加载缓存
OfficeSuiteManager.initialize() SHALL优先从缓存加载配置。

#### Scenario: 缓存命中
- **GIVEN** 缓存数据存在且未过期
- **WHEN** 调用initialize()方法
- **THEN** 系统应直接加载缓存结果
- **AND** 不应执行COM检测
- **AND** 应记录"使用缓存的办公套件配置"日志
- **AND** 初始化应在100ms内完成

#### Scenario: 缓存未命中
- **GIVEN** 缓存数据不存在
- **WHEN** 调用initialize()方法
- **THEN** 系统应执行完整的检测流程
- **AND** 检测完成后应保存新缓存

#### Scenario: 缓存过期
- **GIVEN** 缓存时间戳超过24小时
- **WHEN** 调用initialize()方法
- **THEN** 系统应视为缓存无效
- **AND** 应执行重新检测
- **AND** 应更新缓存时间戳
- **AND** 应记录"缓存已过期，重新检测"日志

### Requirement: 缓存有效性验证
系统 SHALL验证缓存数据完整性和时效性。

#### Scenario: 缓存完整性检查
- **GIVEN** 加载缓存数据
- **WHEN** 验证缓存
- **THEN** 必须包含所有必需字段:
  - available_suites
  - effective_suite或current_suite
  - timestamp
- **AND** 如果字段缺失，应视为无效缓存

#### Scenario: 缓存时效性检查
- **GIVEN** 缓存数据有timestamp
- **WHEN** 检查时效性
- **THEN** 应计算: 当前时间 - timestamp < 86400秒(24小时)
- **AND** 如果超时，应视为过期缓存

#### Scenario: 缓存版本检查
- **GIVEN** 缓存数据包含version字段
- **WHEN** 检查版本兼容性
- **THEN** 如果版本不匹配，应视为无效缓存
- **AND** 应执行重新检测

### Requirement: 缓存更新机制
系统 SHALL在适当时候更新缓存数据。

#### Scenario: 手动刷新
- **GIVEN** 用户点击"刷新"按钮
- **WHEN** 执行refresh_office_suites()
- **THEN** 系统应清除现有缓存
- **AND** 应执行重新检测
- **AND** 应保存新缓存结果

#### Scenario: 自动更新
- **GIVEN** 检测到系统变化(如安装了新办公套件)
- **WHEN** 执行检测
- **THEN** 应更新缓存
- **AND** 时间戳应为当前时间

### Requirement: 缓存清除功能
系统 SHALL提供清除缓存的机制。

#### Scenario: 清除指定缓存
- **GIVEN** 缓存数据存在
- **WHEN** 调用clear_office_cache()
- **THEN** 应从配置文件中移除office_suite_cache项
- **AND** 后续初始化应执行完整检测

#### Scenario: 清除所有缓存
- **GIVEN** 配置文件中存在多个缓存
- **WHEN** 调用clear_all_cache()
- **THEN** 应清除所有相关缓存数据
- **AND** 配置文件应保持其他设置不变

### Requirement: OfficeSuiteManager.initialize()行为变更
initialize()方法 SHALL首先尝试加载缓存，只有在缓存无效时才执行检测。

#### Scenario: 新的初始化流程
- **GIVEN** 调用initialize()
- **WHEN** 方法执行
- **THEN** 应执行以下步骤:
  1. 尝试加载缓存
  2. 验证缓存有效性
  3. 如果缓存有效，加载并返回
  4. 如果缓存无效，执行检测
  5. 保存新缓存结果

#### Scenario: 初始化性能目标
- **GIVEN** 缓存有效
- **WHEN** 执行initialize()
- **THEN** 总耗时应 < 100ms

- **GIVEN** 缓存无效
- **WHEN** 执行initialize()
- **THEN** 总耗时可能较长(2-3秒)
- **AND** 但仅在必要时执行

### Requirement: 刷新按钮功能增强
refresh_office_suites()方法 SHALL清除缓存并重新检测。

#### Scenario: 增强刷新逻辑
- **GIVEN** 用户点击"刷新"按钮
- **WHEN** 执行refresh_office_suites()
- **THEN** 应执行以下操作:
  1. 清除缓存数据
  2. 重置初始化状态
  3. 重新执行initialize()
  4. 刷新UI显示
  5. 更新日志信息

#### Scenario: 刷新状态提示
- **GIVEN** 正在刷新
- **WHEN** 用户等待
- **THEN** 日志应显示:
  - "正在清除缓存并重新检测..."
  - "使用缓存的办公套件配置"(如果缓存有效)
  - 或检测结果(如果重新检测)

## Design Constraints

1. **配置兼容性**: 缓存机制必须向后兼容无缓存的配置文件
2. **原子性**: 缓存保存操作应具有原子性，避免文件损坏
3. **错误恢复**: 缓存损坏时应自动回退到检测模式
4. **磁盘空间**: 缓存数据应尽可能小(<1KB)

## Performance Targets

| 指标 | 目标值 | 延迟初始化后 |
|------|--------|-------------|
| 首次启动时间 | <0.5秒 | ~0.5秒 |
| 后续启动时间 | <0.1秒 | ~0.5秒 |
| 缓存加载耗时 | <50ms | N/A |
| 缓存保存耗时 | <100ms | N/A |
| 总优化提升 | 97%+ | 83%+ |

## Related Capabilities
- startup-performance-optimization.lazy-initialization (前置条件)
- startup-performance-optimization.interface-optimization (UI增强)

---

**Dependencies**
- ConfigManager支持缓存读写
- OfficeSuiteManager支持缓存注入
- MainWindow支持缓存清除操作
## Requirements
### Requirement: 配置文件缓存机制
OfficeSuiteManager SHALL在配置文件中缓存检测结果，避免重复检测。

#### Scenario: 缓存检测结果
- **GIVEN** 执行办公套件检测
- **WHEN** 检测完成并获得结果
- **THEN** 系统应将结果保存到配置文件
- **AND** 缓存数据应包含可用套件列表
- **AND** 缓存数据应包含当前选择套件
- **AND** 缓存数据应包含时间戳

#### Scenario: 缓存数据格式
- **GIVEN** 需要保存缓存
- **WHEN** 保存操作执行
- **THEN** 应使用以下JSON结构:
  ```json
  {
    "available_suites": ["MICROSOFT", "WPS"],
    "effective_suite": "MICROSOFT",
    "current_suite": "AUTO",
    "timestamp": 1731398400,
    "version": "1.0"
  }
  ```

#### Scenario: 缓存存储位置
- **GIVEN** 保存缓存数据
- **WHEN** 系统执行保存
- **THEN** 应存储在config.yaml文件中
- **AND** 使用键名"office_suite_cache"

### Requirement: 启动时加载缓存
OfficeSuiteManager.initialize() SHALL优先从缓存加载配置。

#### Scenario: 缓存命中
- **GIVEN** 缓存数据存在且未过期
- **WHEN** 调用initialize()方法
- **THEN** 系统应直接加载缓存结果
- **AND** 不应执行COM检测
- **AND** 应记录"使用缓存的办公套件配置"日志
- **AND** 初始化应在100ms内完成

#### Scenario: 缓存未命中
- **GIVEN** 缓存数据不存在
- **WHEN** 调用initialize()方法
- **THEN** 系统应执行完整的检测流程
- **AND** 检测完成后应保存新缓存

#### Scenario: 缓存过期
- **GIVEN** 缓存时间戳超过24小时
- **WHEN** 调用initialize()方法
- **THEN** 系统应视为缓存无效
- **AND** 应执行重新检测
- **AND** 应更新缓存时间戳
- **AND** 应记录"缓存已过期，重新检测"日志

### Requirement: 缓存有效性验证
系统 SHALL验证缓存数据完整性和时效性。

#### Scenario: 缓存完整性检查
- **GIVEN** 加载缓存数据
- **WHEN** 验证缓存
- **THEN** 必须包含所有必需字段:
  - available_suites
  - effective_suite或current_suite
  - timestamp
- **AND** 如果字段缺失，应视为无效缓存

#### Scenario: 缓存时效性检查
- **GIVEN** 缓存数据有timestamp
- **WHEN** 检查时效性
- **THEN** 应计算: 当前时间 - timestamp < 86400秒(24小时)
- **AND** 如果超时，应视为过期缓存

#### Scenario: 缓存版本检查
- **GIVEN** 缓存数据包含version字段
- **WHEN** 检查版本兼容性
- **THEN** 如果版本不匹配，应视为无效缓存
- **AND** 应执行重新检测

### Requirement: 缓存更新机制
系统 SHALL在适当时候更新缓存数据。

#### Scenario: 手动刷新
- **GIVEN** 用户点击"刷新"按钮
- **WHEN** 执行refresh_office_suites()
- **THEN** 系统应清除现有缓存
- **AND** 应执行重新检测
- **AND** 应保存新缓存结果

#### Scenario: 自动更新
- **GIVEN** 检测到系统变化(如安装了新办公套件)
- **WHEN** 执行检测
- **THEN** 应更新缓存
- **AND** 时间戳应为当前时间

### Requirement: 缓存清除功能
系统 SHALL提供清除缓存的机制。

#### Scenario: 清除指定缓存
- **GIVEN** 缓存数据存在
- **WHEN** 调用clear_office_cache()
- **THEN** 应从配置文件中移除office_suite_cache项
- **AND** 后续初始化应执行完整检测

#### Scenario: 清除所有缓存
- **GIVEN** 配置文件中存在多个缓存
- **WHEN** 调用clear_all_cache()
- **THEN** 应清除所有相关缓存数据
- **AND** 配置文件应保持其他设置不变

### Requirement: OfficeSuiteManager.initialize()行为变更
initialize()方法 SHALL首先尝试加载缓存，只有在缓存无效时才执行检测。

#### Scenario: 新的初始化流程
- **GIVEN** 调用initialize()
- **WHEN** 方法执行
- **THEN** 应执行以下步骤:
  1. 尝试加载缓存
  2. 验证缓存有效性
  3. 如果缓存有效，加载并返回
  4. 如果缓存无效，执行检测
  5. 保存新缓存结果

#### Scenario: 初始化性能目标
- **GIVEN** 缓存有效
- **WHEN** 执行initialize()
- **THEN** 总耗时应 < 100ms

- **GIVEN** 缓存无效
- **WHEN** 执行initialize()
- **THEN** 总耗时可能较长(2-3秒)
- **AND** 但仅在必要时执行

### Requirement: 刷新按钮功能增强
refresh_office_suites()方法 SHALL清除缓存并重新检测。

#### Scenario: 增强刷新逻辑
- **GIVEN** 用户点击"刷新"按钮
- **WHEN** 执行refresh_office_suites()
- **THEN** 应执行以下操作:
  1. 清除缓存数据
  2. 重置初始化状态
  3. 重新执行initialize()
  4. 刷新UI显示
  5. 更新日志信息

#### Scenario: 刷新状态提示
- **GIVEN** 正在刷新
- **WHEN** 用户等待
- **THEN** 日志应显示:
  - "正在清除缓存并重新检测..."
  - "使用缓存的办公套件配置"(如果缓存有效)
  - 或检测结果(如果重新检测)

