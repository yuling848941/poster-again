# WPS Office支持 - 架构设计

## 1. 架构概览

### 1.1 设计目标
- 在现有Microsoft Office支持基础上扩展WPS Office支持
- 保持API兼容性和用户体验一致性
- 实现办公套件的自动检测和优雅降级
- 最小化代码重复和维护成本

### 1.2 核心原则
- **抽象优先**: 通过抽象层隐藏办公套件差异
- **工厂模式**: 动态创建合适的处理器实例
- **策略模式**: 根据环境选择最优处理策略
- **适配器模式**: 统一不同COM接口的差异

## 2. 系统架构

### 2.1 整体架构图

```
┌─────────────────────────────────────────┐
│              GUI界面层                    │
├─────────────────────────────────────────┤
│          PPTGenerator (协调器)            │
├─────────────────────────────────────────┤
│    OfficeSuiteFactory (工厂类)           │
├─────────────────────────────────────────┤
│    ITemplateProcessor (抽象接口)          │
├───────────────┬─────────────────────────┤
│PowerPointImpl │   WPSImpl (新增)         │
│(Microsoft)    │   (WPS Office)          │
├───────────────┴─────────────────────────┤
│            COM接口层                      │
├─────────────────────────────────────────┤
│  PowerPoint.Application │ WPS.Application │
└─────────────────────────────────────────┘
```

### 2.2 关键组件设计

#### 2.2.1 办公套件检测器 (OfficeDetector)
```python
class OfficeDetector:
    """办公套件自动检测器"""

    @staticmethod
    def detect_available_suites() -> List[OfficeSuite]:
        """检测系统中可用的办公套件"""

    @staticmethod
    def get_preferred_suite() -> Optional[OfficeSuite]:
        """获取用户首选或最优的办公套件"""

    @staticmethod
    def validate_suite_availability(suite: OfficeSuite) -> bool:
        """验证特定办公套件的可用性"""
```

#### 2.2.2 抽象处理器接口 (ITemplateProcessor)
```python
from abc import ABC, abstractmethod

class ITemplateProcessor(ABC):
    """模板处理器抽象接口"""

    @abstractmethod
    def connect_to_application(self) -> bool:
        """连接到办公套件应用程序"""

    @abstractmethod
    def load_template(self, template_path: str) -> bool:
        """加载模板文件"""

    @abstractmethod
    def find_placeholders(self) -> List[str]:
        """查找占位符"""

    @abstractmethod
    def replace_placeholders(self, data: Dict[str, str]) -> bool:
        """替换占位符内容"""

    @abstractmethod
    def save_presentation(self, output_path: str) -> bool:
        """保存演示文稿"""

    @abstractmethod
    def export_to_images(self, output_dir: str, format: str, quality: float) -> bool:
        """导出为图片"""
```

#### 2.2.3 WPS处理器实现 (WPSProcessor)
```python
class WPSProcessor(ITemplateProcessor):
    """WPS Office模板处理器实现"""

    def __init__(self):
        self.wps_app = None
        self.presentation = None
        self.logger = logging.getLogger(__name__)

    def connect_to_application(self) -> bool:
        """连接到WPS应用程序"""
        try:
            pythoncom.CoInitialize()
            # WPS演示程序的ProgID通常是"WPS.Application"或"KWps.Application"
            self.wps_app = win32com.client.Dispatch("WPS.Application")
            return self._validate_connection()
        except Exception as e:
            self.logger.error(f"连接WPS失败: {e}")
            return False
```

#### 2.2.4 办公套件工厂 (OfficeSuiteFactory)
```python
class OfficeSuiteFactory:
    """办公套件处理器工厂"""

    @staticmethod
    def create_processor(suite_type: OfficeSuite) -> ITemplateProcessor:
        """根据办公套件类型创建处理器"""
        if suite_type == OfficeSuite.MICROSOFT:
            return MicrosoftProcessor()
        elif suite_type == OfficeSuite.WPS:
            return WPSProcessor()
        else:
            raise ValueError(f"不支持的办公套件类型: {suite_type}")

    @staticmethod
    def create_auto_processor() -> ITemplateProcessor:
        """自动检测并创建最适合的处理器"""
        available_suites = OfficeDetector.detect_available_suites()
        preferred = OfficeDetector.get_preferred_suite()

        if preferred in available_suites:
            return OfficeSuiteFactory.create_processor(preferred)
        elif available_suites:
            return OfficeSuiteFactory.create_processor(available_suites[0])
        else:
            raise RuntimeError("未检测到可用的办公套件")
```

## 3. COM接口差异处理

### 3.1 主要差异点

| 功能特性 | Microsoft PowerPoint | WPS演示 | 处理策略 |
|---------|---------------------|---------|----------|
| ProgID | PowerPoint.Application | WPS.Application | 配置化映射 |
| 文件格式常量 | ppSaveAsDefault=11 | 类似常量 | 动态获取 |
| 图片导出 | Export方法 | 类似方法 | 适配器模式 |
| 对象模型 | Presentations集合 | 类似集合 | 统一接口 |

### 3.2 适配器实现策略

#### 3.2.1 常量映射适配器
```python
class ConstantsAdapter:
    """常量映射适配器"""

    def __init__(self, suite_type: OfficeSuite):
        self.suite_type = suite_type
        self._init_constants()

    def _init_constants(self):
        if self.suite_type == OfficeSuite.MICROSOFT:
            self.SAVE_AS_DEFAULT = 11
            self.SAVE_AS_PNG = 18
            self.SAVE_AS_JPG = 17
        elif self.suite_type == OfficeSuite.WPS:
            # WPS的常量值可能不同，需要测试确定
            self.SAVE_AS_DEFAULT = self._detect_wps_constant("default")
            self.SAVE_AS_PNG = self._detect_wps_constant("png")
            self.SAVE_AS_JPG = self._detect_wps_constant("jpg")
```

#### 3.2.2 方法调用适配器
```python
class MethodAdapter:
    """方法调用适配器"""

    def __init__(self, app_object, suite_type: OfficeSuite):
        self.app = app_object
        self.suite_type = suite_type

    def export_to_image(self, presentation, path, format_type):
        """统一的图片导出方法"""
        if self.suite_type == OfficeSuite.MICROSOFT:
            return presentation.Export(path, format_type, 0, 0)
        elif self.suite_type == OfficeSuite.WPS:
            # WPS可能使用不同的方法签名
            return self._wps_export(presentation, path, format_type)
```

## 4. 配置管理设计

### 4.1 配置结构
```json
{
  "office_suite": {
    "preferred": "auto",  // auto, microsoft, wps
    "fallback_enabled": true,
    "detection_timeout": 5000
  },
  "wps_specific": {
    "prog_id": "WPS.Application",
    "connection_retries": 3,
    "compatibility_mode": true
  }
}
```

### 4.2 配置管理器扩展
```python
class OfficeConfigManager:
    """办公套件配置管理器"""

    def __init__(self, config_manager: ConfigManager):
        self.config = config_manager

    def get_preferred_suite(self) -> OfficeSuite:
        """获取用户首选办公套件"""

    def is_fallback_enabled(self) -> bool:
        """是否启用降级支持"""

    def get_suite_config(self, suite: OfficeSuite) -> Dict:
        """获取特定办公套件的配置"""
```

## 5. GUI界面设计

### 5.1 办公套件选择界面
- 添加办公套件选择下拉菜单
- 显示检测到的可用套件
- 提供自动检测选项
- 显示当前使用的套件状态

### 5.2 状态反馈设计
- 实时显示当前使用的办公套件
- 在切换时提供进度提示
- 错误情况下提供清晰的错误信息

## 6. 错误处理和降级策略

### 6.1 连接失败处理
```python
class GracefulDegradation:
    """优雅降级处理器"""

    @staticmethod
    def handle_connection_failure(preferred: OfficeSuite) -> Optional[ITemplateProcessor]:
        """处理连接失败，尝试降级方案"""
        available = OfficeDetector.detect_available_suites()

        # 尝试其他可用套件
        for suite in available:
            if suite != preferred:
                try:
                    processor = OfficeSuiteFactory.create_processor(suite)
                    if processor.connect_to_application():
                        logging.info(f"成功降级到 {suite.value}")
                        return processor
                except Exception as e:
                    logging.warning(f"降级到 {suite.value} 失败: {e}")

        return None
```

### 6.2 兼容性检查
```python
class CompatibilityChecker:
    """兼容性检查器"""

    @staticmethod
    def check_feature_support(processor: ITemplateProcessor, feature: str) -> bool:
        """检查特定功能是否支持"""

    @staticmethod
    def get_supported_formats(processor: ITemplateProcessor) -> List[str]:
        """获取支持的文件格式列表"""
```

## 7. 测试策略

### 7.1 单元测试
- 每个处理器的独立测试
- COM接口模拟测试
- 错误处理测试

### 7.2 集成测试
- 端到端功能测试
- 多套件环境测试
- 降级场景测试

### 7.3 兼容性测试
- 不同WPS版本测试
- 文件格式兼容性测试
- API差异测试

## 8. 部署和迁移

### 8.1 向后兼容性
- 现有代码无需修改
- 配置文件自动升级
- 渐进式功能启用

### 8.2 迁移策略
- 默认保持原有行为
- 用户主动启用WPS支持
- 提供平滑的迁移路径

## 9. 性能考虑

### 9.1 检测优化
- 缓存检测结果
- 异步检测可选套件
- 最小化启动时间影响

### 9.2 资源管理
- 按需加载办公套件
- 及时释放COM资源
- 连接池管理可选

## 10. 未来扩展性

### 10.1 其他办公套件支持
- LibreOffice支持预留接口
- 在线办公服务支持架构
- 插件化办公套件架构

### 10.2 功能增强
- 批量处理多套件支持
- 跨套件格式转换
- 智能套件选择算法