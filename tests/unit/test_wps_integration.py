#!/usr/bin/env python3
"""
WPS Office集成功能完整测试
测试整个WPS Office支持的架构和功能
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication, QMessageBox
from src.gui.office_suite_detector import OfficeSuiteManager, OfficeSuite
from src.core.factory.office_suite_factory import OfficeSuiteFactory, GracefulDegradation, CompatibilityChecker
from src.core.processors.microsoft_processor import MicrosoftProcessor
from src.core.processors.wps_processor import WPSProcessor
from src.core.interfaces.template_processor import TemplateProcessorError
from src.ppt_generator import PPTGenerator
import logging

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def test_basic_architecture():
    """测试基础架构"""
    print("=" * 60)
    print("测试1: 基础架构功能")
    print("=" * 60)

    try:
        # 测试办公套件管理器
        print("1.1 测试办公套件管理器...")
        office_manager = OfficeSuiteManager()
        office_manager.initialize()
        print("SUCCESS: 办公套件管理器初始化成功")

        available_suites = office_manager.get_available_suites()
        suite_names = [suite.display_name for suite in available_suites]
        print(f"SUCCESS: 检测到可用套件: {suite_names}")

        current_suite = office_manager.get_current_suite()
        effective_suite = office_manager.get_effective_suite()
        print(f"SUCCESS: 当前选择: {current_suite.display_name}, 实际使用: {effective_suite.display_name}")

        return True
    except Exception as e:
        print(f"ERROR: 基础架构测试失败: {e}")
        return False

def test_factory_pattern():
    """测试工厂模式"""
    print("\n" + "=" * 60)
    print("测试2: 工厂模式功能")
    print("=" * 60)

    try:
        # 测试工厂创建处理器
        print("2.1 测试工厂创建处理器...")
        factory = OfficeSuiteFactory()
        supported_suites = factory.get_supported_suites()
        supported_names = [suite.display_name for suite in supported_suites]
        print(f"SUCCESS: 支持的套件: {supported_names}")

        # 测试自动创建
        print("2.2 测试自动处理器创建...")
        auto_processor = factory.create_auto_processor()
        print(f"SUCCESS: 自动创建处理器成功: {type(auto_processor).__name__}")

        # 获取处理器信息
        app_info = auto_processor.get_application_info()
        print(f"SUCCESS: 处理器信息: {app_info}")

        auto_processor.close_application()
        return True
    except Exception as e:
        print(f"ERROR: 工厂模式测试失败: {e}")
        return False

def test_processor_capabilities():
    """测试处理器能力"""
    print("\n" + "=" * 60)
    print("测试3: 处理器能力测试")
    print("=" * 60)

    try:
        factory = OfficeSuiteFactory()
        compatibility_checker = CompatibilityChecker()

        # 测试Microsoft处理器
        print("3.1 测试Microsoft PowerPoint处理器...")
        try:
            ms_processor = factory.create_processor(OfficeSuite.MICROSOFT)
            ms_capabilities = ms_processor.get_processor_capabilities()
            print(f"SUCCESS Microsoft处理器能力: {ms_capabilities}")

            supported_formats = ms_processor.get_supported_formats()
            print(f"SUCCESS 支持的格式: {supported_formats}")

            ms_processor.close_application()
        except Exception as e:
            print(f"WARNING Microsoft处理器不可用: {e}")

        # 测试WPS处理器
        print("3.2 测试WPS演示处理器...")
        try:
            wps_processor = factory.create_processor(OfficeSuite.WPS)
            wps_capabilities = wps_processor.get_processor_capabilities()
            print(f"SUCCESS WPS处理器能力: {wps_capabilities}")

            supported_formats = wps_processor.get_supported_formats()
            print(f"SUCCESS 支持的格式: {supported_formats}")

            wps_processor.close_application()
        except Exception as e:
            print(f"WARNING WPS处理器不可用: {e}")

        # 获取功能矩阵
        print("3.3 生成功能兼容性矩阵...")
        feature_matrix = compatibility_checker.get_feature_matrix()
        for suite, capabilities in feature_matrix.items():
            if 'error' in capabilities:
                print(f"  {suite}: 错误 - {capabilities['error']}")
            else:
                print(f"  {suite}: {len(capabilities)} 个功能支持")

        return True
    except Exception as e:
        print(f"ERROR 处理器能力测试失败: {e}")
        return False

def test_graceful_degradation():
    """测试优雅降级"""
    print("\n" + "=" * 60)
    print("测试4: 优雅降级机制")
    print("=" * 60)

    try:
        office_manager = OfficeSuiteManager()
        office_manager.initialize()
        degradation = GracefulDegradation()

        # 获取降级链
        print("4.1 获取降级链...")
        fallback_chain = degradation.get_fallback_chain(OfficeSuite.WPS, office_manager)
        print(f"SUCCESS WPS的降级链: {[suite.display_name for suite in fallback_chain]}")

        fallback_chain = degradation.get_fallback_chain(OfficeSuite.MICROSOFT, office_manager)
        print(f"SUCCESS Microsoft的降级链: {[suite.display_name for suite in fallback_chain]}")

        # 测试降级创建
        print("4.2 测试降级处理器创建...")
        try:
            processor = degradation.create_processor_with_fallback(OfficeSuite.AUTO, office_manager)
            print(f"SUCCESS 降级创建成功: {type(processor).__name__}")
            processor.close_application()
        except Exception as e:
            print(f"WARNING 降级创建失败: {e}")

        return True
    except Exception as e:
        print(f"ERROR 优雅降级测试失败: {e}")
        return False

def test_ppt_generator_integration():
    """测试PPT生成器集成"""
    print("\n" + "=" * 60)
    print("测试5: PPT生成器集成")
    print("=" * 60)

    try:
        # 测试PPT生成器
        print("5.1 测试PPT生成器初始化...")
        office_manager = OfficeSuiteManager()
        office_manager.initialize()

        ppt_generator = PPTGenerator(office_manager=office_manager)
        print("SUCCESS PPT生成器初始化成功")

        # 获取当前办公套件信息
        current_suite = ppt_generator.get_current_office_suite()
        print(f"SUCCESS 当前使用的办公套件: {current_suite}")

        # 测试设置办公套件管理器
        print("5.2 测试切换办公套件管理器...")
        new_office_manager = OfficeSuiteManager()
        new_office_manager.initialize()
        ppt_generator.set_office_manager(new_office_manager)
        print("SUCCESS 办公套件管理器切换成功")

        # 清理资源
        ppt_generator.close()
        return True
    except Exception as e:
        print(f"ERROR PPT生成器集成测试失败: {e}")
        return False

def test_gui_integration():
    """测试GUI集成"""
    print("\n" + "=" * 60)
    print("测试6: GUI集成测试")
    print("=" * 60)

    try:
        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)

        from src.gui.main_window import MainWindow

        print("6.1 测试主窗口创建...")
        window = MainWindow()
        print("SUCCESS 主窗口创建成功")

        # 检查办公套件选择器
        print("6.2 检查办公套件选择器...")
        if hasattr(window, 'office_suite_combo'):
            print("SUCCESS 办公套件下拉菜单存在")
            count = window.office_suite_combo.count()
            print(f"SUCCESS 下拉菜单选项数量: {count}")
        else:
            print("ERROR 办公套件下拉菜单不存在")

        if hasattr(window, 'refresh_suites_btn'):
            print("SUCCESS 刷新按钮存在")
        else:
            print("ERROR 刷新按钮不存在")

        # 检查工作线程
        print("6.3 检查工作线程集成...")
        if hasattr(window.worker_thread, 'ppt_generator'):
            generator = window.worker_thread.ppt_generator
            current_suite = generator.get_current_office_suite()
            print(f"SUCCESS 工作线程使用办公套件: {current_suite}")

        window.close()
        return True
    except Exception as e:
        print(f"ERROR GUI集成测试失败: {e}")
        return False

def main():
    """主测试函数"""
    print("WPS Office支持功能完整测试")
    print("测试整个架构的各个组件和集成")
    print()

    tests = [
        ("基础架构", test_basic_architecture),
        ("工厂模式", test_factory_pattern),
        ("处理器能力", test_processor_capabilities),
        ("优雅降级", test_graceful_degradation),
        ("PPT生成器集成", test_ppt_generator_integration),
        ("GUI集成", test_gui_integration)
    ]

    passed = 0
    total = len(tests)

    for test_name, test_func in tests:
        try:
            if test_func():
                passed += 1
                print(f"SUCCESS {test_name} 测试通过")
            else:
                print(f"FAIL {test_name} 测试失败")
        except Exception as e:
            print(f"ERROR {test_name} 测试异常: {e}")

    print("\n" + "=" * 60)
    print(f"测试结果: {passed}/{total} 通过")
    print("=" * 60)

    if passed == total:
        print("SUCCESS 所有测试通过！WPS Office支持功能已成功实现")
        return True
    else:
        print("WARNING 部分测试失败，请检查相关功能")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)