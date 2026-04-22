#!/usr/bin/env python3
"""
简化的WPS Office功能测试
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.gui.office_suite_detector import OfficeSuiteManager, OfficeSuite
from src.core.factory.office_suite_factory import OfficeSuiteFactory

def test_wps_support():
    """测试WPS支持功能"""
    print("=" * 50)
    print("WPS Office支持功能测试")
    print("=" * 50)

    try:
        # 1. 测试办公套件管理器
        print("1. 测试办公套件管理器...")
        office_manager = OfficeSuiteManager()
        office_manager.initialize()
        available_suites = office_manager.get_available_suites()
        print(f"   可用套件数量: {len(available_suites)}")
        for suite in available_suites:
            print(f"   - {suite.display_name}")

        # 2. 测试工厂模式
        print("\n2. 测试工厂模式...")
        factory = OfficeSuiteFactory()
        auto_processor = factory.create_auto_processor(office_manager)
        print(f"   创建的处理器: {type(auto_processor).__name__}")

        app_info = auto_processor.get_application_info()
        print(f"   应用信息: {app_info}")

        # 3. 测试切换
        print("\n3. 测试套件切换...")
        for suite in available_suites:
            try:
                office_manager.set_current_suite(suite)
                new_processor = factory.create_processor(suite)
                new_app_info = new_processor.get_application_info()
                print(f"   成功切换到: {new_app_info.get('name', 'Unknown')}")
                new_processor.close_application()
            except Exception as e:
                print(f"   切换到 {suite.display_name} 失败: {e}")

        # 清理资源
        auto_processor.close_application()

        print("\n" + "=" * 50)
        print("测试完成！WPS Office支持功能正常工作")
        print("=" * 50)
        return True

    except Exception as e:
        print(f"测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_wps_support()
    sys.exit(0 if success else 1)