#!/usr/bin/env python3
"""
简化的办公套件选择器测试
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.gui.office_suite_detector import OfficeSuiteManager, OfficeSuite
from src.config import ConfigManager

def test_office_suite_detector():
    """测试办公套件检测器"""
    print("=" * 50)
    print("办公套件检测器测试")
    print("=" * 50)

    try:
        # 创建配置管理器
        config_manager = ConfigManager()
        print("SUCCESS: 配置管理器创建成功")

        # 创建办公套件管理器
        office_manager = OfficeSuiteManager(config_manager)
        print("SUCCESS: 办公套件管理器创建成功")

        # 初始化
        office_manager.initialize()
        print("SUCCESS: 办公套件管理器初始化成功")

        # 获取可用套件
        available_suites = office_manager.get_available_suites()
        print(f"SUCCESS: 检测到 {len(available_suites)} 个可用套件:")
        for suite in available_suites:
            info = office_manager.detector.get_suite_info(suite)
            print(f"  - {suite.display_name}: {info.get('name', '未知')} {info.get('version', '')}")

        # 获取当前选择
        current_suite = office_manager.get_current_suite()
        effective_suite = office_manager.get_effective_suite()
        print(f"SUCCESS: 当前选择: {current_suite.display_name}")
        print(f"SUCCESS: 实际使用: {effective_suite.display_name}")

        # 测试切换功能
        if available_suites:
            first_suite = available_suites[0]
            print(f"\n测试切换到: {first_suite.display_name}")
            success = office_manager.set_current_suite(first_suite)
            if success:
                print("SUCCESS: 切换成功")
                new_current = office_manager.get_current_suite()
                print(f"SUCCESS: 新的选择: {new_current.display_name}")
            else:
                print("ERROR: 切换失败")

        # 测试自动检测
        print("\n测试切换到自动检测")
        success = office_manager.set_current_suite(OfficeSuite.AUTO)
        if success:
            print("SUCCESS: 切换到自动检测成功")
            effective = office_manager.get_effective_suite()
            print(f"SUCCESS: 自动检测选择: {effective.display_name}")
        else:
            print("ERROR: 切换到自动检测失败")

        print("\n" + "=" * 50)
        print("所有测试完成！")
        print("=" * 50)
        return True

    except Exception as e:
        print(f"ERROR: 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_office_suite_detector()
    sys.exit(0 if success else 1)