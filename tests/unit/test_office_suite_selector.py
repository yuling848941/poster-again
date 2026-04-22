#!/usr/bin/env python3
"""
测试办公套件选择器功能
"""

import sys
import os
import logging

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication
from src.gui.main_window import MainWindow

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

def test_office_suite_selector():
    """测试办公套件选择器功能"""
    print("=" * 50)
    print("办公套件选择器功能测试")
    print("=" * 50)

    # 创建应用
    app = QApplication(sys.argv)

    # 创建主窗口
    try:
        window = MainWindow()
        print("SUCCESS: 主窗口创建成功")

        # 检查办公套件管理器
        office_manager = window.office_suite_manager
        print("SUCCESS: 办公套件管理器初始化成功")

        # 获取可用套件
        available_suites = office_manager.get_available_suites()
        suite_names = [suite.display_name for suite in available_suites]
        print(f"SUCCESS: 检测到可用套件: {suite_names}")

        # 获取当前选择
        current_suite = office_manager.get_current_suite()
        effective_suite = office_manager.get_effective_suite()
        print(f"SUCCESS: 当前选择: {current_suite.display_name}")
        print(f"SUCCESS: 实际使用: {effective_suite.display_name}")

        # 检查UI组件
        if hasattr(window, 'office_suite_combo'):
            print("SUCCESS: 办公套件下拉菜单创建成功")
            print(f"SUCCESS: 下拉菜单选项数量: {window.office_suite_combo.count()}")

            # 显示选项
            for i in range(window.office_suite_combo.count()):
                text = window.office_suite_combo.itemText(i)
                data = window.office_suite_combo.itemData(i)
                print(f"  - {text}: {data.value if data else None}")
        else:
            print("ERROR: 办公套件下拉菜单未创建")

        if hasattr(window, 'refresh_suites_btn'):
            print("SUCCESS: 刷新按钮创建成功")
        else:
            print("ERROR: 刷新按钮未创建")

        # 显示窗口进行手动测试
        window.show()
        print("\nSUCCESS: 测试窗口已显示，请手动测试以下功能:")
        print("  1. 查看办公套件下拉菜单中的选项")
        print("  2. 尝试切换不同的办公套件")
        print("  3. 点击刷新按钮重新检测")
        print("  4. 查看日志输出")
        print("\n按Ctrl+C关闭测试窗口")

        # 运行应用
        return app.exec()

    except Exception as e:
        print(f"ERROR: 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    exit_code = test_office_suite_selector()
    sys.exit(exit_code)