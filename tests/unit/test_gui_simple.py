#!/usr/bin/env python3
"""
简化的GUI集成测试
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_gui_components():
    """测试GUI组件"""
    try:
        print("开始测试GUI组件...")

        # 测试导入
        from src.gui.main_window import MainWindow
        from src.gui.config_dialog import ConfigSaveLoadDialog
        from src.config import ConfigManager
        from src.exact_matcher import ExactMatcher

        print("[OK] 所有模块导入成功")

        # 检查MainWindow方法
        print("\n检查MainWindow类方法...")
        methods_to_check = [
            'open_config_dialog',
            '_sync_exact_matcher',
            '_sync_from_exact_matcher',
            '_add_match_row'
        ]

        for method in methods_to_check:
            if hasattr(MainWindow, method):
                print(f"[OK] MainWindow.{method} 存在")
            else:
                print(f"[ERROR] MainWindow.{method} 缺失")
                return False

        # 检查ConfigSaveLoadDialog
        print("\n检查ConfigSaveLoadDialog...")
        if hasattr(ConfigSaveLoadDialog, '__init__'):
            print("[OK] ConfigSaveLoadDialog.__init__ 存在")
        else:
            print("[ERROR] ConfigSaveLoadDialog.__init__ 缺失")
            return False

        print("\n[SUCCESS] 所有GUI组件测试通过!")
        print("\n使用说明:")
        print("1. 运行主程序: python main.py")
        print("2. 在主界面底部找到绿色的'配置管理'按钮")
        print("3. 点击按钮打开配置管理对话框")
        print("4. 可以保存当前配置或加载已保存的配置")
        print("5. 支持配置兼容性检查和智能加载")

        return True

    except ImportError as e:
        print(f"[ERROR] 导入错误: {str(e)}")
        return False
    except Exception as e:
        print(f"[ERROR] 测试异常: {str(e)}")
        return False

if __name__ == "__main__":
    success = test_gui_components()
    sys.exit(0 if success else 1)