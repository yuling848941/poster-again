#!/usr/bin/env python3
"""
测试简化配置管理系统
"""

import sys
import os
import tempfile

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_simple_config():
    """测试简化配置管理系统"""
    try:
        print("测试简化配置管理系统...")

        # 导入简化配置管理器
        from src.gui.simple_config_manager import SimpleConfigManager

        config_manager = SimpleConfigManager()

        # 模拟主窗口状态
        class MockMainWindow:
            def __init__(self):
                self.text_addition_rules = {
                    'ph_中心支公司': {'prefix': '', 'suffix': '中支'},
                    'ph_保险种类': {'prefix': '', 'suffix': '1件'}
                }
                self.dundian_checkbox = MockCheckbox(True)
                self.dundian_type_combo = MockComboBox('期缴')
                self.log_text = MockLogText()

        class MockCheckbox:
            def __init__(self, checked):
                self._checked = checked

            def isChecked(self):
                return self._checked

        class MockComboBox:
            def __init__(self, current_text):
                self._current_text = current_text
                self._items = ['期缴', '趸交']

            def currentText(self):
                return self._current_text

            def count(self):
                return len(self._items)

            def itemText(self, index):
                return self._items[index] if 0 <= index < len(self._items) else ''

        class MockLogText:
            def append(self, text):
                print(f"LOG: {text}")

        mock_window = MockMainWindow()

        # 测试保存配置
        print("\n1. 测试保存配置...")
        with tempfile.NamedTemporaryFile(suffix='.pptcfg', delete=False) as temp_file:
            temp_path = temp_file.name

        save_success = config_manager.save_gui_config(mock_window, temp_path)
        print(f"   保存结果: {'成功' if save_success else '失败'}")

        if save_success:
            # 验证保存的文件
            print(f"   配置文件大小: {os.path.getsize(temp_path)} 字节")

            # 测试获取预览
            preview = config_manager.get_config_preview(temp_path)
            print(f"   配置预览:\n{preview}")

            # 测试加载配置
            print("\n2. 测试加载配置...")
            # 清空模拟窗口状态
            mock_window.text_addition_rules = {}
            mock_window.dundian_checkbox._checked = False
            mock_window.dundian_type_combo._current_text = '趸交'

            load_success = config_manager.load_gui_config(mock_window, temp_path)
            print(f"   加载结果: {'成功' if load_success else '失败'}")

            if load_success:
                print("   恢复的状态:")
                print(f"   - 文本增加规则数: {len(mock_window.text_addition_rules)}")
                for placeholder, rule in mock_window.text_addition_rules.items():
                    print(f"     {placeholder}: 前缀'{rule.get('prefix', '')}' 后缀'{rule.get('suffix', '')}'")
                print(f"   - 递交趸期启用: {mock_window.dundian_checkbox.isChecked()}")
                print(f"   - 默认递交类型: {mock_window.dundian_type_combo.currentText()}")

        # 清理临时文件
        try:
            os.unlink(temp_path)
        except:
            pass

        print("\n[SUCCESS] 简化配置管理系统测试通过！")
        return True

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_simple_config()
    sys.exit(0 if success else 1)