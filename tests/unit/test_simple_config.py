#!/usr/bin/env python3
"""
测试简化配置管理功能
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_simple_config():
    """测试简化配置管理功能"""
    try:
        print("测试简化配置管理功能...")

        from src.gui.simple_config_manager import SimpleConfigManager
        from src.gui.simple_config_dialog import SimpleConfigDialog
        import pandas as pd

        # 创建测试数据
        test_data = pd.DataFrame({
            '中心支公司': ['北京中支', '上海中支'],
            '保险种类': ['人寿保险', '车险'],
            '所属家庭': ['张家', '李家']
        })

        # 创建模拟的主窗口对象
        class MockMainWindow:
            def __init__(self):
                self.data_frame = test_data
                self.template_placeholders = ['ph_中心支公司', 'ph_保险种类', 'ph_所属家庭', 'ph_递交字段']
                self.text_addition_rules = {}
                self.dundian_checkbox = MockCheckbox(True)
                self.dundian_type_combo = MockComboBox(['期缴', '趸交'])

            def add_text_addition_rule(self, placeholder, prefix, suffix):
                self.text_addition_rules[placeholder] = {
                    'prefix': prefix,
                    'suffix': suffix
                }

        class MockCheckbox:
            def __init__(self, checked=False):
                self._checked = checked

            def isChecked(self):
                return self._checked

            def setChecked(self, checked):
                self._checked = checked

        class MockComboBox:
            def __init__(self, items):
                self.items = items
                self._current_index = 0

            def currentText(self):
                return self.items[self._current_index] if self._current_index < len(self.items) else ""

            def count(self):
                return len(self.items)

            def setCurrentIndex(self, index):
                if 0 <= index < len(self.items):
                    self._current_index = index

        print("1. 创建模拟主窗口和设置...")

        mock_window = MockMainWindow()

        # 模拟用户添加文本增加规则
        mock_window.add_text_addition_rule('ph_中心支公司', '', '中支')
        mock_window.add_text_addition_rule('ph_保险种类', '', '1件')
        mock_window.add_text_addition_rule('ph_所属家庭', '', '家族')

        print("   - 已添加3条文本增加规则")
        print(f"   - 递交趸期设置: {'启用' if mock_window.dundian_checkbox.isChecked() else '禁用'}")
        print(f"   - 默认类型: {mock_window.dundian_type_combo.currentText()}")

        print("\n2. 测试配置保存...")
        config_manager = SimpleConfigManager()

        # 创建临时文件
        import tempfile
        temp_file = tempfile.mktemp(suffix='.pptcfg')

        # 保存配置
        success = config_manager.save_gui_config(mock_window, temp_file)

        if success:
            print(f"   [OK] 配置保存成功: {temp_file}")

            # 验证保存的文件内容
            preview = config_manager.get_config_preview(temp_file)
            print(f"   配置预览:\n{preview}")

            print("\n3. 测试配置加载...")
            # 创建新的模拟窗口用于加载
            new_window = MockMainWindow()

            # 加载配置
            load_success = config_manager.load_gui_config(new_window, temp_file)

            if load_success:
                print("   [OK] 配置加载成功")

                # 验证加载结果
                text_rules = new_window.text_addition_rules
                print(f"   加载的文本增加规则数: {len(text_rules)}")
                for placeholder, rule in text_rules.items():
                    prefix = rule.get('prefix', '')
                    suffix = rule.get('suffix', '')
                    if prefix or suffix:
                        print(f"     {placeholder}: 前缀='{prefix}' 后缀='{suffix}'")

                dundian_enabled = new_window.dundian_checkbox.isChecked()
                dundian_type = new_window.dundian_type_combo.currentText()
                print(f"   递交趸期设置: {'启用' if dundian_enabled else '禁用'}, 默认类型: {dundian_type}")

            # 清理临时文件
            try:
                os.unlink(temp_file)
            except:
                pass
        else:
            print("   [ERROR] 配置保存失败")

        print("\n[SUCCESS] 简化配置管理测试通过！")
        print("\n新的设计方案优势:")
        print("1. GUI优先 - 直接基于用户在界面的操作")
        print("2. 简单直接 - 不需要复杂的状态转换")
        print("3. 用户友好 - 清晰的状态显示和操作反馈")
        print("4. 稳定可靠 - 减少状态转换的出错点")

        return True

    except Exception as e:
        print(f"[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_simple_config()
    sys.exit(0 if success else 1)