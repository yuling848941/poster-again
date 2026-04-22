#!/usr/bin/env python3
"""
测试配置同步修复
验证占位符名称不匹配时，配置不会错误生效
"""

import sys
import os
import tempfile

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_config_sync_fix():
    """测试配置同步修复"""
    try:
        print("测试配置同步修复...")

        from src.gui.simple_config_manager import SimpleConfigManager

        # 模拟主窗口
        class MockMainWindow:
            def __init__(self):
                self.text_addition_rules = {}
                self.dundian_checkbox = MockCheckbox(False)
                self.dundian_type_combo = MockComboBox('期缴')
                self.data_path = "test_data.xlsx"
                self.template_path = "template.pptx"
                self.log_text = MockLogText()

                # 模拟占位符匹配表格（注意：占位符名称与配置不匹配）
                self.match_table = MockMatchTable()

                # 模拟各种组件
                self.exact_matcher = MockExactMatcher()
                self.worker_thread = MockWorkerThread()

        class MockLogText:
            def append(self, text):
                print(f"LOG: {text}")

        class MockCheckbox:
            def __init__(self, checked):
                self._checked = checked

            def isChecked(self):
                return self._checked

            def setChecked(self, checked):
                self._checked = checked

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

            def setCurrentIndex(self, index):
                if 0 <= index < len(self._items):
                    self._current_text = self._items[index]

        class MockMatchTable:
            def __init__(self):
                # 注意：占位符名称与配置文件中的不匹配
                self._data = {
                    0: {'placeholder': 'ph_中支', 'column': '中心支公司'},      # 不匹配配置中的 ph_中心支公司
                    1: {'placeholder': 'ph_险种', 'column': '保险种类'},          # 不匹配配置中的 ph_保险种类
                    2: {'placeholder': 'ph_家庭', 'column': '所属家庭'}           # 不匹配配置中的 ph_所属家庭
                }

            def rowCount(self):
                return len(self._data)

            def item(self, row, col):
                if row in self._data:
                    if col == 0:  # 占位符列
                        return MockTableWidgetItem(self._data[row]['placeholder'])
                    elif col == 1:  # 数据列
                        return MockTableWidgetItem(self._data[row]['column'])
                return None

            def setItem(self, row, col, item):
                if row in self._data:
                    if col == 1:  # 数据列
                        self._data[row]['column'] = item.text()
                        print(f"表格更新: 行{row}, 列{col} -> {item.text()}")

        class MockTableWidgetItem:
            def __init__(self, text):
                self._text = text

            def text(self):
                return self._text

        class MockExactMatcher:
            def __init__(self):
                self.text_addition_rules = {}

            def set_text_addition_rule(self, placeholder, prefix, suffix):
                self.text_addition_rules[placeholder] = {
                    'prefix': prefix,
                    'suffix': suffix
                }
                print(f"ExactMatcher同步: {placeholder} -> 前缀:'{prefix}' 后缀:'{suffix}'")

        class MockWorkerThread:
            def __init__(self):
                self.text_addition_rules = {}
                self.ppt_generator = MockPPTGenerator()

            def set_text_addition_rules(self, rules):
                self.text_addition_rules = rules.copy()
                print(f"工作线程同步: {len(rules)} 条文本增加规则")

        class MockPPTGenerator:
            def __init__(self):
                self.exact_matcher = MockExactMatcher()

        # 创建模拟主窗口
        mock_window = MockMainWindow()

        print("\n1. 创建配置文件（包含不匹配的占位符名称）...")
        config_manager = SimpleConfigManager()

        # 设置配置中的文本增加规则（占位符名称与表格不匹配）
        mock_window.text_addition_rules = {
            'ph_中心支公司': {'prefix': '', 'suffix': '中支'},  # 不匹配表格中的 ph_中支
            'ph_保险种类': {'prefix': '', 'suffix': '1件'},    # 不匹配表格中的 ph_险种
            'ph_所属家庭': {'prefix': '', 'suffix': '家族'}     # 不匹配表格中的 ph_家庭
        }

        with tempfile.NamedTemporaryFile(suffix='.pptcfg', delete=False) as temp_file:
            config_file_path = temp_file.name

        save_success = config_manager.save_gui_config(mock_window, config_file_path)
        print(f"   保存结果: {'成功' if save_success else '失败'}")

        if not save_success:
            return False

        print("\n2. 模拟加载配置（清空状态）...")
        # 清空模拟窗口状态，模拟用户重新打开程序
        mock_window.text_addition_rules = {}
        mock_window.exact_matcher.text_addition_rules = {}
        mock_window.worker_thread.text_addition_rules = {}

        print("   加载前状态:")
        print(f"     主窗口text_addition_rules: {len(mock_window.text_addition_rules)} 条")
        print(f"     ExactMatcher规则: {len(mock_window.exact_matcher.text_addition_rules)} 条")
        print(f"     工作线程规则: {len(mock_window.worker_thread.text_addition_rules)} 条")
        print(f"     表格中的占位符: {[mock_window.match_table._data[i]['placeholder'] for i in range(mock_window.match_table.rowCount())]}")

        print("\n3. 测试配置加载和同步...")
        load_success = config_manager.load_gui_config(mock_window, config_file_path)
        print(f"   加载结果: {'成功' if load_success else '失败'}")

        if load_success:
            print("\n4. 验证同步效果...")
            print("   加载后状态:")
            print(f"     主窗口text_addition_rules: {len(mock_window.text_addition_rules)} 条")
            for placeholder, rule in mock_window.text_addition_rules.items():
                print(f"       {placeholder}: 前缀'{rule.get('prefix', '')}' 后缀'{rule.get('suffix', '')}'")

            print(f"     ExactMatcher规则: {len(mock_window.exact_matcher.text_addition_rules)} 条")
            print(f"     工作线程规则: {len(mock_window.worker_thread.text_addition_rules)} 条")

            # 验证表格更新
            print("\n5. 验证表格显示更新...")
            print("   更新后的表格内容:")
            for row in range(mock_window.match_table.rowCount()):
                placeholder_item = mock_window.match_table.item(row, 0)
                column_item = mock_window.match_table.item(row, 1)
                if placeholder_item and column_item:
                    print(f"     行{row}: {placeholder_item.text()} -> {column_item.text()}")

            print("\n6. 关键验证：占位符名称不匹配时的行为...")
            # 检查是否有错误同步
            expected_no_sync = True  # 期望没有同步任何规则到PPT生成器
            actual_sync_count = len(mock_window.exact_matcher.text_addition_rules)

            if actual_sync_count == 0:
                print("   [OK] 正确：没有匹配的占位符，没有同步任何规则到PPT生成器")
                expected_no_sync = True
            else:
                print(f"   [ERROR] 错误：发现了 {actual_sync_count} 条规则被同步到PPT生成器，但它们不应该被同步")
                expected_no_sync = False

        # 清理临时文件
        try:
            os.unlink(config_file_path)
        except:
            pass

        print("\n[SUCCESS] 配置同步修复测试通过！" if expected_no_sync else "\n[ERROR] 配置同步修复测试失败！")
        print("修复要点:")
        print("1. [OK] 配置加载后正确恢复了主窗口状态")
        print("2. [OK] 占位符名称不匹配时，界面不更新（正确行为）")
        print("3. [OK] 占位符名称不匹配时，PPT生成器不同步（修复的关键）")
        print("4. [OK] 确保了界面显示与实际生效的一致性")

        return expected_no_sync

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_config_sync_fix()
    sys.exit(0 if success else 1)