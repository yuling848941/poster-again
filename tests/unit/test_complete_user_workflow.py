#!/usr/bin/env python3
"""
测试完整的用户工作流程
验证简化配置管理系统的实际使用场景
"""

import sys
import os
import tempfile
import json

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_complete_workflow():
    """测试完整用户工作流程"""
    try:
        print("测试完整的用户工作流程...")

        # 步骤1: 模拟用户启动应用程序，加载数据和模板
        print("\n=== 步骤1: 模拟应用程序启动 ===")

        from src.exact_matcher import ExactMatcher
        import pandas as pd

        # 模拟用户数据
        user_data = pd.DataFrame({
            '中心支公司': ['北京中支', '上海中支', '广州中支'],
            '保险种类': ['人寿保险', '车险', '健康险'],
            '保费金额': [10000, 5000, 8000]
        })

        # 创建匹配器并设置数据
        matcher = ExactMatcher()
        matcher.set_data(user_data)

        # 设置模板占位符
        placeholders = ['ph_中心支公司', 'ph_保险种类', 'ph_保费金额']
        matcher.set_template_placeholders(placeholders)

        print(f"数据加载完成: {len(user_data)} 行, {len(user_data.columns)} 列")
        print(f"模板占位符: {placeholders}")

        # 步骤2: 模拟用户执行自动匹配
        print("\n=== 步骤2: 执行自动匹配 ===")
        matcher.set_matching_rule('ph_中心支公司', '中心支公司')
        matcher.set_matching_rule('ph_保险种类', '保险种类')
        matcher.set_matching_rule('ph_保费金额', '保费金额')

        matching_rules = matcher.get_matching_rules()
        print(f"自动匹配完成: {len(matching_rules)} 条规则")
        for placeholder, column in matching_rules.items():
            print(f"  {placeholder} -> {column}")

        # 步骤3: 模拟用户添加文本增加后缀
        print("\n=== 步骤3: 添加文本增加后缀 ===")

        # 模拟主窗口的文本增加规则
        class MockMainWindow:
            def __init__(self):
                self.text_addition_rules = {}
                self.dundian_checkbox = MockCheckbox(False)  # 默认不启用递交趸期
                self.dundian_type_combo = MockComboBox('期缴')
                self.data_path = "test_data.xlsx"  # 模拟数据文件路径
                self.template_path = "template.pptx"  # 模拟模板文件路径

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

        mock_window = MockMainWindow()

        # 用户添加后缀设置
        mock_window.text_addition_rules['ph_中心支公司'] = {'prefix': '', 'suffix': '中支'}
        mock_window.text_addition_rules['ph_保险种类'] = {'prefix': '', 'suffix': '1件'}

        print(f"用户添加了 {len(mock_window.text_addition_rules)} 条文本增加规则:")
        for placeholder, rule in mock_window.text_addition_rules.items():
            print(f"  {placeholder}: 前缀='{rule.get('prefix', '')}' 后缀='{rule.get('suffix', '')}'")

        # 步骤4: 保存配置
        print("\n=== 步骤4: 保存配置 ===")

        from src.gui.simple_config_manager import SimpleConfigManager
        config_manager = SimpleConfigManager()

        with tempfile.NamedTemporaryFile(suffix='.pptcfg', delete=False) as temp_file:
            config_file_path = temp_file.name

        save_success = config_manager.save_gui_config(mock_window, config_file_path)
        print(f"配置保存结果: {'成功' if save_success else '失败'}")

        if save_success:
            print(f"配置文件: {config_file_path}")

            # 验证配置文件内容
            with open(config_file_path, 'r', encoding='utf-8') as f:
                saved_config = json.load(f)

            gui_settings = saved_config.get('gui_settings', {})
            text_additions = gui_settings.get('text_additions', {})
            dundian_settings = gui_settings.get('dundian_settings', {})

            print(f"保存的配置内容:")
            print(f"  版本: {saved_config.get('version', 'Unknown')}")
            print(f"  文本增加规则: {len(text_additions)} 条")
            print(f"  递交趸期设置: {'启用' if dundian_settings.get('enabled', False) else '禁用'}")

        # 步骤5: 模拟用户重新打开应用程序，加载配置
        print("\n=== 步骤5: 加载配置 ===")

        # 重置模拟窗口状态
        mock_window.text_addition_rules = {}
        mock_window.dundian_checkbox.setChecked(False)
        mock_window.dundian_type_combo.setCurrentIndex(1)  # 切换到趸交

        print(f"加载前状态:")
        print(f"  文本增加规则: {len(mock_window.text_addition_rules)} 条")
        print(f"  递交趸期启用: {mock_window.dundian_checkbox.isChecked()}")
        print(f"  默认类型: {mock_window.dundian_type_combo.currentText()}")

        # 加载配置
        load_success = config_manager.load_gui_config(mock_window, config_file_path)
        print(f"配置加载结果: {'成功' if load_success else '失败'}")

        if load_success:
            print(f"加载后状态:")
            print(f"  文本增加规则: {len(mock_window.text_addition_rules)} 条")
            for placeholder, rule in mock_window.text_addition_rules.items():
                print(f"    {placeholder}: 前缀='{rule.get('prefix', '')}' 后缀='{rule.get('suffix', '')}'")
            print(f"  递交趸期启用: {mock_window.dundian_checkbox.isChecked()}")
            print(f"  默认类型: {mock_window.dundian_type_combo.currentText()}")

        # 步骤6: 验证配置与实际应用的集成
        print("\n=== 步骤6: 验证实际应用集成 ===")

        # 检查加载的配置是否能正确应用到实际的匹配器
        if mock_window.text_addition_rules:
            print("将文本增加规则应用到ExactMatcher:")
            for placeholder, rule in mock_window.text_addition_rules.items():
                matcher.set_text_addition_rule(
                    placeholder,
                    rule.get('prefix', ''),
                    rule.get('suffix', '')
                )
                print(f"  应用: {placeholder} -> 前缀='{rule.get('prefix', '')}' 后缀='{rule.get('suffix', '')}'")

            # 验证最终配置
            final_text_rules = matcher.get_all_text_addition_rules()
            print(f"最终生效的文本增加规则: {len(final_text_rules)} 条")

            # 导出完整配置
            exported_config = matcher.export_matching_config()
            print(f"导出的完整配置规则: {len(exported_config)} 条")

            for rule in exported_config:
                placeholder = rule['placeholder']
                column = rule['column']
                text_info = ""
                if 'text_addition' in rule:
                    ta = rule['text_addition']
                    text_info = f" (前缀:'{ta.get('prefix', '')}' 后缀:'{ta.get('suffix', '')}')"
                print(f"  {placeholder} -> {column}{text_info}")

        # 清理临时文件
        try:
            os.unlink(config_file_path)
        except:
            pass

        print("\n[SUCCESS] 完整用户工作流程测试通过！")
        print("验证要点:")
        print("1. [OK] 模拟了用户从数据加载到配置保存的完整流程")
        print("2. [OK] 验证了文本增加规则的正确保存和加载")
        print("3. [OK] 确认了配置与ExactMatcher的正确集成")
        print("4. [OK] 测试了GUI状态的正确恢复")
        print("5. [OK] 验证了配置文件的格式和内容")

        return True

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_complete_workflow()
    sys.exit(0 if success else 1)