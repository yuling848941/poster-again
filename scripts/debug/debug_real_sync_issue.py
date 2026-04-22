#!/usr/bin/env python3
"""
调试实际的配置同步问题
模拟用户的完整工作流程
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def debug_real_sync_issue():
    """调试实际的配置同步问题"""
    try:
        print("调试实际的配置同步问题...")

        # 1. 加载实际的配置文件
        print("\n1. 加载实际配置文件...")
        import json
        config_file = "src/gui_config.pptcfg"

        if not os.path.exists(config_file):
            print(f"配置文件不存在: {config_file}")
            return False

        with open(config_file, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        gui_settings = config_data.get('gui_settings', {})
        text_additions = gui_settings.get('text_additions', {})

        print(f"   配置中的文本增加规则:")
        for placeholder, rule in text_additions.items():
            prefix = rule.get('prefix', '')
            suffix = rule.get('suffix', '')
            print(f"     {placeholder}: 前缀='{prefix}' 后缀='{suffix}'")

        # 2. 获取模板中的占位符
        print("\n2. 获取模板占位符...")
        from src.core.template_processor import TemplateProcessor

        template_file = "10月递交贺报(3).pptx"
        processor = TemplateProcessor()
        processor.load_template(template_file)

        placeholders = processor.find_placeholders()
        print(f"   模板中的占位符:")
        for ph in placeholders:
            print(f"     {ph}")

        # 3. 检查匹配情况
        print("\n3. 检查匹配情况...")
        matched_placeholders = []
        unmatched_config = []

        for config_ph in text_additions.keys():
            if config_ph in placeholders:
                matched_placeholders.append(config_ph)
                print(f"   [匹配] {config_ph}")
            else:
                unmatched_config.append(config_ph)
                # 检查相似占位符
                similar = [p for p in placeholders if any(keyword in p for keyword in config_ph.split('_')[1:])]
                if similar:
                    print(f"   [不匹配] {config_ph} -> 相似占位符: {similar}")
                else:
                    print(f"   [不匹配] {config_ph} -> 无相似占位符")

        # 4. 模拟配置加载过程
        print("\n4. 模拟配置加载过程...")
        from src.gui.simple_config_manager import SimpleConfigManager

        config_manager = SimpleConfigManager()

        # 模拟主窗口
        class MockMainWindow:
            def __init__(self):
                self.text_addition_rules = {}
                self.match_table = MockMatchTable(placeholders)
                self.exact_matcher = MockExactMatcher()
                self.worker_thread = MockWorkerThread()
                self.log_text = MockLogText()

        class MockLogText:
            def append(self, text):
                print(f"   LOG: {text}")

        class MockMatchTable:
            def __init__(self, placeholders):
                self._placeholders = placeholders

            def rowCount(self):
                return len(self._placeholders)

            def item(self, row, col):
                if col == 0:  # 占位符列
                    return MockTableWidgetItem(self._placeholders[row])
                return None

            def setItem(self, row, col, item):
                if col == 1:
                    print(f"   表格更新: {self._placeholders[row]} -> {item.text()}")

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
                    'prefix': prefix, 'suffix': suffix
                }
                print(f"   ExactMatcher同步: {placeholder} -> 前缀:'{prefix}' 后缀:'{suffix}'")

        class MockWorkerThread:
            def __init__(self):
                self.text_addition_rules = {}
                self.ppt_generator = MockPPTGenerator()

            def set_text_addition_rules(self, rules):
                self.text_addition_rules = rules.copy()
                print(f"   工作线程同步: {len(rules)} 条规则")

        class MockPPTGenerator:
            def __init__(self):
                self.exact_matcher = MockExactMatcher()

        mock_window = MockMainWindow()

        # 执行配置加载
        print("\n   执行配置加载...")
        load_success = config_manager.load_gui_config(mock_window, config_file)
        print(f"   加载结果: {'成功' if load_success else '失败'}")

        # 5. 检查最终状态
        print("\n5. 检查最终状态...")
        print(f"   主窗口text_addition_rules: {len(mock_window.text_addition_rules)} 条")
        print(f"   ExactMatcher规则: {len(mock_window.exact_matcher.text_addition_rules)} 条")
        print(f"   工作线程规则: {len(mock_window.worker_thread.text_addition_rules)} 条")

        # 6. 关键检查：是否有错误同步
        print("\n6. 关键检查：错误同步分析...")
        expected_no_sync = len(matched_placeholders) == 0
        actual_sync_count = len(mock_window.exact_matcher.text_addition_rules)

        print(f"   匹配的占位符数量: {len(matched_placeholders)}")
        print(f"   实际同步的规则数量: {actual_sync_count}")

        if expected_no_sync and actual_sync_count == 0:
            print(f"   [OK] 正确：没有匹配的占位符，也没有同步任何规则")
            result = True
        elif not expected_no_sync and actual_sync_count == len(matched_placeholders):
            print(f"   [OK] 正确：只同步了匹配的规则")
            result = True
        else:
            print(f"   [ERROR] 错误：同步了 {actual_sync_count} 条规则，但只应该同步 {len(matched_placeholders)} 条")
            result = False

        return result

    except Exception as e:
        print(f"\n[ERROR] 调试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = debug_real_sync_issue()
    sys.exit(0 if success else 1)