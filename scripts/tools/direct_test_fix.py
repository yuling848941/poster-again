#!/usr/bin/env python3
"""
直接测试配置修复逻辑
使用简单的模拟来验证修复是否正确
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def direct_test_fix():
    """直接测试配置修复逻辑"""
    try:
        print("=== 直接测试配置修复逻辑 ===")

        from src.gui.simple_config_manager import SimpleConfigManager

        # 模拟用户的实际场景
        print("\n1. 模拟用户场景...")

        # 配置文件内容（来自实际文件）
        text_additions = {
            "ph_中心支公司": {"prefix": "", "suffix": "中支"},
            "ph_保险种类": {"prefix": "", "suffix": "1件"},
            "ph_所属家庭": {"prefix": "", "suffix": "家族"}
        }

        # 模板占位符（来自实际模板）
        placeholders = ["ph_标准编号", "ph_姓名", "ph_递交字段", "ph_中支", "ph_保险种类", "ph_所属家庭"]

        print(f"   配置中的文本增加规则: {list(text_additions.keys())}")
        print(f"   模板中的占位符: {placeholders}")

        # 模拟主窗口和表格
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
                elif col == 1:  # 数据列 - 模拟显示
                    return MockTableWidgetItem("")  # 空的数据列
                return None

            def setItem(self, row, col, item):
                if col == 1:
                    placeholder = self._placeholders[row]
                    print(f"   表格更新: {placeholder} -> {item.text()}")

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
        config_manager = SimpleConfigManager()

        print("\n2. 执行配置加载和同步...")

        # 执行配置加载的核心逻辑
        print(f"   加载前状态: ExactMatcher规则 {len(mock_window.exact_matcher.text_addition_rules)} 条")

        # 模拟 _restore_gui_state 的关键逻辑
        gui_settings = {"text_additions": text_additions}
        restored_count = len(text_additions)

        # 恢复文本增加设置
        for placeholder, rule in text_additions.items():
            mock_window.text_addition_rules[placeholder] = {
                'prefix': rule.get('prefix', ''),
                'suffix': rule.get('suffix', '')
            }

        # 执行表格更新和同步逻辑
        matched_text_additions = config_manager._update_match_table_display(mock_window, text_additions)

        # 只同步匹配成功的规则
        if matched_text_additions:
            config_manager._sync_to_ppt_generator(mock_window, matched_text_additions)
        else:
            print("   没有匹配的文本增加规则，跳过PPT生成器同步")

        print(f"   加载后状态: ExactMatcher规则 {len(mock_window.exact_matcher.text_addition_rules)} 条")

        print("\n3. 验证同步结果...")

        # 分析匹配结果
        matched_placeholders = []
        for placeholder in placeholders:
            if placeholder in text_additions:
                matched_placeholders.append(placeholder)
                print(f"   [匹配] {placeholder}")
            else:
                # 检查相似性
                similar = [p for p in text_additions.keys() if any(keyword in placeholder for keyword in p.split('_')[1:])]
                if similar:
                    print(f"   [不匹配] {placeholder} <- 配置中有相似的: {similar}")
                else:
                    print(f"   [不匹配] {placeholder}")

        print(f"\n4. 关键验证...")

        # 验证同步的规则数量
        expected_sync_count = len(matched_placeholders)
        actual_sync_count = len(mock_window.exact_matcher.text_addition_rules)

        print(f"   期望同步的规则数: {expected_sync_count}")
        print(f"   实际同步的规则数: {actual_sync_count}")

        # 检查具体同步的规则
        print(f"   实际同步的规则:")
        for placeholder, rule in mock_window.exact_matcher.text_addition_rules.items():
            print(f"     {placeholder}: 后缀'{rule.get('suffix', '')}'")

        # 关键检查：ph_中支 是否被错误同步
        if "ph_中支" in mock_window.exact_matcher.text_addition_rules:
            rule = mock_window.exact_matcher.text_addition_rules["ph_中支"]
            if rule.get('suffix') == "中支":
                print(f"   [错误] ph_中支 被错误同步了 '中支' 后缀!")
                return False

        if "ph_中心支公司" not in mock_window.exact_matcher.text_addition_rules:
            print(f"   [正确] ph_中心支公司 没有被同步（因为模板中没有这个占位符）")

        if actual_sync_count == expected_sync_count:
            print(f"   [正确] 同步的规则数量正确")
            return True
        else:
            print(f"   [错误] 同步的规则数量不正确")
            return False

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = direct_test_fix()
    print(f"\n{'测试通过' if success else '测试失败'}")
    sys.exit(0 if success else 1)