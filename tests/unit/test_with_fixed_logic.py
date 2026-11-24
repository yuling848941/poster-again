#!/usr/bin/env python3
"""
使用修复后的逻辑验证文本增加规则只应用到匹配的占位符

核心验证点：配置中不存在的占位符（如 ph_中支）不应被应用文本增加规则。
"""

import os
import sys
import json

# 将项目根目录加入 sys.path，确保 src 包可导入
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
sys.path.insert(0, PROJECT_ROOT)

from src.gui.simple_config_manager import SimpleConfigManager


def test_with_fixed_logic():
    """使用修复后的逻辑验证文本增加规则匹配"""
    try:
        print("=== 验证文本增加规则只应用到匹配的占位符 ===")

        # 1. 加载实际配置
        config_file = os.path.join(PROJECT_ROOT, "samples/presets/递交.pptcfg")
        if not os.path.exists(config_file):
            print(f"[ERROR] 配置文件不存在: {config_file}")
            return False

        with open(config_file, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        gui_settings = config_data.get('gui_settings', {})
        text_additions = gui_settings.get('text_additions', {})

        print("   配置中的文本增加规则:")
        for placeholder, rule in text_additions.items():
            suffix = rule.get('suffix', '')
            print(f"     {placeholder}: 后缀='{suffix}'")

        # 2. 模拟匹配表格：配置中已有的占位符 + 一个配置中不存在的 ph_中支
        placeholders = ["ph_中心支公司", "ph_保险种类", "ph_所属家庭", "ph_中支"]

        class MockTableItem:
            def __init__(self, text):
                self._text = text

            def text(self):
                return self._text

        class MockTable:
            def __init__(self):
                self.rows = [[ph, ph.replace("ph_", "")] for ph in placeholders]

            def rowCount(self):
                return len(self.rows)

            def item(self, row, col):
                if 0 <= row < len(self.rows):
                    return MockTableItem(self.rows[row][col])
                return None

            def setItem(self, row, col, item):
                if 0 <= row < len(self.rows):
                    self.rows[row][col] = item

        class MockWindow:
            def __init__(self):
                self.match_table = MockTable()

        # 3. 使用修复后的配置加载逻辑
        print("\n3. 使用修复后的配置加载逻辑...")
        config_manager = SimpleConfigManager()
        matched = config_manager._update_match_table_display(MockWindow(), text_additions)

        print(f"   匹配的规则数量: {len(matched)}")
        for placeholder, rule in matched.items():
            print(f"     {placeholder}: 后缀'{rule.get('suffix', '')}'")

        # 4. 关键验证
        print("\n4. 验证结果...")
        success = True

        # 配置中不存在的占位符不应被应用
        if "ph_中支" in matched:
            print("   [ERROR] ph_中支 被错误应用了文本增加规则")
            success = False
        else:
            print("   [OK] ph_中支 没有被应用文本增加规则")

        # 配置中存在的占位符且表格中存在，应被应用
        for expected in ["ph_保险种类", "ph_所属家庭"]:
            if expected in matched:
                print(f"   [OK] {expected} 正确被应用")
            else:
                print(f"   [ERROR] {expected} 应该被应用但没有")
                success = False

        print(f"\n{'修复验证成功' if success else '修复验证失败'}")
        return success

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = test_with_fixed_logic()
    sys.exit(0 if success else 1)
