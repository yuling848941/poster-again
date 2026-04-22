#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证承保趸期配置功能
"""

import sys
import os
import json
import tempfile
import shutil

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.gui.simple_config_manager import SimpleConfigManager

class MockMainWindow:
    def __init__(self):
        self.text_addition_rules = {
            "ph_保费金额": {"prefix": "$", "suffix": ""}
        }
        self.dundian_mappings = {
            "ph_投保人": "dundian_data"
        }
        self.chengbao_mappings = {
            "ph_保险费": "chengbao_data",
            "ph_首年保费": "chengbao_data"
        }
        self.log_text = MockLogText()
        self.match_table = MockMatchTable()
        self.worker_thread = MockWorkerThread()

class MockLogText:
    def __init__(self):
        self.logs = []
    def append(self, message):
        self.logs.append(message)

class MockMatchTable:
    def __init__(self):
        self.rows = [
            ["ph_保费金额", "保费金额"],
            ["ph_保险费", "保险费"],
            ["ph_首年保费", "首年保费"],
            ["ph_投保人", "投保人"]
        ]
    def rowCount(self):
        return len(self.rows)
    def item(self, row, col):
        if 0 <= row < len(self.rows):
            return MockTableItem(self.rows[row][col])
        return None
    def setItem(self, row, col, item):
        if 0 <= row < len(self.rows):
            self.rows[row][col] = item.text

class MockTableItem:
    def __init__(self, text):
        self._text = text

    def text(self):
        return self._text

class MockWorkerThread:
    def __init__(self):
        self.matching_rules = {}
        self.ppt_generator = MockPPTGenerator()
    def get_ppt_generator(self):
        return self.ppt_generator

class MockPPTGenerator:
    def __init__(self):
        self.rules = {}
    def add_matching_rule(self, placeholder, column):
        self.rules[placeholder] = column

def test():
    print("=" * 60)
    print("Verify Chengbao Config Functionality")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    config_file = os.path.join(temp_dir, "test_config.pptcfg")

    try:
        main_window = MockMainWindow()
        config_manager = SimpleConfigManager()

        # Test save
        print("\n[Test 1] Save config...")
        success = config_manager.save_gui_config(main_window, config_file)
        print(f"    Result: {'PASS' if success else 'FAIL'}")

        # Verify file structure
        print("\n[Test 2] Verify config structure...")
        with open(config_file, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        gui_settings = config_data.get('gui_settings', {})
        text_additions = gui_settings.get('text_additions', {})
        dundian_mappings = gui_settings.get('dundian_mappings', {})
        chengbao_mappings = gui_settings.get('chengbao_mappings', {})

        print(f"    Text additions: {len(text_additions)} items")
        print(f"    Dundian mappings: {len(dundian_mappings)} items")
        print(f"    Chengbao mappings: {len(chengbao_mappings)} items")

        # Test preview
        print("\n[Test 3] Config preview...")
        preview = config_manager.get_config_preview(config_file)
        has_chengbao = '承保趸期设置' in preview
        print(f"    Has Chengbao info: {'YES' if has_chengbao else 'NO'}")

        # Test load
        print("\n[Test 4] Load config...")
        new_window = MockMainWindow()
        new_window.text_addition_rules = {}
        new_window.dundian_mappings = {}
        new_window.chengbao_mappings = {}

        success = config_manager.load_gui_config(new_window, config_file)
        print(f"    Result: {'PASS' if success else 'FAIL'}")

        # Verify restoration
        print("\n[Test 5] Verify restoration...")
        chengbao_count = len(new_window.chengbao_mappings)
        print(f"    Chengbao mappings restored: {chengbao_count}")

        # Summary
        print("\n" + "=" * 60)
        all_pass = (
            len(text_additions) > 0 and
            len(dundian_mappings) > 0 and
            len(chengbao_mappings) > 0 and
            has_chengbao and
            chengbao_count > 0
        )
        print(f"Result: {'SUCCESS - All tests passed!' if all_pass else 'FAIL - Some tests failed'}")
        print("=" * 60)

        return all_pass

    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    success = test()
    sys.exit(0 if success else 1)
