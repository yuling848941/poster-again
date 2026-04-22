#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试承保趵期配置保存和加载功能
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
    """Mock main window"""
    def __init__(self):
        # Initialize various attributes
        self.text_addition_rules = {
            "ph_保费金额": {"prefix": "¥", "suffix": "元"}
        }
        self.dundian_mappings = {
            "ph_投保人": "期趸数据"
        }
        self.chengbao_mappings = {
            "ph_保险费": "承保趸期数据"
        }
        self.log_text = MockLogText()
        self.match_table = MockMatchTable()
        self.worker_thread = MockWorkerThread()

class MockLogText:
    """Mock log text"""
    def __init__(self):
        self.logs = []

    def append(self, message):
        self.logs.append(message)
        print(f"[LOG] {message}")

class MockMatchTable:
    """Mock match table"""
    def __init__(self):
        self.rows = [
            ["ph_保费金额", "保费金额"],
            ["ph_保险费", "保险费"],
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
    """Mock table item"""
    def __init__(self, text):
        self.text = text

class MockWorkerThread:
    """Mock worker thread"""
    def __init__(self):
        self.matching_rules = {}
        self.ppt_generator = MockPPTGenerator()

    def get_ppt_generator(self):
        return self.ppt_generator

class MockPPTGenerator:
    """Mock PPT generator"""
    def __init__(self):
        self.rules = {}

    def add_matching_rule(self, placeholder, column):
        self.rules[placeholder] = column
        print(f"[PPTGenerator] Add matching rule: {placeholder} -> {column}")

def test_config_save_and_load():
    """Test config save and load"""
    print("=" * 60)
    print("Test Chengbao Config Save and Load")
    print("=" * 60)

    # Create temp directory
    temp_dir = tempfile.mkdtemp()
    config_file = os.path.join(temp_dir, "test_config.pptcfg")

    try:
        # Create mock main window
        main_window = MockMainWindow()

        # Create config manager
        config_manager = SimpleConfigManager()

        # Test 1: Save config
        print("\n[Test 1] Save config...")
        success = config_manager.save_gui_config(main_window, config_file)
        if success:
            print("    [OK] Config saved successfully")
        else:
            print("    [FAIL] Config save failed")
            return False

        # Verify saved file
        print("\n[Verify] Check saved config file...")
        with open(config_file, 'r', encoding='utf-8') as f:
            saved_config = json.load(f)

        gui_settings = saved_config.get('gui_settings', {})
        print(f"    Text addition rules: {len(gui_settings.get('text_additions', {}))}")
        print(f"    Dundian mappings: {len(gui_settings.get('dundian_mappings', {}))}")
        print(f"    Chengbao mappings: {len(gui_settings.get('chengbao_mappings', {}))}")

        # Test 2: Config preview
        print("\n[Test 2] Get config preview...")
        preview = config_manager.get_config_preview(config_file)
        print("    Preview info:")
        for line in preview.split('\n'):
            print(f"      {line}")

        # Test 3: Load config to new window
        print("\n[Test 3] Load config to new window...")
        new_window = MockMainWindow()
        # Clear existing data
        new_window.text_addition_rules = {}
        new_window.dundian_mappings = {}
        new_window.chengbao_mappings = {}

        success = config_manager.load_gui_config(new_window, config_file)
        if success:
            print("    [OK] Config loaded successfully")
        else:
            print("    [FAIL] Config load failed")
            return False

        # Verify loaded data
        print("\n[Verify] Check loaded config...")
        print(f"    Text addition rules: {len(new_window.text_addition_rules)}")
        for placeholder, rule in new_window.text_addition_rules.items():
            print(f"      {placeholder}: {rule}")

        print(f"    Dundian mappings: {len(new_window.dundian_mappings)}")
        for placeholder, column in new_window.dundian_mappings.items():
            print(f"      {placeholder} -> {column}")

        print(f"    Chengbao mappings: {len(new_window.chengbao_mappings)}")
        for placeholder, column in new_window.chengbao_mappings.items():
            print(f"      {placeholder} -> {column}")

        # Test 4: Verify data consistency
        print("\n[Test 4] Verify data consistency...")
        checks = []

        # Check text addition rules
        if new_window.text_addition_rules == main_window.text_addition_rules:
            checks.append("[OK] Text addition rules match")
        else:
            checks.append("[FAIL] Text addition rules mismatch")

        # Check dundian mappings
        if new_window.dundian_mappings == main_window.dundian_mappings:
            checks.append("[OK] Dundian mappings match")
        else:
            checks.append("[FAIL] Dundian mappings mismatch")

        # Check chengbao mappings
        if new_window.chengbao_mappings == main_window.chengbao_mappings:
            checks.append("[OK] Chengbao mappings match")
        else:
            checks.append("[FAIL] Chengbao mappings mismatch")

        for check in checks:
            print(f"    {check}")

        # Summary
        print("\n" + "=" * 60)
        all_passed = all("[OK]" in check for check in checks)
        if all_passed:
            print("[SUCCESS] All tests passed!")
        else:
            print("[FAIL] Some tests failed")
        print("=" * 60)

        return all_passed

    finally:
        # Clean up temp directory
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            print(f"\n[Cleanup] Removed temp dir: {temp_dir}")

if __name__ == "__main__":
    success = test_config_save_and_load()
    sys.exit(0 if success else 1)
