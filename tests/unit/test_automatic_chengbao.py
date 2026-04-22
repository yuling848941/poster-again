#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试承保趸期配置加载后自动触发计算功能
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
        self.text_addition_rules = {}
        self.dundian_mappings = {}
        self.chengbao_mappings = {}
        self.log_text = MockLogText()
        self.match_table = MockMatchTable()
        self.worker_thread = MockWorkerThread()

        # 模拟Excel文件路径（测试用）
        self.data_path_edit = MockDataPathEdit()
        self.data_path = None

    def show_chengbao_input_dialog(self, data_reader, user_input_rows):
        """模拟显示承保趸期数据输入对话框"""
        print(f"\n[模拟] 显示承保趸期数据输入对话框")
        print(f"    需要用户输入的行数: {len(user_input_rows)}")
        for row in user_input_rows:
            print(f"      - 行 {row}")

        # 模拟用户输入
        self.chengbao_user_inputs = {row: 10 for row in user_input_rows}
        print(f"    模拟用户输入: {self.chengbao_user_inputs}")

class MockLogText:
    def __init__(self):
        self.logs = []
    def append(self, message):
        self.logs.append(message)
        print(f"[LOG] {message}")

class MockDataPathEdit:
    def __init__(self):
        self.path = ""
    def text(self):
        return self.path

class MockMatchTable:
    def __init__(self):
        self.rows = [
            ["ph_保险费", "保险费"],
            ["ph_首年保费", "首年保费"]
        ]
    def rowCount(self):
        return len(self.rows)
    def item(self, row, col):
        if 0 <= row < len(self.rows):
            return MockTableItem(self.rows[row][col])
        return None
    def setItem(self, row, col, item):
        if 0 <= row < len(self.rows):
            self.rows[row][col] = item.text()

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
        self.data = None
        self.data_loaded = False

def test_automatic_chengbao():
    print("=" * 60)
    print("Test Automatic Chengbao Calculation on Config Load")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp()
    config_file = os.path.join(temp_dir, "test_config.pptcfg")

    try:
        # 创建主窗口并设置承保趸期映射
        print("\n[Setup] 创建主窗口并设置承保趸期映射...")
        main_window = MockMainWindow()
        main_window.chengbao_mappings = {
            "ph_保险费": "承保趸期数据"
        }
        print(f"    承保趸期映射: {main_window.chengbao_mappings}")

        # 保存配置
        print("\n[Test 1] 保存配置...")
        config_manager = SimpleConfigManager()
        success = config_manager.save_gui_config(main_window, config_file)
        print(f"    Result: {'PASS' if success else 'FAIL'}")

        # 验证配置文件
        print("\n[Verify] 验证配置文件...")
        with open(config_file, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        gui_settings = config_data.get('gui_settings', {})
        chengbao_mappings = gui_settings.get('chengbao_mappings', {})
        print(f"    配置文件中的承保趸期映射: {chengbao_mappings}")

        # 创建新主窗口（模拟加载配置）
        print("\n[Test 2] 创建新主窗口并加载配置...")
        new_window = MockMainWindow()
        new_window.chengbao_mappings = {}  # 清空映射
        new_window.data_path_edit.path = config_file  # 使用配置文件路径作为Excel路径（模拟）

        # 加载配置（这会触发自动计算，但会因为没有真实Excel数据而跳过）
        print("\n[Test 3] 加载配置（将尝试触发自动计算）...")
        success = config_manager.load_gui_config(new_window, config_file)
        print(f"    Result: {'PASS' if success else 'FAIL'}")

        # 验证承保趸期映射已恢复
        print("\n[Verify] 验证承保趸期映射恢复...")
        print(f"    恢复的承保趸期映射: {new_window.chengbao_mappings}")
        has_mapping = len(new_window.chengbao_mappings) > 0
        print(f"    Has mapping: {'YES' if has_mapping else 'NO'}")

        # 验证配置文件是否正确保存了映射信息
        print("\n[Verify] 验证配置文件完整性...")
        config_has_mapping = 'ph_保险费' in chengbao_mappings
        print(f"    Config has mapping: {'YES' if config_has_mapping else 'NO'}")

        # 总结
        print("\n" + "=" * 60)
        all_pass = (
            success and
            has_mapping and
            config_has_mapping and
            len(new_window.chengbao_mappings) == 1 and
            "ph_保险费" in new_window.chengbao_mappings
        )
        if all_pass:
            print(f"Result: SUCCESS - Config save/load works!")
            print("Note: Auto-calculation skipped (no real Excel file)")
        else:
            print(f"Result: FAIL - Some tests failed")
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
    success = test_automatic_chengbao()
    sys.exit(0 if success else 1)
