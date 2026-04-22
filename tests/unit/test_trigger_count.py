#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证承保趸期计算只触发一次
"""

import sys
import os

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.gui.simple_config_manager import SimpleConfigManager

class TestMainWindow:
    """测试用主窗口"""
    def __init__(self):
        self.text_addition_rules = {}
        self.dundian_mappings = {}
        self.chengbao_mappings = {}
        self.logs = []
        self.dialog_calls = []

    def show_chengbao_input_dialog(self, data_reader, user_input_rows):
        """记录对话框调用"""
        self.dialog_calls.append(len(user_input_rows))
        print(f"[Dialog] 需要用户输入的行数: {len(user_input_rows)}")

class TestDataPathEdit:
    def __init__(self, path):
        self.path = path
        self.call_count = 0
    def text(self):
        self.call_count += 1
        return self.path

def test_trigger_count():
    """测试承保趸期计算触发次数"""
    print("=" * 70)
    print("Test Chengbao Calculation Trigger Count")
    print("=" * 70)

    config_manager = SimpleConfigManager()
    main_window = TestMainWindow()

    # 使用真实存在的文件作为路径
    excel_file = "KA承保快报.xlsx"
    if not os.path.exists(excel_file):
        print(f"[WARN] Excel文件不存在，将使用虚假路径进行测试: {excel_file}")
        excel_file = "fake_path.xlsx"

    main_window.data_path_edit = TestDataPathEdit(excel_file)

    # 模拟第一次触发
    print("\n[Test 1] 第一次调用 _trigger_chengbao_calculation()...")
    config_manager._trigger_chengbao_calculation(main_window)

    # 模拟第二次触发
    print("\n[Test 2] 第二次调用 _trigger_chengbao_calculation()...")
    config_manager._trigger_chengbao_calculation(main_window)

    # 模拟第三次触发
    print("\n[Test 3] 第三次调用 _trigger_chengbao_calculation()...")
    config_manager._trigger_chengbao_calculation(main_window)

    # 检查调用次数
    print(f"\n[Result] 实际触发次数: {config_manager._calculation_call_count}")
    print(f"预期触发次数: 1")

    if config_manager._calculation_call_count == 1:
        print("\n[SUCCESS] 防重复触发机制工作正常！")
    else:
        print(f"\n[FAIL] 触发次数为 {config_manager._calculation_call_count}，不符合预期")

    print("=" * 70)

if __name__ == "__main__":
    test_trigger_count()
