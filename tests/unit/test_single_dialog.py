#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证修复后的承保趸期配置只弹一次对话框
"""

import sys
import os

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from PySide6.QtWidgets import QWidget

from src.gui.simple_config_manager import SimpleConfigManager

class MainWindowWithCounter(QWidget):
    """带对话框计数的主窗口"""
    def __init__(self):
        super().__init__()
        self.text_addition_rules = {}
        self.dundian_mappings = {}
        self.chengbao_mappings = {}
        self.logs = []
        self.dialog_count = 0
        self.user_input_rows_list = []  # 记录每次弹框的用户输入行数

        # 模拟占位符表格
        self.match_table = MockMatchTableWithPhTest()

        # 模拟工作线程
        self.worker_thread = MockWorkerThreadWithData()

        # 模拟数据路径
        self.data_path_edit = MockDataPathEdit("KA承保快报.xlsx")

    def show_chengbao_input_dialog(self, data_reader, user_input_rows):
        """显示承保趸期数据输入对话框（带计数）"""
        self.dialog_count += 1
        self.user_input_rows_list.append(len(user_input_rows))

        print(f"\n[Dialog #{self.dialog_count}] 显示承保趸期数据输入对话框")
        print(f"    需要用户输入的行数: {len(user_input_rows)}")

        # 模拟用户输入
        for row in user_input_rows:
            print(f"      - 行 {row}")

        # 模拟自动填入数据
        if not hasattr(self, 'chengbao_user_inputs'):
            self.chengbao_user_inputs = {}

        for row in user_input_rows:
            self.chengbao_user_inputs[row] = 10  # 模拟用户输入10年

class MockMatchTableWithPhTest:
    """模拟包含ph_test的占位符表格"""
    def __init__(self):
        self.rows = [
            ["ph_test", "测试字段"]
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

class MockWorkerThreadWithData:
    """模拟工作线程"""
    def __init__(self):
        self.matching_rules = {}
        self._ppt_generator = MockPPTGenerator()

    def get_ppt_generator(self):
        return self._ppt_generator

class MockPPTGenerator:
    """模拟PPT生成器"""
    def __init__(self):
        self.rules = {}
        self.data = None

    def add_matching_rule(self, placeholder, column):
        self.rules[placeholder] = column

class MockTableItem:
    def __init__(self, text):
        self._text = text
    def text(self):
        return self._text

class MockDataPathEdit:
    def __init__(self, path):
        self.path = path
    def text(self):
        return self.path

def test_single_dialog():
    """测试只弹一次对话框"""
    print("=" * 70)
    print("Test Single Dialog After Config Load")
    print("=" * 70)

    config_file = "承保.pptcfg"
    excel_file = "KA承保快报.xlsx"

    if not os.path.exists(config_file):
        print(f"ERROR: 配置文件不存在 - {config_file}")
        return False

    if not os.path.exists(excel_file):
        print(f"ERROR: Excel文件不存在 - {excel_file}")
        return False

    try:
        # 创建主窗口
        print("\n[Step 1] 创建主窗口...")
        main_window = MainWindowWithCounter()

        # 创建配置管理器
        print("[Step 2] 创建配置管理器...")
        config_manager = SimpleConfigManager()

        # 加载配置
        print("[Step 3] 加载配置...")
        success = config_manager.load_gui_config(main_window, config_file)

        print(f"\n[Step 4] 加载结果: {'成功' if success else '失败'}")

        # 验证结果
        print(f"\n[Verify] 验证结果...")
        print(f"  对话框显示次数: {main_window.dialog_count}")
        print(f"  各次用户输入行数: {main_window.user_input_rows_list}")

        # 验证承保趸期映射
        print(f"\n  承保趸期映射: {len(main_window.chengbao_mappings)} 条")
        for placeholder, column in main_window.chengbao_mappings.items():
            print(f"    {placeholder} -> {column}")

        # 验证匹配规则
        print(f"\n  PPT生成器匹配规则: {len(main_window.worker_thread._ppt_generator.rules)} 条")
        for placeholder, column in main_window.worker_thread._ppt_generator.rules.items():
            print(f"    {placeholder} -> {column}")

        # 总结
        print("\n" + "=" * 70)
        if success and main_window.dialog_count == 1:
            print("[SUCCESS] 只弹了一次对话框！")
            if len(main_window.user_input_rows_list) > 0:
                rows = main_window.user_input_rows_list[0]
                print(f"  ✓ 用户输入行数: {rows} 行")
                if rows == 1:
                    print(f"  ✓ 正确检测到1行需要输入")
                else:
                    print(f"  ⚠ 预期1行，实际{rows}行")
        elif main_window.dialog_count > 1:
            print(f"[FAIL] 弹了 {main_window.dialog_count} 次对话框")
            for i, rows in enumerate(main_window.user_input_rows_list, 1):
                print(f"  第{i}次: {rows} 行")
        else:
            print("[WARN] 未触发对话框")

        print("=" * 70)

        return success and main_window.dialog_count == 1

    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_single_dialog()
    sys.exit(0 if success else 1)
