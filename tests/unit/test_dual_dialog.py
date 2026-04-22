#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试双弹框现象 - 使用用户指定的真实文件
"""

import sys
import os

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton
from src.gui.simple_config_manager import SimpleConfigManager
import pandas as pd

class TestMainWindow(QWidget):
    """测试用主窗口 - 模拟真实场景"""
    def __init__(self):
        super().__init__()
        self.text_addition_rules = {}
        self.dundian_mappings = {}
        self.chengbao_mappings = {}
        self.logs = []
        self.dialog_calls = []  # 记录对话框调用
        self.data_path_edit = MockDataPathEdit("KA承保快报.xlsx")
        self.match_table = MockMatchTableWithPhTest()
        self.worker_thread = MockWorkerThreadWithData()

        self.dialog_count = 0
        self.dialog_details = []

    def show_chengbao_input_dialog(self, data_reader, user_input_rows):
        """显示承保趸期数据输入对话框（详细记录）"""
        self.dialog_count += 1
        self.dialog_calls.append(len(user_input_rows))

        # 记录详细信息
        detail = {
            'dialog_number': self.dialog_count,
            'user_input_rows_count': len(user_input_rows),
            'user_input_rows': user_input_rows.copy(),
            'excel_data_shape': data_reader.data.shape if data_reader.data is not None else None,
            'chengbao_data_column_exists': '承保趸期数据' in data_reader.data.columns if data_reader.data is not None else False,
        }

        # 检查R=1的行数
        if detail['chengbao_data_column_exists']:
            chengbao_col = data_reader.data['承保趸期数据']
            r1_rows_in_data = []
            for i, val in enumerate(chengbao_col):
                if pd.notna(val) and '年交SFYP' in str(val):
                    r1_rows_in_data.append(i)
            detail['r1_rows_in_data'] = r1_rows_in_data

        self.dialog_details.append(detail)

        print(f"\n{'='*70}")
        print(f"[Dialog #{self.dialog_count}] 显示承保趸期数据输入对话框")
        print(f"{'='*70}")
        print(f"  需要用户输入的行数: {len(user_input_rows)}")
        print(f"  Excel数据形状: {detail['excel_data_shape']}")
        print(f"  已有承保趸期数据列: {'是' if detail['chengbao_data_column_exists'] else '否'}")

        if 'r1_rows_in_data' in detail:
            print(f"  数据中R=1的行: {detail['r1_rows_in_data']}")

        print(f"  用户输入行: {user_input_rows}")

        # 模拟用户输入
        if not hasattr(self, 'chengbao_user_inputs'):
            self.chengbao_user_inputs = {}

        for row in user_input_rows:
            if row not in self.chengbao_user_inputs:
                # 模拟用户输入"10"
                self.chengbao_user_inputs[row] = "10"
                print(f"    [模拟输入] 行{row} -> 10年交SFYP")

        return self.chengbao_user_inputs

class MockMatchTableWithPhTest:
    """模拟包含ph_test的占位符表格"""
    def __init__(self):
        self.rows = [
            ["ph_险种", "险种"],
            ["ph_支公司", "支公司"],
            ["ph_test", "测试字段"]  # 对应承保趸期映射
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

def test_dual_dialog():
    """测试双弹框现象"""
    print("="*70)
    print("测试双弹框现象")
    print("="*70)

    config_file = "承保.pptcfg"
    excel_file = "KA承保快报.xlsx"

    if not os.path.exists(config_file):
        print(f"\nERROR: 配置文件不存在 - {config_file}")
        return False

    if not os.path.exists(excel_file):
        print(f"\nERROR: Excel文件不存在 - {excel_file}")
        return False

    try:
        # 创建应用
        app = QApplication(sys.argv)

        # 创建主窗口
        print(f"\n[Step 1] 创建主窗口...")
        main_window = TestMainWindow()

        # 创建配置管理器
        print(f"[Step 2] 创建配置管理器...")
        config_manager = SimpleConfigManager()

        # 加载配置
        print(f"[Step 3] 加载配置...")
        print(f"  配置文件: {config_file}")
        print(f"  Excel文件: {excel_file}")
        success = config_manager.load_gui_config(main_window, config_file)

        print(f"\n[Step 4] 加载结果: {'成功' if success else '失败'}")

        # 分析结果
        print(f"\n{'='*70}")
        print("分析结果")
        print(f"{'='*70}")
        print(f"对话框总次数: {main_window.dialog_count}")

        if len(main_window.dialog_details) > 0:
            for detail in main_window.dialog_details:
                print(f"\n  第{detail['dialog_number']}次弹框:")
                print(f"    检测到的用户输入行数: {detail['user_input_rows_count']}")
                print(f"    Excel数据形状: {detail['excel_data_shape']}")
                print(f"    已有承保趸期数据列: {detail['chengbao_data_column_exists']}")
                if 'r1_rows_in_data' in detail:
                    print(f"    数据中R=1的行: {detail['r1_rows_in_data']}")

        print(f"\n{'='*70}")
        if main_window.dialog_count == 1:
            print("[SUCCESS] 只弹了1次对话框！")
        elif main_window.dialog_count == 2:
            print("[INFO] 弹了2次对话框，需要分析原因")
            if len(main_window.dialog_details) >= 2:
                first = main_window.dialog_details[0]
                second = main_window.dialog_details[1]
                print(f"\n  第1次: 检测到{first['user_input_rows_count']}行")
                print(f"  第2次: 检测到{second['user_input_rows_count']}行")

                if second['user_input_rows_count'] > first['user_input_rows_count']:
                    print(f"\n  [问题] 第二次弹框检测到更多行！")
                    print(f"  可能原因: Excel数据行数从{first['excel_data_shape'][0]}行变为{second['excel_data_shape'][0]}行")
        else:
            print(f"[FAIL] 弹了{main_window.dialog_count}次对话框")

        print(f"{'='*70}")

        return success

    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_dual_dialog()
    sys.exit(0 if success else 1)
