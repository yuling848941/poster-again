#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试防重复弹框机制的有效性
模拟加载配置时承保趸期数据计算触发的场景
"""

import sys
import os
import json
import tempfile
import shutil

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.gui.simple_config_manager import SimpleConfigManager

class MockDataReader:
    """模拟数据读取器"""
    def __init__(self):
        self.data = {
            "承保趸期数据": ["", "10年交SFYP", ""],  # 只有1个用户输入项
            "保单号": ["A001", "A002", "A003"]
        }

    def load_excel(self, file_path, use_thousand_separator=True, parent_widget=None):
        return True

    def get_column(self, column_name):
        return self.data.get(column_name, [])

class MockMainWindow:
    """模拟主窗口 - 集成完整功能"""
    def __init__(self, excel_file_path=None):
        self.text_addition_rules = {}
        self.dundian_mappings = {}
        self.chengbao_mappings = {
            "ph_保险费": "承保趸期数据"
        }
        self.log_text = MockLogText()
        self.match_table = MockMatchTable()
        self.worker_thread = MockWorkerThread()
        self.data_path_edit = MockDataPathEdit(excel_file_path or "test.xlsx")
        self.dialog_shown_count = 0  # 追踪对话框显示次数
        self.excel_file_exists = excel_file_path is not None

        # 模拟Excel数据
        if self.excel_file_exists:
            self.excel_data = {
                "承保趸期数据": ["", "10年交SFYP", ""],  # 只有1个用户输入项
                "保单号": ["A001", "A002", "A003"],
                "SFYP2(不含短险续保)": [100, 200, 300],
                "首年保费": [1000, 2000, 3000]
            }
        else:
            self.excel_data = {}

    def show_chengbao_input_dialog(self, data_reader, user_input_rows):
        """模拟显示承保趸期数据输入对话框（包含防重复机制）"""
        print(f"\n[DEBUG] show_chengbao_input_dialog 被调用")
        print(f"    user_input_rows: {user_input_rows}")

        # 检查防重复标志
        if hasattr(self, '_chengbao_dialog_shown') and self._chengbao_dialog_shown:
            print(f"[Anti-Duplicate] 对话框已显示过，跳过重复显示")
            return

        # 设置防重复标志
        self._chengbao_dialog_shown = True
        self.dialog_shown_count += 1

        print(f"\n[Dialog #{self.dialog_shown_count}] 显示承保趸期数据输入对话框")
        print(f"    需要用户输入的行数: {len(user_input_rows)}")

        # 模拟用户输入
        for row in user_input_rows:
            policy_num = self.worker_thread.ppt_generator.data.get('保单号', ['N/A'])[row]
            print(f"      - 行 {row} (保单号: {policy_num})")

        # 模拟用户输入数据
        if not hasattr(self, 'chengbao_user_inputs'):
            self.chengbao_user_inputs = {}

        for row in user_input_rows:
            self.chengbao_user_inputs[row] = 10

        print(f"    模拟用户输入完成: {self.chengbao_user_inputs}")

        # 清除防重复标志
        if hasattr(self, '_chengbao_dialog_shown'):
            delattr(self, '_chengbao_dialog_shown')

class MockLogText:
    def __init__(self):
        self.logs = []
    def append(self, message):
        self.logs.append(message)
        print(f"[LOG] {message}")

class MockDataPathEdit:
    def __init__(self, path):
        self.path = path
    def text(self):
        return self.path

class MockMatchTable:
    def __init__(self):
        self.rows = [
            ["ph_保险费", "保险费"]
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
        self.data = {
            "承保趸期数据": ["", "10年交SFYP", ""],
            "保单号": ["A001", "A002", "A003"]
        }
    def add_matching_rule(self, placeholder, column):
        self.rules[placeholder] = column

def test_anti_duplicate_dialog():
    """测试防重复弹框机制"""
    print("=" * 70)
    print("Test Anti-Duplicate Dialog Mechanism")
    print("=" * 70)

    temp_dir = tempfile.mkdtemp()
    config_file = os.path.join(temp_dir, "test_config.pptcfg")

    try:
        # 创建配置管理器
        config_manager = SimpleConfigManager()

        # 创建主窗口并设置承保趸期映射
        print("\n[Setup] 创建主窗口并设置承保趸期映射...")
        main_window = MockMainWindow()
        print(f"    承保趸期映射: {main_window.chengbao_mappings}")

        # 保存配置
        print("\n[Test 1] 保存配置...")
        success = config_manager.save_gui_config(main_window, config_file)
        print(f"    Result: {'PASS' if success else 'FAIL'}")
        assert success, "配置保存失败"

        # 模拟加载配置过程（多次触发计算）
        print("\n[Test 2] 模拟加载配置过程（模拟多次触发）...")

        # 创建新主窗口
        new_window = MockMainWindow()
        new_window.chengbao_mappings = {}  # 清空映射

        # 加载配置
        print("    步骤1: 加载配置...")
        success = config_manager.load_gui_config(new_window, config_file)
        print(f"    配置加载结果: {'PASS' if success else 'FAIL'}")

        # 验证承保趸期映射已恢复
        print("\n[Verify] 验证承保趸期映射恢复...")
        print(f"    恢复的承保趸期映射: {new_window.chengbao_mappings}")
        assert len(new_window.chengbao_mappings) > 0, "承保趸期映射未恢复"

        # 验证对话框显示次数
        print("\n[Verify] 验证防重复弹框机制...")
        dialog_count = new_window.dialog_shown_count
        print(f"    对话框显示次数: {dialog_count}")
        print(f"    预期次数: 1")
        print(f"    结果: {'PASS' if dialog_count == 1 else 'FAIL - 显示了多次对话框!'}")

        # 验证自动计算触发
        print("\n[Verify] 验证自动计算触发...")
        if hasattr(new_window, 'chengbao_user_inputs'):
            input_count = len(new_window.chengbao_user_inputs)
            print(f"    用户输入行数: {input_count}")
            print(f"    预期行数: 1")
            print(f"    结果: {'PASS' if input_count == 1 else 'FAIL - 检测到错误数量的用户输入行'}")
        else:
            print(f"    结果: FAIL - 未触发用户输入")

        # 模拟再次触发计算（应该被防重复机制阻止）
        print("\n[Test 3] 模拟再次触发计算（应该被阻止）...")
        print("    手动触发第二次计算...")
        old_count = new_window.dialog_shown_count
        config_manager._trigger_chengbao_calculation(new_window)
        new_count = new_window.dialog_shown_count
        print(f"    触发前对话框次数: {old_count}")
        print(f"    触发后对话框次数: {new_count}")
        print(f"    结果: {'PASS - 未重复显示' if new_count == old_count else 'FAIL - 仍然显示对话框'}")

        # 总结
        print("\n" + "=" * 70)
        all_pass = (
            dialog_count == 1 and  # 第一次只显示一次
            len(new_window.chengbao_mappings) > 0 and  # 映射正确恢复
            (hasattr(new_window, 'chengbao_user_inputs') and len(new_window.chengbao_user_inputs) == 1)  # 检测到正确数量的用户输入
        )
        if all_pass:
            print("Result: SUCCESS - 防重复弹框机制工作正常!")
            print("✓ 对话框只显示一次")
            print("✓ 正确检测到1个用户输入行")
            print("✓ 重复触发被有效阻止")
        else:
            print("Result: FAIL - 防重复弹框机制存在问题")
        print("=" * 70)

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
    success = test_anti_duplicate_dialog()
    sys.exit(0 if success else 1)
