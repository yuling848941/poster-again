#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试配置管理界面显示承保趸期设置
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
    """模拟主窗口"""
    def __init__(self):
        # 初始化各种属性
        self.text_addition_rules = {
            "ph_保费金额": {"prefix": "¥", "suffix": "元"}
        }
        self.dundian_mappings = {
            "ph_投保人": "期趸数据"
        }
        self.chengbao_mappings = {
            "ph_保险费": "承保趸期数据",
            "ph_首年保费": "承保趸期数据"
        }
        self.log_text = MockLogText()
        self.match_table = MockMatchTable()
        self.worker_thread = MockWorkerThread()

class MockLogText:
    """模拟日志文本"""
    def __init__(self):
        self.logs = []

    def append(self, message):
        self.logs.append(message)
        print(f"[LOG] {message}")

class MockMatchTable:
    """模拟匹配表格"""
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
    """模拟表格项"""
    def __init__(self, text):
        self.text = text

class MockWorkerThread:
    """模拟工作线程"""
    def __init__(self):
        self.matching_rules = {}
        self.ppt_generator = MockPPTGenerator()

    def get_ppt_generator(self):
        return self.ppt_generator

class MockPPTGenerator:
    """模拟PPT生成器"""
    def __init__(self):
        self.rules = {}

    def add_matching_rule(self, placeholder, column):
        self.rules[placeholder] = column
        print(f"[PPTGenerator] 添加匹配规则: {placeholder} -> {column}")

def test_config_dialog_display():
    """测试配置管理界面显示"""
    print("=" * 60)
    print("测试配置管理界面显示承保趸期设置")
    print("=" * 60)

    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    config_file = os.path.join(temp_dir, "test_config.pptcfg")

    try:
        # 创建模拟主窗口
        main_window = MockMainWindow()

        # 创建配置管理器
        config_manager = SimpleConfigManager()

        # 保存配置
        print("\n[步骤1] 保存配置...")
        success = config_manager.save_gui_config(main_window, config_file)
        if success:
            print("    [OK] 配置保存成功")
        else:
            print("    [FAIL] 配置保存失败")
            return False

        # 读取配置文件
        print("\n[步骤2] 读取配置文件...")
        with open(config_file, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        gui_settings = config_data.get('gui_settings', {})
        print(f"    配置文件结构:")
        print(f"      text_additions: {list(gui_settings.get('text_additions', {}).keys())}")
        print(f"      dundian_mappings: {list(gui_settings.get('dundian_mappings', {}).keys())}")
        print(f"      chengbao_mappings: {list(gui_settings.get('chengbao_mappings', {}).keys())}")

        # 测试配置预览
        print("\n[步骤3] 获取配置预览（模拟配置管理界面显示）...")
        preview = config_manager.get_config_preview(config_file)
        print("    预览信息:")
        print("    " + "-" * 50)
        for line in preview.split('\n'):
            print(f"    {line}")
        print("    " + "-" * 50)

        # 验证承保趸期信息是否显示
        print("\n[步骤4] 验证承保趸期信息显示...")
        chengbao_lines = [line for line in preview.split('\n') if '承保趸期' in line]
        for line in chengbao_lines:
            print(f"    [显示] {line}")

        if len(chengbao_lines) >= 2:
            print("    [OK] 承保趸期信息显示正确")
        else:
            print("    [FAIL] 承保趸期信息显示不完整")
            return False

        # 测试加载配置
        print("\n[步骤5] 测试配置加载...")
        new_window = MockMainWindow()
        new_window.text_addition_rules = {}
        new_window.dundian_mappings = {}
        new_window.chengbao_mappings = {}

        success = config_manager.load_gui_config(new_window, config_file)
        if success:
            print("    [OK] 配置加载成功")
        else:
            print("    [FAIL] 配置加载失败")
            return False

        # 验证承保趸期映射已恢复
        print("\n[步骤6] 验证承保趸期映射恢复...")
        if new_window.chengbao_mappings:
            print(f"    [OK] 承保趸期映射已恢复: {len(new_window.chengbao_mappings)} 条")
            for placeholder, column in new_window.chengbao_mappings.items():
                print(f"      - {placeholder} -> {column}")
        else:
            print("    [FAIL] 承保趸期映射未恢复")
            return False

        # 总结
        print("\n" + "=" * 60)
        print("[SUCCESS] 配置管理界面显示测试通过！")
        print("=" * 60)

        return True

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理临时目录
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            print(f"\n[清理] 已删除临时目录: {temp_dir}")

if __name__ == "__main__":
    success = test_config_dialog_display()
    sys.exit(0 if success else 1)
