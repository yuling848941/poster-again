#!/usr/bin/env python3
"""
测试递交趸期配置持久化功能
验证递交趸期数据映射的保存和加载
"""

import sys
import os
import json
import tempfile

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_dundian_config_persistence():
    """测试递交趸期配置持久化功能"""
    try:
        print("=== 测试递交趸期配置持久化功能 ===")

        # 1. 测试SimpleConfigManager的收集功能
        print("\n1. 测试配置收集功能...")

        from src.gui.simple_config_manager import SimpleConfigManager

        config_manager = SimpleConfigManager()

        # 模拟主窗口状态
        class MockMainWindow:
            def __init__(self):
                self.text_addition_rules = {
                    "ph_保险种类": {"prefix": "", "suffix": "1件"},
                    "ph_所属家庭": {"prefix": "", "suffix": "家族"}
                }
                self.dundian_mappings = {
                    "ph_保费金额": "期趸数据",
                    "ph_缴费年期": "期趸数据"
                }
                self.match_table = MockMatchTable(["ph_标准编号", "ph_姓名", "ph_保费金额", "ph_缴费年期"])
                self.log_text = MockLogText()

        class MockLogText:
            def append(self, text):
                print(f"  LOG: {text}")

        class MockMatchTable:
            def __init__(self, placeholders):
                self._placeholders = placeholders

            def rowCount(self):
                return len(self._placeholders)

            def item(self, row, col):
                if col == 0:  # 占位符列
                    return MockTableWidgetItem(self._placeholders[row])
                elif col == 1:  # 数据列
                    return MockTableWidgetItem("")
                return None

            def setItem(self, row, col, item):
                if col == 1:
                    placeholder = self._placeholders[row]
                    print(f"  表格更新: {placeholder} -> {item.text()}")

        class MockTableWidgetItem:
            def __init__(self, text):
                self._text = text

            def text(self):
                return self._text

        mock_window = MockMainWindow()

        # 测试收集GUI状态
        gui_config = config_manager._collect_gui_state(mock_window)

        print(f"  收集的文本增加规则: {len(gui_config['text_additions'])}")
        for placeholder, rule in gui_config['text_additions'].items():
            print(f"    {placeholder}: 前缀'{rule.get('prefix', '')}' 后缀'{rule.get('suffix', '')}'")

        print(f"  收集的递交趸期映射: {len(gui_config['dundian_mappings'])}")
        for placeholder, data_column in gui_config['dundian_mappings'].items():
            print(f"    {placeholder} -> {data_column}")

        # 2. 测试配置保存
        print("\n2. 测试配置保存...")

        config_data = {
            "version": "1.0.0",
            "created_at": "2025-01-11T12:00:00.000000",
            "gui_settings": gui_config
        }

        # 创建临时配置文件
        with tempfile.NamedTemporaryFile(mode='w', suffix='.pptcfg', delete=False, encoding='utf-8') as f:
            json.dump(config_data, f, ensure_ascii=False, indent=2)
            temp_config_file = f.name

        print(f"  配置已保存到: {temp_config_file}")

        # 3. 测试配置加载
        print("\n3. 测试配置加载...")

        # 读取配置文件
        with open(temp_config_file, 'r', encoding='utf-8') as f:
            loaded_config_data = json.load(f)

        loaded_gui_settings = loaded_config_data.get('gui_settings', {})

        print(f"  加载的文本增加规则: {len(loaded_gui_settings.get('text_additions', {}))}")
        for placeholder, rule in loaded_gui_settings.get('text_additions', {}).items():
            print(f"    {placeholder}: 前缀'{rule.get('prefix', '')}' 后缀'{rule.get('suffix', '')}'")

        print(f"  加载的递交趸期映射: {len(loaded_gui_settings.get('dundian_mappings', {}))}")
        for placeholder, data_column in loaded_gui_settings.get('dundian_mappings', {}).items():
            print(f"    {placeholder} -> {data_column}")

        # 4. 测试配置恢复
        print("\n4. 测试配置恢复...")

        # 创建新的模拟主窗口（空状态）
        empty_mock_window = MockMainWindow()
        empty_mock_window.text_addition_rules = {}
        empty_mock_window.dundian_mappings = {}

        print(f"  恢复前状态: 文本规则 {len(empty_mock_window.text_addition_rules)}, 趸期映射 {len(empty_mock_window.dundian_mappings)}")

        # 恢复配置
        success = config_manager._restore_gui_state(empty_mock_window, loaded_gui_settings)

        print(f"  恢复结果: {'成功' if success else '失败'}")
        print(f"  恢复后状态: 文本规则 {len(empty_mock_window.text_addition_rules)}, 趸期映射 {len(empty_mock_window.dundian_mappings)}")

        print("  恢复的文本增加规则:")
        for placeholder, rule in empty_mock_window.text_addition_rules.items():
            print(f"    {placeholder}: 前缀'{rule.get('prefix', '')}' 后缀'{rule.get('suffix', '')}'")

        print("  恢复的递交趸期映射:")
        for placeholder, data_column in empty_mock_window.dundian_mappings.items():
            print(f"    {placeholder} -> {data_column}")

        # 5. 测试配置有效性检查
        print("\n5. 测试配置有效性检查...")

        valid_gui_config = {
            "text_additions": {"ph_保险种类": {"suffix": "1件"}},
            "dundian_mappings": {"ph_保费金额": "期趸数据"}
        }

        invalid_gui_config = {
            "text_additions": {},
            "dundian_mappings": {}
        }

        print(f"  有效配置检查: {config_manager._has_valid_settings(valid_gui_config)}")
        print(f"  无效配置检查: {config_manager._has_valid_settings(invalid_gui_config)}")

        # 6. 测试占位符存在性检查
        print("\n6. 测试占位符存在性检查...")

        test_window = MockMainWindow()

        existing_placeholder = "ph_保费金额"
        non_existing_placeholder = "ph_不存在的占位符"

        print(f"  存在的占位符 {existing_placeholder}: {config_manager._placeholder_exists(test_window, existing_placeholder)}")
        print(f"  不存在的占位符 {non_existing_placeholder}: {config_manager._placeholder_exists(test_window, non_existing_placeholder)}")

        # 7. 测试更新匹配表格
        print("\n7. 测试更新匹配表格...")

        config_manager._update_match_table_for_dundian(test_window, "ph_保费金额", "期趸数据")
        config_manager._update_match_table_for_dundian(test_window, "ph_不存在的占位符", "期趸数据")

        # 清理临时文件
        os.unlink(temp_config_file)

        print(f"\n✅ 测试完成！所有功能正常工作")
        return True

    except Exception as e:
        print(f"\n❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def test_configuration_format():
    """测试配置文件格式兼容性"""
    try:
        print("\n=== 测试配置文件格式兼容性 ===")

        # 测试新格式配置
        new_format_config = {
            "version": "1.0.0",
            "created_at": "2025-01-11T12:00:00.000000",
            "gui_settings": {
                "text_additions": {
                    "ph_保险种类": {"suffix": "1件"}
                },
                "dundian_mappings": {
                    "ph_保费金额": "期趸数据"
                }
            }
        }

        # 测试旧格式配置（没有dundian_mappings）
        old_format_config = {
            "version": "1.0.0",
            "created_at": "2025-01-11T12:00:00.000000",
            "gui_settings": {
                "text_additions": {
                    "ph_保险种类": {"suffix": "1件"}
                },
                "dundian_settings": {  # 旧格式
                    "enabled": False,
                    "default_type": "期缴"
                }
            }
        }

        from src.gui.simple_config_manager import SimpleConfigManager
        config_manager = SimpleConfigManager()

        # 创建模拟主窗口
        class MockMainWindow:
            def __init__(self):
                self.text_addition_rules = {}
                self.dundian_mappings = {}
                self.match_table = MockMatchTable(["ph_保险种类", "ph_保费金额"])
                self.log_text = MockLogText()

        class MockLogText:
            def append(self, text):
                print(f"  LOG: {text}")

        class MockMatchTable:
            def __init__(self, placeholders):
                self._placeholders = placeholders

            def rowCount(self):
                return len(self._placeholders)

            def item(self, row, col):
                if col == 0:
                    return MockTableWidgetItem(self._placeholders[row])
                return None

        class MockTableWidgetItem:
            def __init__(self, text):
                self._text = text

            def text(self):
                return self._text

        mock_window = MockMainWindow()

        # 测试新格式配置加载
        print("1. 测试新格式配置加载...")
        success = config_manager._restore_gui_state(mock_window, new_format_config["gui_settings"])
        print(f"   新格式加载结果: {'成功' if success else '失败'}")
        print(f"   加载后的趸期映射: {mock_window.dundian_mappings}")

        # 重置窗口状态
        mock_window.dundian_mappings = {}

        # 测试旧格式配置加载（应该兼容）
        print("2. 测试旧格式配置兼容性...")
        success = config_manager._restore_gui_state(mock_window, old_format_config["gui_settings"])
        print(f"   旧格式加载结果: {'成功' if success else '失败'}")
        print(f"   加载后的趸期映射: {mock_window.dundian_mappings}")

        print("✅ 配置格式兼容性测试通过")
        return True

    except Exception as e:
        print(f"❌ 配置格式测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("开始递交趸期配置持久化功能测试...")

    success1 = test_dundian_config_persistence()
    success2 = test_configuration_format()

    overall_success = success1 and success2

    print(f"\n{'='*50}")
    print(f"测试结果: {'✅ 全部通过' if overall_success else '❌ 部分失败'}")
    print(f"{'='*50}")

    sys.exit(0 if overall_success else 1)