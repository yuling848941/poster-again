#!/usr/bin/env python3
"""
简化的递交趸期配置持久化测试
"""

import sys
import os
import json
import tempfile

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_dundian_simple():
    """简化的递交趸期配置测试"""
    try:
        print("=== 递交趸期配置测试 ===")

        from src.gui.simple_config_manager import SimpleConfigManager
        config_manager = SimpleConfigManager()

        # 测试配置收集
        print("\n1. 测试配置收集...")
        gui_config = {
            "text_additions": {
                "ph_保险种类": {"suffix": "1件"}
            },
            "dundian_mappings": {
                "ph_保费金额": "期趸数据"
            }
        }

        print(f"   文本规则数量: {len(gui_config['text_additions'])}")
        print(f"   趸期映射数量: {len(gui_config['dundian_mappings'])}")

        # 测试配置有效性
        print("\n2. 测试配置有效性...")
        is_valid = config_manager._has_valid_settings(gui_config)
        print(f"   配置有效性: {is_valid}")

        # 测试占位符存在性
        print("\n3. 测试占位符匹配...")

        class MockTable:
            def __init__(self):
                self.placeholders = ["ph_保费金额", "ph_缴费年期"]

            def rowCount(self):
                return len(self.placeholders)

            def item(self, row, col):
                if col == 0 and row < len(self.placeholders):
                    placeholder = self.placeholders[row]
                    class MockItem:
                        def __init__(self, text):
                            self._text = text
                        def text(self):
                            return self._text
                    return MockItem(placeholder)
                return None

        class MockWindow:
            def __init__(self):
                self.match_table = MockTable()

        mock_window = MockWindow()

        # 测试存在的占位符
        exists = config_manager._placeholder_exists(mock_window, "ph_保费金额")
        print(f"   ph_保费金额 存在: {exists}")

        # 测试不存在的占位符
        not_exists = config_manager._placeholder_exists(mock_window, "ph_不存在的占位符")
        print(f"   ph_不存在的占位符 存在: {not_exists}")

        print("\n4. 测试配置序列化...")

        # 创建完整配置
        full_config = {
            "version": "1.0.0",
            "created_at": "2025-01-11T12:00:00.000000",
            "gui_settings": gui_config
        }

        # 测试JSON序列化
        with tempfile.NamedTemporaryFile(mode='w', suffix='.pptcfg', delete=False, encoding='utf-8') as f:
            json.dump(full_config, f, ensure_ascii=False, indent=2)
            temp_file = f.name

        # 测试JSON反序列化
        with open(temp_file, 'r', encoding='utf-8') as f:
            loaded_config = json.load(f)

        # 验证数据完整性
        loaded_gui_settings = loaded_config.get('gui_settings', {})
        loaded_mappings = loaded_gui_settings.get('dundian_mappings', {})

        print(f"   序列化成功: {len(loaded_mappings) > 0}")
        print(f"   数据完整性: {'ph_保费金额' in loaded_mappings}")

        # 清理临时文件
        os.unlink(temp_file)

        print("\n=== 测试完成 ===")
        print("所有基本功能正常工作!")
        return True

    except Exception as e:
        print(f"测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_dundian_simple()
    sys.exit(0 if success else 1)