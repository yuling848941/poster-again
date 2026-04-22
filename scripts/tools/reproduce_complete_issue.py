#!/usr/bin/env python3
"""
完整重现用户的配置同步问题
使用实际的模板、数据和配置文件
"""

import sys
import os
import json

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def reproduce_complete_issue():
    """完整重现配置同步问题"""
    try:
        print("=== 重现配置同步问题 ===")

        # 1. 加载配置文件
        print("\n1. 加载配置文件...")
        config_file = "src/gui_config.pptcfg"
        with open(config_file, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        gui_settings = config_data.get('gui_settings', {})
        text_additions = gui_settings.get('text_additions', {})

        print("   配置中的文本增加规则:")
        for placeholder, rule in text_additions.items():
            prefix = rule.get('prefix', '')
            suffix = rule.get('suffix', '')
            print(f"     {placeholder}: 前缀='{prefix}' 后缀='{suffix}'")

        # 2. 加载Excel数据
        print("\n2. 加载Excel数据...")
        from src.data_reader import DataReader
        data_reader = DataReader()
        data_reader.load_excel("KA递交模板.xlsx")
        data = data_reader.get_data_preview(num_rows=10)  # 获取前10行数据

        print(f"   数据列: {list(data.columns)}")
        print(f"   数据行数: {len(data)}")
        print("   前3行数据:")
        print(data.head(3).to_string())

        # 3. 加载模板并获取占位符
        print("\n3. 加载模板并获取占位符...")
        from src.core.template_processor import TemplateProcessor
        from src.exact_matcher import ExactMatcher

        template_processor = TemplateProcessor()
        template_processor.load_template("10月递交贺报(3).pptx")

        exact_matcher = ExactMatcher()
        exact_matcher.set_data(data)

        placeholders = template_processor.find_placeholders()
        print(f"   模板占位符: {placeholders}")

        # 4. 执行自动匹配
        print("\n4. 执行自动匹配...")
        exact_matcher.set_template_placeholders(placeholders)

        # 使用自动匹配逻辑
        for placeholder in placeholders:
            placeholder_clean = placeholder.replace('ph_', '').lower()

            # 寻找匹配的列
            for column in data.columns:
                if column.lower() == placeholder_clean:
                    exact_matcher.set_matching_rule(placeholder, column)
                    print(f"   匹配: {placeholder} -> {column}")
                    break

        matching_rules = exact_matcher.get_matching_rules()
        print(f"   匹配结果: {matching_rules}")

        # 5. 模拟配置加载过程 - 这是最关键的部分
        print("\n5. 模拟配置加载过程...")
        from src.gui.simple_config_manager import SimpleConfigManager

        config_manager = SimpleConfigManager()

        # 模拟主窗口状态
        class MockMainWindow:
            def __init__(self):
                self.text_addition_rules = {}
                self.exact_matcher = exact_matcher  # 使用实际的ExactMatcher
                self.worker_thread = MockWorkerThread()
                self.match_table = MockMatchTable(placeholders)
                self.log_text = MockLogText()

        class MockLogText:
            def append(self, text):
                print(f"   LOG: {text}")

        class MockMatchTable:
            def __init__(self, placeholders):
                self._placeholders = placeholders

            def rowCount(self):
                return len(self._placeholders)

            def item(self, row, col):
                if col == 0:  # 占位符列
                    return MockTableWidgetItem(self._placeholders[row])
                elif col == 1:  # 数据列 - 模拟显示匹配的列名
                    placeholder = self._placeholders[row]
                    # 获取匹配的列名
                    if placeholder in matching_rules:
                        return MockTableWidgetItem(matching_rules[placeholder])
                    else:
                        return MockTableWidgetItem("")
                return None

            def setItem(self, row, col, item):
                if col == 1:
                    print(f"   表格更新: {self._placeholders[row]} -> {item.text()}")

        class MockTableWidgetItem:
            def __init__(self, text):
                self._text = text

            def text(self):
                return self._text

        class MockWorkerThread:
            def __init__(self):
                self.text_addition_rules = {}
                self.ppt_generator = MockPPTGenerator()

            def set_text_addition_rules(self, rules):
                self.text_addition_rules = rules.copy()
                print(f"   工作线程同步: {len(rules)} 条规则")
                # 同步到PPT生成器
                for placeholder, rule in rules.items():
                    prefix = rule.get('prefix', '')
                    suffix = rule.get('suffix', '')
                    self.ppt_generator.exact_matcher.set_text_addition_rule(
                        placeholder, prefix, suffix
                    )

        class MockPPTGenerator:
            def __init__(self):
                self.exact_matcher = ExactMatcher()  # 新的ExactMatcher实例

        mock_window = MockMainWindow()

        print("   加载配置前:")
        print(f"     ExactMatcher规则: {len(exact_matcher.text_addition_rules)} 条")

        # 执行配置加载
        print("\n   执行配置加载...")
        load_success = config_manager.load_gui_config(mock_window, config_file)
        print(f"   加载结果: {'成功' if load_success else '失败'}")

        print("   加载配置后:")
        print(f"     ExactMatcher规则: {len(exact_matcher.text_addition_rules)} 条")
        for placeholder, rule in exact_matcher.text_addition_rules.items():
            prefix = rule.get('prefix', '')
            suffix = rule.get('suffix', '')
            print(f"       {placeholder}: 前缀='{prefix}' 后缀='{suffix}'")

        # 6. 执行实际的PPT生成
        print("\n6. 执行PPT生成...")

        # 准备数据：包含文本增加规则的数据
        data_with_text = data.copy()

        # 应用文本增加规则
        for placeholder, rule in exact_matcher.text_addition_rules.items():
            # 检查是否有匹配的数据列
            if placeholder in matching_rules:
                data_column = matching_rules[placeholder]
                prefix = rule.get('prefix', '')
                suffix = rule.get('suffix', '')

                print(f"   应用文本规则: {placeholder} -> {data_column} (前缀:'{prefix}' 后缀:'{suffix}')")

                # 应用前缀和后缀到数据列
                if data_column in data_with_text.columns:
                    if prefix:
                        data_with_text[data_column] = prefix + data_with_text[data_column].astype(str)
                    if suffix:
                        data_with_text[data_column] = data_with_text[data_column].astype(str) + suffix

        # 显示应用文本规则后的数据
        print("   应用文本规则后的数据:")
        for column in data_with_text.columns:
            if column in matching_rules.values():
                print(f"     {column}: {data_with_text[column].iloc[0]}")

        # 构建替换数据
        replacement_data = {}
        for placeholder in placeholders:
            if placeholder in matching_rules:
                data_column = matching_rules[placeholder]
                if data_column in data_with_text.columns:
                    replacement_data[placeholder] = str(data_with_text[data_column].iloc[0])

        print(f"   替换数据: {replacement_data}")

        # 执行替换
        template_processor.replace_placeholders(replacement_data)

        # 保存结果
        output_file = "output/debug_test.pptx"
        template_processor.save_presentation(output_file)
        print(f"   生成文件: {output_file}")

        # 7. 关键分析
        print("\n7. 关键问题分析...")

        # 检查 ph_中支 是否有自定义文本
        if "ph_中支" in replacement_data:
            text = replacement_data["ph_中支"]
            if "中支" in text:
                print(f"   [问题确认] ph_中支 包含自定义文本: '{text}'")
                print(f"   [原因分析] 这说明 'ph_中心支公司' 的配置被错误应用到了 'ph_中支'")
                return False
            else:
                print(f"   [正常] ph_中支 不包含自定义文本: '{text}'")
        else:
            print(f"   [注意] ph_中支 不在替换数据中")

        return True

    except Exception as e:
        print(f"\n[ERROR] 重现失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = reproduce_complete_issue()
    sys.exit(0 if success else 1)