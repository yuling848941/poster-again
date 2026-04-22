#!/usr/bin/env python3
"""
完全模拟用户的操作流程
自动匹配 -> 加载配置 -> 批量生成
"""

import sys
import os
import json

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def simulate_user_workflow():
    """完全模拟用户的操作流程"""
    try:
        print("=== 完全模拟用户操作流程 ===")
        print("步骤：自动匹配 → 加载配置 → 批量生成")

        # 第一步：自动匹配
        print("\n=== 第一步：自动匹配 ===")

        # 加载Excel数据
        from src.data_reader import DataReader
        data_reader = DataReader()
        data_reader.load_excel("KA递交模板.xlsx")
        data = data_reader.get_data_preview(num_rows=5)  # 获取前5行数据用于测试

        print(f"数据列: {list(data.columns)}")
        print(f"数据行数: {len(data)}")
        print("第一行数据:")
        for col in data.columns:
            print(f"  {col}: {data[col].iloc[0]}")

        # 加载PPT模板
        from src.core.template_processor import TemplateProcessor
        from src.exact_matcher import ExactMatcher

        template_processor = TemplateProcessor()
        template_processor.load_template("10月递交贺报(3).pptx")

        exact_matcher = ExactMatcher()
        exact_matcher.set_data(data)

        # 获取模板占位符
        placeholders = template_processor.find_placeholders()
        print(f"模板占位符: {placeholders}")

        # 设置占位符并执行自动匹配
        exact_matcher.set_template_placeholders(placeholders)

        # 模拟自动匹配逻辑
        for placeholder in placeholders:
            placeholder_clean = placeholder.replace('ph_', '').lower()
            for column in data.columns:
                if column.lower() == placeholder_clean:
                    exact_matcher.set_matching_rule(placeholder, column)
                    print(f"自动匹配: {placeholder} -> {column}")
                    break

        matching_rules = exact_matcher.get_matching_rules()
        print(f"自动匹配结果: {matching_rules}")

        # 第二步：加载配置
        print("\n=== 第二步：加载配置 ===")

        # 加载配置文件
        config_file = "src/gui_config.pptcfg"
        with open(config_file, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        gui_settings = config_data.get('gui_settings', {})
        text_additions = gui_settings.get('text_additions', {})

        print("配置中的文本增加规则:")
        for placeholder, rule in text_additions.items():
            prefix = rule.get('prefix', '')
            suffix = rule.get('suffix', '')
            print(f"  {placeholder}: 前缀='{prefix}' 后缀='{suffix}'")

        # 模拟GUI界面状态
        class MockMainWindow:
            def __init__(self):
                self.text_addition_rules = {}
                self.exact_matcher = exact_matcher  # 使用实际的ExactMatcher
                self.worker_thread = MockWorkerThread()
                self.match_table = MockMatchTable(placeholders)
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
                    placeholder = self._placeholders[row]
                    if placeholder in matching_rules:
                        return MockTableWidgetItem(matching_rules[placeholder])
                    else:
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

        class MockWorkerThread:
            def __init__(self):
                self.text_addition_rules = {}
                self.ppt_generator = MockPPTGenerator()

            def set_text_addition_rules(self, rules):
                self.text_addition_rules = rules.copy()
                print(f"  工作线程同步: {len(rules)} 条规则")
                # 同步到PPT生成器
                for placeholder, rule in rules.items():
                    prefix = rule.get('prefix', '')
                    suffix = rule.get('suffix', '')
                    self.ppt_generator.exact_matcher.set_text_addition_rule(
                        placeholder, prefix, suffix
                    )

        class MockPPTGenerator:
            def __init__(self):
                self.exact_matcher = exact_matcher  # 使用相同的ExactMatcher实例

        mock_window = MockMainWindow()

        print("加载配置前:")
        print(f"  ExactMatcher规则: {len(exact_matcher.text_addition_rules)} 条")

        # 使用实际配置管理器加载配置
        from src.gui.simple_config_manager import SimpleConfigManager
        config_manager = SimpleConfigManager()

        load_success = config_manager.load_gui_config(mock_window, config_file)
        print(f"配置加载结果: {'成功' if load_success else '失败'}")

        print("加载配置后:")
        print(f"  ExactMatcher规则: {len(exact_matcher.text_addition_rules)} 条")
        print("  具体规则:")
        for placeholder, rule in exact_matcher.text_addition_rules.items():
            prefix = rule.get('prefix', '')
            suffix = rule.get('suffix', '')
            print(f"    {placeholder}: 前缀='{prefix}' 后缀='{suffix}'")

        # 第三步：批量生成
        print("\n=== 第三步：批量生成 ===")

        # 准备数据：应用文本增加规则
        print("应用文本增加规则到数据...")
        replacement_data = {}

        for placeholder in placeholders:
            if placeholder in matching_rules:
                data_column = matching_rules[placeholder]
                if data_column in data.columns:
                    # 获取原始数据
                    original_value = str(data[data_column].iloc[0])

                    # 检查是否有文本增加规则
                    if placeholder in exact_matcher.text_addition_rules:
                        rule = exact_matcher.text_addition_rules[placeholder]
                        prefix = rule.get('prefix', '')
                        suffix = rule.get('suffix', '')

                        # 应用前缀和后缀
                        modified_value = prefix + original_value + suffix

                        print(f"  应用规则: {placeholder} -> {original_value} -> {modified_value}")
                        replacement_data[placeholder] = modified_value
                    else:
                        replacement_data[placeholder] = original_value
                        print(f"  无规则: {placeholder} -> {original_value}")

        print(f"\n最终替换数据:")
        for placeholder, value in replacement_data.items():
            print(f"  {placeholder}: '{value}'")

        # 关键检查：ph_中支的情况
        if "ph_中支" in replacement_data:
            value = replacement_data["ph_中支"]
            if "中支" in value:
                print(f"\n[问题确认] ph_中支 包含自定义后缀: '{value}'")
                print("这说明配置被错误应用了！")
                return False
            else:
                print(f"\n[正常] ph_中支 不包含自定义后缀: '{value}'")

        # 执行PPT替换
        print("\n执行PPT替换...")
        template_processor.replace_placeholders(replacement_data)

        # 保存结果
        output_file = "output/simulated_workflow_test.pptx"
        template_processor.save_presentation(output_file)
        print(f"生成文件: {output_file}")

        # 关键分析
        print("\n=== 关键问题分析 ===")

        # 检查匹配情况
        print("占位符匹配分析:")
        matched_config = []
        for config_placeholder in text_additions.keys():
            if config_placeholder in placeholders:
                matched_config.append(config_placeholder)
                print(f"  [匹配] {config_placeholder}")
            else:
                # 检查相似占位符
                similar = [p for p in placeholders if any(keyword in p for keyword in config_placeholder.split('_')[1:])]
                if similar:
                    print(f"  [不匹配] {config_placeholder} -> 找到相似: {similar}")
                else:
                    print(f"  [不匹配] {config_placeholder} -> 无相似占位符")

        print(f"\n同步分析:")
        print(f"  期望同步的规则数: {len(matched_config)}")
        print(f"  实际同步的规则数: {len(exact_matcher.text_addition_rules)}")

        # 检查是否有错误同步
        error_synced_rules = []
        for synced_rule in exact_matcher.text_addition_rules.keys():
            if synced_rule not in matched_config:
                error_synced_rules.append(synced_rule)
                print(f"  [错误同步] {synced_rule} - 不应该被同步")

        if error_synced_rules:
            print(f"\n❌ 发现错误同步: {len(error_synced_rules)} 个规则被错误同步")
            return False
        else:
            print(f"\n✅ 同步正确: 只有匹配的规则被同步")
            return True

    except Exception as e:
        print(f"\n[ERROR] 模拟失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = simulate_user_workflow()
    print(f"\n{'模拟测试通过' if success else '模拟测试失败'}")
    sys.exit(0 if success else 1)