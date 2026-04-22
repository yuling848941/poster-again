#!/usr/bin/env python3
"""
使用修复后的逻辑生成测试文件
验证修复是否真正解决了问题
"""

import sys
import os
import json

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_with_fixed_logic():
    """使用修复后的逻辑生成测试文件"""
    try:
        print("=== 使用修复后的逻辑生成测试文件 ===")

        # 1. 加载配置和数据
        print("\n1. 加载实际配置和数据...")

        # 加载配置
        config_file = "src/gui_config.pptcfg"
        with open(config_file, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        gui_settings = config_data.get('gui_settings', {})
        text_additions = gui_settings.get('text_additions', {})

        print("   配置中的文本增加规则:")
        for placeholder, rule in text_additions.items():
            suffix = rule.get('suffix', '')
            print(f"     {placeholder}: 后缀='{suffix}'")

        # 加载数据
        from src.data_reader import DataReader
        data_reader = DataReader()
        data_reader.load_excel("KA递交模板.xlsx")
        data = data_reader.get_data_preview(num_rows=1)  # 只取第一行数据

        print(f"   数据列: {list(data.columns)}")

        # 2. 获取模板占位符并执行匹配
        print("\n2. 执行自动匹配...")
        from src.core.template_processor import TemplateProcessor
        from src.exact_matcher import ExactMatcher

        template_processor = TemplateProcessor()
        template_processor.load_template("10月递交贺报(3).pptx")
        exact_matcher = ExactMatcher()
        exact_matcher.set_data(data)

        placeholders = template_processor.find_placeholders()
        exact_matcher.set_template_placeholders(placeholders)

        # 自动匹配
        for placeholder in placeholders:
            placeholder_clean = placeholder.replace('ph_', '').lower()
            for column in data.columns:
                if column.lower() == placeholder_clean:
                    exact_matcher.set_matching_rule(placeholder, column)
                    print(f"   匹配: {placeholder} -> {column}")
                    break

        matching_rules = exact_matcher.get_matching_rules()
        print(f"   匹配结果: {len(matching_rules)} 条")

        # 3. 使用修复后的配置加载逻辑
        print("\n3. 使用修复后的配置加载逻辑...")
        from src.gui.simple_config_manager import SimpleConfigManager

        config_manager = SimpleConfigManager()

        # 模拟配置加载过程
        matched_text_additions = config_manager._update_match_table_display(
            type('MockWindow', (), {
                'match_table': type('MockTable', (), {
                    'rowCount': lambda: len(placeholders),
                    'item': lambda self, row, col: type('Item', (), {'text':
                        lambda: placeholders[row] if col == 0 else matching_rules.get(placeholders[row], "")
                    })() if col in [0, 1] else None
                })(),
            })(),
            text_additions
        )

        print(f"   匹配的规则数量: {len(matched_text_additions)}")
        print("   匹配的规则:")
        for placeholder, rule in matched_text_additions.items():
            print(f"     {placeholder}: 后缀'{rule.get('suffix', '')}'")

        # 4. 应用匹配的文本规则到ExactMatcher
        print("\n4. 应用匹配的文本规则...")
        for placeholder, rule in matched_text_additions.items():
            prefix = rule.get('prefix', '')
            suffix = rule.get('suffix', '')
            exact_matcher.set_text_addition_rule(placeholder, prefix, suffix)
            print(f"   应用: {placeholder} -> 后缀'{suffix}'")

        print(f"   ExactMatcher中的规则: {len(exact_matcher.text_addition_rules)} 条")

        # 5. 生成PPT
        print("\n5. 生成测试PPT...")

        # 准备替换数据
        replacement_data = {}
        for placeholder in placeholders:
            if placeholder in matching_rules:
                data_column = matching_rules[placeholder]
                if data_column in data.columns:
                    value = str(data[data_column].iloc[0])

                    # 应用文本增加规则
                    if placeholder in exact_matcher.text_addition_rules:
                        rule = exact_matcher.text_addition_rules[placeholder]
                        prefix = rule.get('prefix', '')
                        suffix = rule.get('suffix', '')
                        value = prefix + value + suffix

                    replacement_data[placeholder] = value

        print(f"   替换数据:")
        for placeholder, value in replacement_data.items():
            print(f"     {placeholder}: '{value}'")

        # 关键检查：ph_中支的值
        if "ph_中支" in replacement_data:
            value = replacement_data["ph_中支"]
            if "中支" in value:
                print(f"   [问题] ph_中支 包含自定义后缀: '{value}'")
                return False
            else:
                print(f"   [正确] ph_中支 不包含自定义后缀: '{value}'")

        # 执行替换
        template_processor.replace_placeholders(replacement_data)

        # 保存结果
        output_file = "output/fixed_logic_test.pptx"
        template_processor.save_presentation(output_file)
        print(f"   生成测试文件: {output_file}")

        # 6. 最终验证
        print("\n6. 最终验证...")
        expected_matches = ["ph_保险种类", "ph_所属家庭"]  # 应该匹配的
        expected_no_matches = ["ph_中心支公司"]  # 不应该匹配的

        applied_rules = list(exact_matcher.text_addition_rules.keys())

        success = True
        for expected in expected_matches:
            if expected in applied_rules:
                print(f"   [OK] {expected} 正确被应用")
            else:
                print(f"   [ERROR] {expected} 应该被应用但没有")
                success = False

        for expected in expected_no_matches:
            if expected not in applied_rules:
                print(f"   [OK] {expected} 正确没有被应用")
            else:
                print(f"   [ERROR] {expected} 不应该被应用但被应用了")
                success = False

        print(f"\n{'修复验证成功' if success else '修复验证失败'}")
        print(f"测试文件: {output_file}")
        print("请检查生成的文件，确认 ph_中支 是否没有自定义后缀")

        return success

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_with_fixed_logic()
    sys.exit(0 if success else 1)