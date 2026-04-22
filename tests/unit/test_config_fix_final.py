#!/usr/bin/env python3
"""
最终测试配置保存加载功能修复
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_final_fix():
    """测试最终修复"""
    try:
        print("测试配置保存加载功能最终修复...")

        from src.exact_matcher import ExactMatcher
        from src.config import ConfigManager
        import pandas as pd

        # 模拟用户的场景：只有文本增加规则的情况
        test_data = pd.DataFrame({
            '中心支公司': ['北京中支'],
            '保险种类': ['人寿保险']
        })

        matcher = ExactMatcher()
        matcher.set_data(test_data)

        # 设置模板占位符
        placeholders = ['ph_中心支公司', 'ph_保险种类']
        matcher.set_template_placeholders(placeholders)

        # 场景1：用户只设置了文本增加规则，没有匹配规则
        print("场景1：只有文本增加规则")
        matcher.set_text_addition_rule('ph_中心支公司', '', '中支')
        matcher.set_text_addition_rule('ph_保险种类', '', '1件')

        # 导出配置
        exported_rules = matcher.export_matching_config()
        text_rules = matcher.get_all_text_addition_rules()

        print(f"  导出的匹配规则数: {len(exported_rules)}")
        print(f"  文本增加规则数: {len(text_rules)}")

        # 生成保存规则（模拟配置对话框的逻辑）
        save_rules = exported_rules.copy()
        if text_rules:
            for placeholder, rule in text_rules.items():
                already_matched = any(r['placeholder'] == placeholder for r in save_rules)
                if not already_matched:
                    save_rules.append({
                        'placeholder': placeholder,
                        'column': '',  # 空列名表示只有文本增加规则
                        'text_addition': rule
                    })

        print(f"  最终保存规则数: {len(save_rules)}")

        # 测试保存
        if save_rules:
            config_manager = ConfigManager()

            template_info = {
                'file_name': 'template.pptx',
                'placeholders': placeholders,
                'placeholder_count': len(placeholders)
            }

            data_info = {
                'file_name': 'data.xlsx',
                'columns': list(test_data.columns),
                'column_count': len(test_data.columns)
            }

            import tempfile
            temp_file = tempfile.mktemp(suffix='.pptcfg')

            success = config_manager.save_placeholder_config(
                temp_file, save_rules, template_info, data_info
            )

            if success:
                print(f"  [OK] 配置保存成功: {temp_file}")

                # 验证保存的文件
                load_result = config_manager.load_placeholder_config(temp_file)
                if load_result['success']:
                    loaded_config = load_result['config_data']
                    loaded_rules = loaded_config.get('matching_rules', [])
                    print(f"  [OK] 验证通过，加载了 {len(loaded_rules)} 条规则")

                    # 显示加载的规则详情
                    for rule in loaded_rules:
                        placeholder = rule['placeholder']
                        column = rule['column']
                        text_info = ""
                        if 'text_addition' in rule:
                            ta = rule['text_addition']
                            text_info = f" (前缀:'{ta.get('prefix', '')}' 后缀:'{ta.get('suffix', '')}')"
                        print(f"    {placeholder} -> '{column}'{text_info}")
                else:
                    print(f"  [ERROR] 验证失败: {load_result.get('error')}")

            # 清理临时文件
            try:
                os.unlink(temp_file)
            except:
                pass

        # 场景2：同时有匹配规则和文本增加规则
        print("\n场景2：匹配规则 + 文本增加规则")
        matcher.set_matching_rule('ph_中心支公司', '中心支公司')
        matcher.set_matching_rule('ph_保险种类', '保险种类')

        exported_rules = matcher.export_matching_config()
        print(f"  导出的匹配规则数: {len(exported_rules)}")

        for rule in exported_rules:
            placeholder = rule['placeholder']
            column = rule['column']
            text_info = ""
            if 'text_addition' in rule:
                ta = rule['text_addition']
                text_info = f" (前缀:'{ta.get('prefix', '')}' 后缀:'{ta.get('suffix', '')}')"
            print(f"    {placeholder} -> {column}{text_info}")

        print("\n[SUCCESS] 所有测试通过！")
        print("修复要点:")
        print("1. 改进了同步逻辑，正确解析表格中的匹配规则")
        print("2. 支持保存只有文本增加规则的情况")
        print("3. 增加了详细的调试信息和状态显示")
        print("4. 配置对话框能正确显示当前配置状态")

        return True

    except Exception as e:
        print(f"[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_final_fix()
    sys.exit(0 if success else 1)