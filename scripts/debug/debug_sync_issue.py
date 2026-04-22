#!/usr/bin/env python3
"""
调试配置同步问题
模拟用户场景：自动匹配 + 添加后缀 + 保存配置
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def debug_sync_issue():
    """调试同步问题"""
    try:
        print("调试配置同步问题...")

        from src.exact_matcher import ExactMatcher
        import pandas as pd

        # 模拟用户的Excel数据
        test_data = pd.DataFrame({
            '中心支公司': ['北京中支', '上海中支'],
            '保险种类': ['人寿保险', '车险'],
            '所属家庭': ['张家', '李家'],
            '期趸数据列': ['期缴', '趸交']
        })

        # 创建ExactMatcher
        matcher = ExactMatcher()
        matcher.set_data(test_data)

        # 模拟用户自动匹配后的占位符
        placeholders = ['ph_中心支公司', 'ph_保险种类', 'ph_所属家庭', 'ph_递交字段']
        matcher.set_template_placeholders(placeholders)

        print("1. 设置数据模板:")
        print(f"   数据列: {list(test_data.columns)}")
        print(f"   占位符: {placeholders}")

        # 模拟自动匹配结果
        print("\n2. 模拟自动匹配...")
        matcher.set_matching_rule('ph_中心支公司', '中心支公司')
        matcher.set_matching_rule('ph_保险种类', '保险种类')
        matcher.set_matching_rule('ph_所属家庭', '所属家庭')
        matcher.set_matching_rule('ph_递交字段', '期趸数据列')

        print("   自动匹配完成")

        # 模拟用户添加后缀
        print("\n3. 模拟用户添加后缀...")
        matcher.set_text_addition_rule('ph_中心支公司', '', '中支')
        matcher.set_text_addition_rule('ph_保险种类', '', '1件')
        matcher.set_text_addition_rule('ph_所属家庭', '', '家族')

        print("   后缀设置完成")

        # 检查当前状态
        print("\n4. 检查ExactMatcher状态:")
        matching_rules = matcher.get_matching_rules()
        text_rules = matcher.get_all_text_addition_rules()

        print(f"   匹配规则数量: {len(matching_rules)}")
        for placeholder, column in matching_rules.items():
            print(f"     {placeholder} -> {column}")

        print(f"   文本增加规则数量: {len(text_rules)}")
        for placeholder, rule in text_rules.items():
            print(f"     {placeholder}: 前缀='{rule.get('prefix', '')}' 后缀='{rule.get('suffix', '')}'")

        # 测试配置导出
        print("\n5. 测试配置导出...")
        exported_rules = matcher.export_matching_config()

        print(f"   导出规则数量: {len(exported_rules)}")
        for rule in exported_rules:
            placeholder = rule['placeholder']
            column = rule['column']
            text_info = ""
            if 'text_addition' in rule:
                ta = rule['text_addition']
                text_info = f" (前缀:'{ta.get('prefix', '')}' 后缀:'{ta.get('suffix', '')}')"
            print(f"     {placeholder} -> {column}{text_info}")

        # 测试配置保存
        if exported_rules:
            print("\n6. 测试配置保存...")
            from src.config import ConfigManager
            config_manager = ConfigManager()

            # 创建模板和数据信息
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

            # 保存到临时文件
            import tempfile
            temp_file = tempfile.mktemp(suffix='.pptcfg')

            success = config_manager.save_placeholder_config(
                temp_file, exported_rules, template_info, data_info
            )

            if success:
                print(f"   [OK] 配置保存成功: {temp_file}")

                # 验证保存的文件
                load_result = config_manager.load_placeholder_config(temp_file)
                if load_result['success']:
                    loaded_config = load_result['config_data']
                    loaded_rules = loaded_config.get('matching_rules', [])
                    print(f"   验证通过，加载了 {len(loaded_rules)} 条规则")
                else:
                    print(f"   验证失败: {load_result.get('error')}")
            else:
                print("   [ERROR] 配置保存失败")

            # 清理临时文件
            try:
                os.unlink(temp_file)
            except:
                pass
        else:
            print("   [ERROR] 没有导出的规则，无法测试保存")

        print("\n[SUCCESS] 调试完成！")
        print("如果用户界面仍有问题，可能的原因:")
        print("1. 主界面的表格数据格式与同步逻辑不匹配")
        print("2. _sync_exact_matcher方法没有被正确调用")
        print("3. 数据路径或模板路径有问题")

        return True

    except Exception as e:
        print(f"[ERROR] 调试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = debug_sync_issue()
    sys.exit(0 if success else 1)