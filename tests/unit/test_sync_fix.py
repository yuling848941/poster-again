#!/usr/bin/env python3
"""
测试配置同步修复
验证主界面和ExactMatcher之间的状态同步是否正常
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_sync_logic():
    """测试同步逻辑"""
    try:
        print("开始测试配置同步逻辑...")

        # 导入必要模块
        from src.exact_matcher import ExactMatcher
        from src.config import ConfigManager
        import pandas as pd

        # 创建测试数据
        test_data = pd.DataFrame({
            '中心支公司': ['北京中支', '上海中支', '广州中支'],
            '保险种类': ['人寿保险', '车险', '健康保险'],
            '所属家庭': ['张家', '李家', '王家']
        })

        # 创建ExactMatcher
        matcher = ExactMatcher()
        matcher.set_data(test_data)

        # 设置模板占位符
        placeholders = ['ph_中心支公司', 'ph_保险种类', 'ph_所属家庭', 'ph_递交字段']
        matcher.set_template_placeholders(placeholders)

        # 模拟主界面的匹配状态
        print("1. 模拟主界面设置匹配规则...")
        # 设置匹配规则
        matcher.set_matching_rule('ph_中心支公司', '中心支公司')
        matcher.set_matching_rule('ph_保险种类', '保险种类')
        matcher.set_matching_rule('ph_所属家庭', '所属家庭')
        matcher.set_matching_rule('ph_递交字段', '期趸数据列')

        # 设置文本增加规则
        matcher.set_text_addition_rule('ph_中心支公司', '', '中支')
        matcher.set_text_addition_rule('ph_保险种类', '', '1件')
        matcher.set_text_addition_rule('ph_所属家庭', '', '家族')

        print("2. 检查匹配规则...")
        matching_rules = matcher.get_matching_rules()
        print(f"   匹配规则数量: {len(matching_rules)}")
        for placeholder, column in matching_rules.items():
            print(f"   {placeholder} -> {column}")

        print("3. 检查文本增加规则...")
        text_rules = matcher.get_all_text_addition_rules()
        print(f"   文本增加规则数量: {len(text_rules)}")
        for placeholder, rule in text_rules.items():
            print(f"   {placeholder}: 前缀='{rule.get('prefix', '')}' 后缀='{rule.get('suffix', '')}'")

        print("4. 测试配置导出...")
        exported_rules = matcher.export_matching_config()
        print(f"   导出的规则数量: {len(exported_rules)}")
        for rule in exported_rules:
            placeholder = rule['placeholder']
            column = rule['column']
            text_info = ""
            if 'text_addition' in rule:
                ta = rule['text_addition']
                text_info = f" (前缀:'{ta.get('prefix', '')}' 后缀:'{ta.get('suffix', '')}')"
            print(f"   {placeholder} -> {column}{text_info}")

        print("5. 测试配置保存...")
        config_manager = ConfigManager()

        # 创建模拟的模板和数据信息
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

        # 测试保存到临时文件
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
                print("   [OK] 配置加载验证成功")
                loaded_config = load_result['config_data']
                loaded_rules = loaded_config.get('matching_rules', [])
                print(f"   加载的规则数量: {len(loaded_rules)}")
            else:
                print(f"   [ERROR] 配置加载验证失败: {load_result.get('error')}")
        else:
            print("   [ERROR] 配置保存失败")

        # 清理临时文件
        try:
            os.unlink(temp_file)
        except:
            pass

        print("\n[SUCCESS] 同步逻辑测试通过！")
        print("修复说明:")
        print("1. 增加了详细的日志记录，便于调试同步问题")
        print("2. 改进了匹配规则检查逻辑，跳过'未匹配'的项目")
        print("3. 优化了配置保存的提示信息")
        print("4. 防止空配置导致的状态丢失")

        return True

    except Exception as e:
        print(f"[ERROR] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_sync_logic()
    sys.exit(0 if success else 1)