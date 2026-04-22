#!/usr/bin/env python3
"""
配置保存加载功能测试脚本
用于验证占位符配置的保存和加载功能
"""

import os
import sys
import json
import tempfile
from datetime import datetime

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from config_manager import ConfigManager
from exact_matcher import ExactMatcher
import pandas as pd


def test_config_manager_basic():
    """测试ConfigManager的基本功能"""
    print("=" * 50)
    print("测试ConfigManager基本功能")
    print("=" * 50)

    config_manager = ConfigManager()

    # 测试数据
    matching_rules = [
        {
            "placeholder": "ph_name",
            "column": "姓名",
            "text_addition": {
                "prefix": "",
                "suffix": "先生"
            }
        },
        {
            "placeholder": "ph_title",
            "column": "职位",
            "text_addition": {
                "prefix": "职位：",
                "suffix": ""
            }
        }
    ]

    template_info = {
        "file_name": "template.pptx",
        "placeholders": ["ph_name", "ph_title", "ph_company"]
    }

    data_info = {
        "file_name": "data.xlsx",
        "columns": ["姓名", "职位", "部门", "薪资"]
    }

    # 使用临时文件测试保存功能
    with tempfile.NamedTemporaryFile(mode='w', suffix='.pptcfg', delete=False) as f:
        temp_file = f.name

    try:
        # 测试保存配置
        print("1. 测试保存配置...")
        success = config_manager.save_placeholder_config(
            temp_file, matching_rules, template_info, data_info
        )

        if success:
            print("[OK] 配置保存成功")
            print(f"   保存文件: {temp_file}")
        else:
            print("[ERROR] 配置保存失败")
            return False

        # 验证保存的文件内容
        print("\n2. 验证保存的配置文件内容...")
        with open(temp_file, 'r', encoding='utf-8') as f:
            saved_config = json.load(f)

        print(f"   版本: {saved_config.get('version')}")
        print(f"   创建时间: {saved_config.get('created_at')}")
        print(f"   匹配规则数量: {len(saved_config.get('matching_rules', []))}")

        # 测试加载配置
        print("\n3. 测试加载配置...")
        load_result = config_manager.load_placeholder_config(temp_file)

        if load_result["success"]:
            print("✅ 配置加载成功")
            loaded_config = load_result["config_data"]
            print(f"   加载的匹配规则数量: {len(loaded_config.get('matching_rules', []))}")
        else:
            print(f"❌ 配置加载失败: {load_result.get('error')}")
            return False

        # 测试兼容性验证
        print("\n4. 测试兼容性验证...")
        current_template_placeholders = ["ph_name", "ph_title", "ph_department"]
        current_data_columns = ["姓名", "职位", "部门", "联系电话"]

        compatibility_result = config_manager.validate_placeholder_compatibility(
            loaded_config, current_template_placeholders, current_data_columns
        )

        print(f"   整体兼容率: {compatibility_result.get('overall_compatibility_rate', 0):.1%}")
        print(f"   兼容规则数: {compatibility_result.get('statistics', {}).get('compatible_rules_count', 0)}")
        print(f"   不兼容规则数: {compatibility_result.get('statistics', {}).get('incompatible_rules_count', 0)}")

        return True

    finally:
        # 清理临时文件
        if os.path.exists(temp_file):
            os.unlink(temp_file)


def test_exact_matcher_integration():
    """测试ExactMatcher的配置导出导入功能"""
    print("\n" + "=" * 50)
    print("测试ExactMatcher集成功能")
    print("=" * 50)

    # 创建测试数据
    test_data = pd.DataFrame({
        '姓名': ['张三', '李四', '王五'],
        '职位': ['经理', '工程师', '设计师'],
        '部门': ['技术部', '市场部', '设计部']
    })

    matcher = ExactMatcher()
    matcher.set_data(test_data)
    matcher.set_template_placeholders(['ph_name', 'ph_title', 'ph_dept'])

    # 设置一些匹配规则
    matcher.set_matching_rule('ph_name', '姓名')
    matcher.set_matching_rule('ph_title', '职位')

    # 设置文本增加规则
    matcher.set_text_addition_rule('ph_name', suffix=' 先生')
    matcher.set_text_addition_rule('ph_title', prefix='职位: ')

    try:
        # 测试导出配置
        print("1. 测试导出配置...")
        exported_rules = matcher.export_matching_config()

        if exported_rules:
            print(f"✅ 成功导出 {len(exported_rules)} 条规则")
            for rule in exported_rules:
                print(f"   {rule['placeholder']} -> {rule['column']}")
                if 'text_addition' in rule:
                    ta = rule['text_addition']
                    if ta.get('prefix') or ta.get('suffix'):
                        print(f"      文本增加: 前缀'{ta.get('prefix', '')}' 后缀'{ta.get('suffix', '')}'")
        else:
            print("❌ 导出配置失败")
            return False

        # 测试快速兼容性验证
        print("\n2. 测试快速兼容性验证...")
        compatibility_result = matcher.validate_config_compatibility_quick(exported_rules)

        if compatibility_result.get('compatible', False):
            print(f"✅ 配置兼容，兼容率: {compatibility_result.get('compatibility_rate', 0):.1%}")
        else:
            print(f"⚠️ 配置可能不兼容，兼容率: {compatibility_result.get('compatibility_rate', 0):.1%}")

        # 清除现有配置
        print("\n3. 测试导入配置...")
        matcher.clear_matching_rules()
        matcher.text_addition_rules = {}

        # 导入配置
        import_result = matcher.import_matching_config(exported_rules)

        if import_result.get('success', False):
            print("✅ 配置导入成功")
            print(f"   导入: {import_result.get('imported_count', 0)} 条")
            print(f"   跳过: {import_result.get('skipped_count', 0)} 条")
            print(f"   错误: {import_result.get('error_count', 0)} 条")

            if import_result.get('errors'):
                print("   错误详情:")
                for error in import_result['errors']:
                    print(f"     - {error}")
        else:
            print(f"❌ 配置导入失败: {import_result.get('error')}")
            return False

        # 验证导入后的配置
        print("\n4. 验证导入后的配置...")
        current_rules = matcher.get_matching_rules()
        print(f"   当前匹配规则数: {len(current_rules)}")
        for placeholder, column in current_rules.items():
            print(f"   {placeholder} -> {column}")

        current_text_rules = matcher.get_all_text_addition_rules()
        print(f"   文本增加规则数: {len(current_text_rules)}")
        for placeholder, rule in current_text_rules.items():
            print(f"   {placeholder}: 前缀'{rule.get('prefix', '')}' 后缀'{rule.get('suffix', '')}'")

        return True

    except Exception as e:
        print(f"❌ 测试过程中发生错误: {str(e)}")
        return False


def test_config_file_validation():
    """测试配置文件格式验证"""
    print("\n" + "=" * 50)
    print("测试配置文件格式验证")
    print("=" * 50)

    config_manager = ConfigManager()

    # 测试各种无效配置
    invalid_configs = [
        {},  # 空配置
        {"version": "1.0.0"},  # 缺少必需字段
        {"version": "2.0.0", "metadata": {}, "matching_rules": []},  # 不支持的版本
        {"version": "1.0.0", "metadata": {}, "matching_rules": "invalid"},  # 无效的matching_rules格式
        {"version": "1.0.0", "metadata": {}, "matching_rules": [{"placeholder": "ph_name"}]},  # 缺少column字段
        {"version": "1.0.0", "metadata": {}, "matching_rules": [{"column": "姓名"}]},  # 缺少placeholder字段
    ]

    for i, invalid_config in enumerate(invalid_configs, 1):
        print(f"{i}. 测试无效配置 {invalid_config}...")
        validation_result = config_manager._validate_placeholder_config_format(invalid_config)

        if not validation_result["valid"]:
            print(f"   ✅ 正确识别无效配置: {validation_result.get('error')}")
        else:
            print(f"   ❌ 未能识别无效配置")
            return False

    # 测试有效配置
    print(f"{len(invalid_configs) + 1}. 测试有效配置...")
    valid_config = {
        "version": "1.0.0",
        "metadata": {"template_info": {}, "data_info": {}, "rule_count": 1},
        "matching_rules": [
            {
                "placeholder": "ph_name",
                "column": "姓名",
                "text_addition": {"prefix": "", "suffix": "先生"}
            }
        ]
    }

    validation_result = config_manager._validate_placeholder_config_format(valid_config)

    if validation_result["valid"]:
        print("   ✅ 正确识别有效配置")
    else:
        print(f"   ❌ 错误地认为有效配置无效: {validation_result.get('error')}")
        return False

    return True


def main():
    """主测试函数"""
    print("开始配置保存加载功能测试")
    print(f"测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    all_tests_passed = True

    try:
        # 运行各项测试
        tests = [
            ("ConfigManager基本功能", test_config_manager_basic),
            ("ExactMatcher集成功能", test_exact_matcher_integration),
            ("配置文件格式验证", test_config_file_validation)
        ]

        for test_name, test_func in tests:
            print(f"\n开始执行: {test_name}")
            try:
                result = test_func()
                if result:
                    print(f"✅ {test_name} - 测试通过")
                else:
                    print(f"❌ {test_name} - 测试失败")
                    all_tests_passed = False
            except Exception as e:
                print(f"❌ {test_name} - 测试异常: {str(e)}")
                all_tests_passed = False

        # 总结测试结果
        print("\n" + "=" * 60)
        print("测试结果总结")
        print("=" * 60)

        if all_tests_passed:
            print("🎉 所有测试通过！配置保存加载功能工作正常。")
            return 0
        else:
            print("❌ 部分测试失败，请检查实现。")
            return 1

    except Exception as e:
        print(f"❌ 测试执行过程中发生未预期的错误: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())