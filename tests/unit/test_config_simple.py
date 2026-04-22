#!/usr/bin/env python3
"""
简单的配置保存加载功能测试
"""

import os
import sys
import json
import tempfile

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from config_manager import ConfigManager


def test_basic_functionality():
    """测试基本功能"""
    print("开始测试配置保存加载基本功能")

    config_manager = ConfigManager()

    # 测试数据
    matching_rules = [
        {
            "placeholder": "ph_name",
            "column": "姓名",
            "text_addition": {"prefix": "", "suffix": "先生"}
        },
        {
            "placeholder": "ph_title",
            "column": "职位",
            "text_addition": {"prefix": "职位：", "suffix": ""}
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

    # 使用临时文件
    temp_file = tempfile.mktemp(suffix='.pptcfg')

    try:
        # 测试保存
        print("1. 测试保存配置...")
        success = config_manager.save_placeholder_config(
            temp_file, matching_rules, template_info, data_info
        )

        if not success:
            print("[ERROR] 配置保存失败")
            return False

        print("[OK] 配置保存成功")
        print(f"   保存文件: {temp_file}")

        # 验证文件存在且可读
        if not os.path.exists(temp_file):
            print("[ERROR] 保存的文件不存在")
            return False

        # 测试加载
        print("2. 测试加载配置...")
        load_result = config_manager.load_placeholder_config(temp_file)

        if not load_result["success"]:
            print(f"[ERROR] 配置加载失败: {load_result.get('error')}")
            return False

        print("[OK] 配置加载成功")
        loaded_config = load_result["config_data"]
        print(f"   加载的匹配规则数量: {len(loaded_config.get('matching_rules', []))}")

        # 测试兼容性验证
        print("3. 测试兼容性验证...")
        current_placeholders = ["ph_name", "ph_title", "ph_department"]
        current_columns = ["姓名", "职位", "部门", "联系电话"]

        compat_result = config_manager.validate_placeholder_compatibility(
            loaded_config, current_placeholders, current_columns
        )

        compat_rate = compat_result.get('overall_compatibility_rate', 0)
        print(f"   整体兼容率: {compat_rate:.1%}")
        print(f"   兼容规则数: {compat_result.get('statistics', {}).get('compatible_rules_count', 0)}")
        print(f"   不兼容规则数: {compat_result.get('statistics', {}).get('incompatible_rules_count', 0)}")

        return True

    finally:
        # 清理临时文件
        if os.path.exists(temp_file):
            try:
                os.unlink(temp_file)
            except:
                pass


def test_format_validation():
    """测试格式验证"""
    print("\n开始测试配置文件格式验证")

    config_manager = ConfigManager()

    # 测试有效配置
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

    print("1. 测试有效配置格式...")
    result = config_manager._validate_placeholder_config_format(valid_config)

    if result["valid"]:
        print("[OK] 有效配置验证通过")
    else:
        print(f"[ERROR] 有效配置验证失败: {result.get('error')}")
        return False

    # 测试无效配置 - 缺少必需字段
    print("2. 测试无效配置格式（缺少字段）...")
    invalid_config = {"version": "1.0.0"}  # 缺少metadata和matching_rules
    result = config_manager._validate_placeholder_config_format(invalid_config)

    if not result["valid"]:
        print("[OK] 无效配置正确被拒绝")
        print(f"   拒绝原因: {result.get('error')}")
    else:
        print("[ERROR] 无效配置错误地被接受")
        return False

    # 测试无效配置 - 错误版本
    print("3. 测试无效配置格式（错误版本）...")
    invalid_config = {
        "version": "2.0.0",
        "metadata": {},
        "matching_rules": []
    }
    result = config_manager._validate_placeholder_config_format(invalid_config)

    if not result["valid"]:
        print("[OK] 错误版本正确被拒绝")
        print(f"   拒绝原因: {result.get('error')}")
    else:
        print("[ERROR] 错误版本错误地被接受")
        return False

    return True


def main():
    """主测试函数"""
    print("配置保存加载功能测试")
    print("=" * 40)

    all_passed = True

    try:
        # 运行测试
        if not test_basic_functionality():
            all_passed = False

        if not test_format_validation():
            all_passed = False

        # 总结
        print("\n" + "=" * 40)
        if all_passed:
            print("[SUCCESS] 所有测试通过!")
            print("配置保存加载功能正常工作")
        else:
            print("[FAILURE] 部分测试失败")
            print("请检查实现")

        return 0 if all_passed else 1

    except Exception as e:
        print(f"[ERROR] 测试过程中发生异常: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())