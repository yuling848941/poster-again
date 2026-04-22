#!/usr/bin/env python3
"""
配置保存加载功能演示脚本
展示完整的配置保存、加载和兼容性验证流程
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


def demo_config_workflow():
    """演示完整的配置工作流程"""
    print("=" * 60)
    print("PPT占位符配置保存加载功能演示")
    print("=" * 60)
    print(f"演示时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()

    # 步骤1: 模拟用户设置占位符匹配
    print("步骤1: 用户设置占位符匹配规则")
    print("-" * 40)

    # 创建测试数据（模拟Excel数据）
    test_data = pd.DataFrame({
        '姓名': ['张三', '李四', '王五', '赵六'],
        '职位': ['产品经理', '软件工程师', '设计师', '项目经理'],
        '部门': ['产品部', '技术部', '设计部', '产品部'],
        '邮箱': ['zhang@company.com', 'li@company.com', 'wang@company.com', 'zhao@company.com']
    })

    # 创建匹配器并设置数据
    matcher = ExactMatcher()
    matcher.set_data(test_data)

    # 设置模板占位符（模拟从PPT模板提取）
    template_placeholders = ['ph_name', 'ph_title', 'ph_dept', 'ph_email']
    matcher.set_template_placeholders(template_placeholders)

    print(f"   Excel数据列: {list(test_data.columns)}")
    print(f"   PPT模板占位符: {template_placeholders}")

    # 用户手动设置匹配规则
    matcher.set_matching_rule('ph_name', '姓名')
    matcher.set_matching_rule('ph_title', '职位')
    matcher.set_matching_rule('ph_dept', '部门')
    matcher.set_matching_rule('ph_email', '邮箱')

    # 设置文本增加规则
    matcher.set_text_addition_rule('ph_name', '', ' 先生')
    matcher.set_text_addition_rule('ph_title', '职位：', '')

    print("   用户设置的匹配规则:")
    for placeholder, column in matcher.get_matching_rules().items():
        text_rule = matcher.get_text_addition_rule(placeholder)
        prefix = text_rule.get('prefix', '')
        suffix = text_rule.get('suffix', '')
        text_info = f" (前缀:'{prefix}' 后缀:'{suffix}')" if prefix or suffix else ""
        print(f"     {placeholder} -> {column}{text_info}")

    # 步骤2: 保存配置
    print("\n步骤2: 保存配置到文件")
    print("-" * 40)

    config_manager = ConfigManager()

    # 导出配置
    matching_rules = matcher.export_matching_config()
    template_info = matcher.get_template_info("template.pptx")
    data_info = matcher.get_data_info("data.xlsx")

    # 保存到临时文件
    temp_config_file = tempfile.mktemp(suffix='.pptcfg')
    success = config_manager.save_placeholder_config(
        temp_config_file, matching_rules, template_info, data_info
    )

    if success:
        print(f"   [SUCCESS] 配置已保存到: {temp_config_file}")
        print(f"   包含 {len(matching_rules)} 条匹配规则")
    else:
        print("   [ERROR] 配置保存失败")
        return

    # 显示保存的配置文件内容
    print("\n   配置文件内容预览:")
    with open(temp_config_file, 'r', encoding='utf-8') as f:
        config_content = json.load(f)
    print(f"   版本: {config_content.get('version')}")
    print(f"   创建时间: {config_content.get('created_at')}")
    print(f"   模板: {config_content.get('metadata', {}).get('template_info', {}).get('file_name')}")
    print(f"   数据文件: {config_content.get('metadata', {}).get('data_info', {}).get('file_name')}")

    # 步骤3: 模拟更换模板和数据，加载配置
    print("\n步骤3: 更换模板和数据，加载配置")
    print("-" * 40)

    # 创建新的数据（模拟新的Excel文件）
    new_data = pd.DataFrame({
        '员工姓名': ['陈七', '周八', '吴九'],
        '工作职位': ['人事专员', '财务经理', '市场总监'],
        '所在部门': ['人事部', '财务部', '市场部']
    })

    # 设置新的模板占位符（部分匹配，部分不匹配）
    new_template_placeholders = ['ph_name', 'ph_title', 'ph_dept', 'ph_phone']  # ph_phone不匹配
    matcher.set_data(new_data)
    matcher.set_template_placeholders(new_template_placeholders)

    print(f"   新Excel数据列: {list(new_data.columns)}")
    print(f"   新PPT模板占位符: {new_template_placeholders}")

    # 加载配置
    print("\n   加载配置文件...")
    load_result = config_manager.load_placeholder_config(temp_config_file)

    if not load_result["success"]:
        print(f"   [ERROR] 配置加载失败: {load_result.get('error')}")
        return

    loaded_config = load_result["config_data"]
    print("   [SUCCESS] 配置文件加载成功")

    # 步骤4: 兼容性验证
    print("\n步骤4: 兼容性验证")
    print("-" * 40)

    compatibility_result = config_manager.validate_placeholder_compatibility(
        loaded_config, new_template_placeholders, list(new_data.columns)
    )

    overall_compat = compatibility_result.get('overall_compatibility_rate', 0)
    placeholder_compat = compatibility_result.get('placeholder_compatibility_rate', 0)
    data_compat = compatibility_result.get('data_compatibility_rate', 0)

    print(f"   整体兼容率: {overall_compat:.1%}")
    print(f"   占位符兼容率: {placeholder_compat:.1%}")
    print(f"   数据字段兼容率: {data_compat:.1%}")

    stats = compatibility_result.get('statistics', {})
    print(f"   统计信息:")
    print(f"     总规则数: {stats.get('total_rules', 0)}")
    print(f"     兼容规则数: {stats.get('compatible_rules_count', 0)}")
    print(f"     不兼容规则数: {stats.get('incompatible_rules_count', 0)}")

    # 显示兼容和不兼容的规则
    compatible_rules = compatibility_result.get('compatible_rules', [])
    incompatible_rules = compatibility_result.get('incompatible_rules', [])

    if compatible_rules:
        print("\n   兼容的规则（可以加载）:")
        for rule in compatible_rules:
            print(f"     {rule['placeholder']} -> {rule['column']}")

    if incompatible_rules:
        print("\n   不兼容的规则（将被跳过）:")
        for item in incompatible_rules:
            rule = item['rule']
            reason = item['reason']
            print(f"     {rule['placeholder']} -> {rule['column']} ({reason})")

    # 步骤5: 应用兼容的配置
    print("\n步骤5: 应用兼容的配置")
    print("-" * 40)

    if compatible_rules:
        print(f"   导入 {len(compatible_rules)} 条兼容规则...")
        import_result = matcher.import_matching_config(compatible_rules)

        if import_result.get('success'):
            print("   [SUCCESS] 配置导入成功")
            print(f"     导入: {import_result.get('imported_count', 0)} 条")
            print(f"     跳过: {import_result.get('skipped_count', 0)} 条")

            # 显示导入后的配置
            print("\n   导入后的配置:")
            current_rules = matcher.get_matching_rules()
            for placeholder, column in current_rules.items():
                text_rule = matcher.get_text_addition_rule(placeholder)
                prefix = text_rule.get('prefix', '')
                suffix = text_rule.get('suffix', '')
                text_info = f" (前缀:'{prefix}' 后缀:'{suffix}')" if prefix or suffix else ""
                print(f"     {placeholder} -> {column}{text_info}")

            # 测试数据替换
            print("\n   数据替换测试（第一行数据）:")
            test_row = matcher.get_all_data_for_row(0)
            for placeholder, value in test_row.items():
                print(f"     {placeholder}: {value}")
        else:
            print(f"   [ERROR] 配置导入失败: {import_result.get('error')}")
    else:
        print("   没有兼容的配置可以应用")

    # 步骤6: 清理和总结
    print("\n步骤6: 功能演示总结")
    print("-" * 40)

    print("   ✓ 配置保存功能 - 将用户设置的占位符匹配规则保存为.pptcfg文件")
    print("   ✓ 配置加载功能 - 从文件读取之前保存的配置")
    print("   ✓ 格式验证功能 - 验证配置文件的格式和版本兼容性")
    print("   ✓ 兼容性检查功能 - 检查配置与当前模板/数据的匹配程度")
    print("   ✓ 智能加载功能 - 只加载兼容的配置项，跳过不兼容项")
    print("   ✓ 错误处理功能 - 提供详细的错误信息和原因说明")

    print(f"\n   配置文件保存在: {temp_config_file}")
    print("   (程序结束后会自动清理临时文件)")

    # 清理临时文件
    try:
        os.unlink(temp_config_file)
        print("\n   临时文件已清理")
    except:
        pass

    print("\n" + "=" * 60)
    print("演示完成！配置保存加载功能已成功实现")
    print("=" * 60)


# 演示辅助函数已在上面直接调用，无需单独定义


if __name__ == "__main__":
    try:
        demo_config_workflow()
    except KeyboardInterrupt:
        print("\n\n演示被用户中断")
    except Exception as e:
        print(f"\n\n演示过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        input("\n按Enter键退出...")