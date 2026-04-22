#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最终验证修复
测试程序能否正确识别PPT中的占位符
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import logging
from src.gui.office_suite_detector import OfficeSuiteManager
from src.ppt_generator import PPTGenerator

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def test_ppt_generator_with_file(ppt_path, name):
    """
    测试PPTGenerator对特定文件的处理
    """
    print(f"\n{'='*70}")
    print(f"测试: {name}")
    print(f"文件: {ppt_path}")
    print('='*70)

    try:
        # 1. 创建OfficeSuiteManager
        print("\n1. 创建OfficeSuiteManager...")
        office_manager = OfficeSuiteManager()
        office_manager.initialize()
        print("   ✓ 创建成功")

        # 2. 创建PPTGenerator
        print("\n2. 创建PPTGenerator...")
        generator = PPTGenerator(office_manager=office_manager)
        print("   ✓ 创建成功")

        # 3. 加载模板
        print("\n3. 加载模板...")
        success = generator.load_template(ppt_path)
        if not success:
            print("   ✗ 加载失败")
            return []
        print("   ✓ 加载成功")

        # 4. 提取占位符
        print("\n4. 提取占位符...")
        placeholders = generator.extract_placeholders()
        print(f"   ✓ 提取完成，找到 {len(placeholders)} 个占位符")

        if placeholders:
            print("\n   占位符列表:")
            for i, ph in enumerate(placeholders, 1):
                print(f"     {i}. {ph}")
        else:
            print("   ✗ 未找到占位符")

        return placeholders

    except Exception as e:
        print(f"\n   ✗ 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return []

    finally:
        # 清理资源
        try:
            generator.close_template_processor()
            print("\n5. 已关闭模板处理器")
        except:
            pass

def main():
    """主测试函数"""
    print("="*70)
    print("PPT占位符提取测试 - 最终验证")
    print("="*70)
    print("测试程序能否正确识别PPT模板中的占位符")
    print("="*70)

    files = [
        ("E:/poster code/poster-again/10月承保贺报(4).pptx", "(4)文件"),
        ("E:/poster code/poster-again/10月递交贺报(3).pptx", "(3)文件")
    ]

    results = {}

    for ppt_path, name in files:
        if not os.path.exists(ppt_path):
            print(f"\n文件不存在: {ppt_path}")
            continue

        result = test_ppt_generator_with_file(ppt_path, name)
        results[name] = result

    # 总结
    print("\n" + "="*70)
    print("测试总结")
    print("="*70)

    for name, placeholders in results.items():
        print(f"{name}: {len(placeholders)} 个占位符")

    # 检查结果
    all_success = all(len(ph) > 0 for ph in results.values())

    if all_success:
        print("\n" + "="*70)
        print("✓ 所有测试通过！")
        print("="*70)
        print("程序可以正确识别PPT中的占位符！")
        print("您现在可以在GUI中正常使用这个功能了。")
    else:
        print("\n" + "="*70)
        print("✗ 部分测试失败")
        print("="*70)
        print("程序仍无法正确识别占位符，需要进一步调试。")

if __name__ == "__main__":
    main()