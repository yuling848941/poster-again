#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模拟程序完整流程
测试实际加载PPT和提取占位符的过程
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import logging
from src.core.template_processor import PPTXProcessor, TemplateProcessor

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def test_pptx_processor_actual(ppt_path):
    """
    实际测试PPTXProcessor加载和提取占位符
    """
    print(f"\n{'='*70}")
    print(f"测试: {os.path.basename(ppt_path)}")
    print('='*70)

    processor = PPTXProcessor()

    # 1. 加载模板
    print("\n1. 加载模板...")
    try:
        success = processor.load_template(ppt_path)
        if not success:
            print("  ✗ 加载失败")
            return []
        print("  ✓ 加载成功")
    except Exception as e:
        print(f"  ✗ 加载异常: {e}")
        return []

    # 2. 提取占位符
    print("\n2. 提取占位符...")
    try:
        placeholders = processor.find_placeholders()
        print(f"  ✓ 提取完成，找到 {len(placeholders)} 个占位符")

        if placeholders:
            print("\n  占位符列表:")
            for i, ph in enumerate(placeholders, 1):
                print(f"    {i}. {ph}")
        else:
            print("  ✗ 未找到占位符")

        return placeholders

    except Exception as e:
        print(f"  ✗ 提取异常: {e}")
        import traceback
        traceback.print_exc()
        return []

def test_template_processor_actual(ppt_path):
    """
    实际测试TemplateProcessor加载和提取占位符
    """
    print(f"\n{'='*70}")
    print(f"测试: {os.path.basename(ppt_path)}")
    print('='*70)

    processor = TemplateProcessor()

    # 1. 连接到PowerPoint
    print("\n1. 连接到PowerPoint...")
    try:
        success = processor.connect_to_powerpoint()
        if not success:
            print("  ✗ 连接失败")
            return []
        print("  ✓ 连接成功")
    except Exception as e:
        print(f"  ✗ 连接异常: {e}")
        return []

    # 2. 加载模板
    print("\n2. 加载模板...")
    try:
        success = processor.load_template(ppt_path)
        if not success:
            print("  ✗ 加载失败")
            return []
        print("  ✓ 加载成功")
    except Exception as e:
        print(f"  ✗ 加载异常: {e}")
        return []

    # 3. 提取占位符
    print("\n3. 提取占位符...")
    try:
        placeholders = processor.find_placeholders()
        print(f"  ✓ 提取完成，找到 {len(placeholders)} 个占位符")

        if placeholders:
            print("\n  占位符列表:")
            for i, ph in enumerate(placeholders, 1):
                print(f"    {i}. {ph}")
        else:
            print("  ✗ 未找到占位符")

        return placeholders

    except Exception as e:
        print(f"  ✗ 提取异常: {e}")
        import traceback
        traceback.print_exc()
        return []

    finally:
        # 关闭PowerPoint
        try:
            processor.close_powerpoint()
        except:
            pass

def main():
    """主测试函数"""
    print("程序流程模拟测试")
    print("="*70)
    print("模拟程序实际加载PPT和提取占位符的完整流程")
    print("="*70)

    files = [
        "E:/poster code/poster-again/10月承保贺报(4).pptx",
        "E:/poster code/poster-again/10月递交贺报(3).pptx"
    ]

    results = {}

    for ppt_path in files:
        if not os.path.exists(ppt_path):
            print(f"文件不存在: {ppt_path}")
            continue

        # 测试PPTXProcessor
        ph_pptx = test_pptx_processor_actual(ppt_path)

        # 测试TemplateProcessor
        ph_template = test_template_processor_actual(ppt_path)

        results[ppt_path] = {
            'pptx': ph_pptx,
            'template': ph_template
        }

    # 总结
    print("\n" + "="*70)
    print("测试总结")
    print("="*70)

    for ppt_path, result in results.items():
        filename = os.path.basename(ppt_path)
        print(f"\n{filename}:")
        print(f"  PPTXProcessor: {len(result['pptx'])} 个占位符")
        print(f"  TemplateProcessor: {len(result['template'])} 个占位符")

    print("\n结论:")
    print("1. 如果两个处理器识别数量不同，说明存在兼容性问题")
    print("2. 如果都识别为0个占位符，说明程序逻辑有问题")
    print("3. 程序日志应该会显示实际使用的处理器类型")

if __name__ == "__main__":
    main()