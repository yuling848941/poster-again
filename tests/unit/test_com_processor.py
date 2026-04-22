#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接测试COM接口处理器
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import logging
from src.core.processors.microsoft_processor import MicrosoftProcessor

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def test_microsoft_processor(ppt_path):
    """
    直接测试MicrosoftProcessor
    """
    print(f"\n测试文件: {ppt_path}")
    print("=" * 70)

    processor = MicrosoftProcessor()

    try:
        # 1. 连接PowerPoint
        print("1. 连接PowerPoint...")
        processor.connect_to_application()
        print("   ✓ 连接成功")

        # 2. 加载模板
        print("2. 加载模板...")
        success = processor.load_template(ppt_path)
        if not success:
            print("   ✗ 加载失败")
            return []
        print("   ✓ 加载成功")

        # 3. 提取占位符
        print("3. 提取占位符...")
        placeholders = processor.find_placeholders()
        print(f"   ✓ 找到 {len(placeholders)} 个占位符")

        if placeholders:
            print("\n   占位符列表:")
            for i, ph in enumerate(placeholders, 1):
                print(f"     {i}. {ph}")
        else:
            print("   ✗ 未找到占位符")

        return placeholders

    except Exception as e:
        print(f"   ✗ 异常: {e}")
        import traceback
        traceback.print_exc()
        return []

    finally:
        # 关闭PowerPoint
        try:
            processor.close_application()
            print("\n4. 已关闭PowerPoint")
        except:
            pass

if __name__ == "__main__":
    # 测试两个文件
    ph1 = test_microsoft_processor("E:/poster code/poster-again/10月承保贺报(4).pptx")
    ph2 = test_microsoft_processor("E:/poster code/poster-again/10月递交贺报(3).pptx")

    print("\n" + "=" * 70)
    print("测试总结")
    print("=" * 70)
    print(f"(4) 文件: {len(ph1)} 个占位符")
    print(f"(3) 文件: {len(ph2)} 个占位符")

    if len(ph1) > 0 and len(ph2) > 0:
        print("\n✓ COM接口可以正确识别占位符")
        print("问题可能在于程序没有使用COM接口，而是降级到了python-pptx")
    elif len(ph1) == 0 and len(ph2) == 0:
        print("\n✗ COM接口无法识别占位符")
        print("问题在于COM接口的find_placeholders()方法")
    else:
        print("\n⚠️ 两个文件识别结果不一致")
        print("说明COM接口对不同文件的处理可能存在差异")
