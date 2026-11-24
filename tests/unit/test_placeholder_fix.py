#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试占位符修复
验证 PPTXProcessor（纯 python-pptx）能否正确识别 PPT 中的占位符
"""

import os
import sys

# 将项目根目录加入 sys.path，确保 src 包可导入
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
sys.path.insert(0, PROJECT_ROOT)

import logging

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

from src.core.processors.pptx_processor import PPTXProcessor


def test_placeholders(ppt_path):
    """测试占位符识别"""
    print(f"测试文件: {ppt_path}")
    print("=" * 70)

    if not os.path.exists(ppt_path):
        print("文件不存在!")
        return False

    processor = PPTXProcessor()
    try:
        if not processor.load_template(ppt_path):
            print("加载模板失败")
            return False

        placeholders = processor.find_placeholders()
        print(f"共找到 {len(placeholders)} 个占位符")
        for i, ph in enumerate(placeholders, 1):
            print(f"  {i}. {ph}")

        return len(placeholders) > 0

    finally:
        processor.close()


if __name__ == "__main__":
    # 使用仓库内实际存在的模板，避免交互式 input()
    candidates = [
        os.path.join(PROJECT_ROOT, "samples/templates/10月递交贺报(3).pptx"),
        os.path.join(PROJECT_ROOT, "samples/templates/10月承保贺报(3).pptx"),
        os.path.join(PROJECT_ROOT, "samples/templates/01、每日贺报 - 发2 .pptx"),
    ]
    ppt_path = next((p for p in candidates if os.path.exists(p)), candidates[0])
    print(f"使用模板: {ppt_path}")

    success = test_placeholders(ppt_path)
    print("\nOK 占位符识别正常" if success else "\nX 占位符识别失败")
    sys.exit(0 if success else 1)
