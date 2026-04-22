#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试占位符修复
验证程序能否正确识别PPT中的占位符
"""

import sys
import os
import win32com.client
import logging

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def test_placeholders(ppt_path):
    """
    测试占位符识别
    """
    try:
        print(f"测试文件: {ppt_path}")
        print("=" * 70)

        if not os.path.exists(ppt_path):
            print("文件不存在!")
            return False

        # 连接到PowerPoint
        print("正在连接PowerPoint...")
        powerpoint_app = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint_app.Visible = False

        # 加载模板
        print("加载模板...")
        presentation = powerpoint_app.Presentations.Open(os.path.abspath(ppt_path))

        print(f"模板已加载，共 {presentation.Slides.Count} 张幻灯片")
        print("-" * 70)

        # 遍历所有幻灯片和形状
        total_placeholders = 0
        found_placeholders = []

        for slide_idx, slide in enumerate(presentation.Slides, 1):
            # 只显示前5张幻灯片
            if slide_idx > 5:
                break

            print(f"\n【幻灯片 {slide_idx}】")
            slide_placeholders = 0

            for shape in slide.Shapes:
                shape_name = shape.Name

                if shape_name and shape_name.startswith("ph_"):
                    slide_placeholders += 1
                    total_placeholders += 1
                    found_placeholders.append(shape_name)
                    print(f"  [占位符] {shape_name}")
                elif shape_name:
                    print(f"  [形状] {shape_name}")

            if slide_placeholders == 0:
                print("  (无占位符)")

        # 关闭PowerPoint
        presentation.Close()
        powerpoint_app.Quit()

        print("\n" + "=" * 70)
        print("测试结果")
        print("=" * 70)
        print(f"总占位符数量: {total_placeholders}")

        if found_placeholders:
            print("\n找到的占位符:")
            for i, ph in enumerate(found_placeholders, 1):
                print(f"  {i}. {ph}")

            print("\n" + "=" * 70)
            print("SUCCESS: 程序可以正确识别占位符!")
            print("=" * 70)
            print("\n建议:")
            print("1. 保存此PPT模板")
            print("2. 在程序中重新加载模板")
            print("3. 运行自动匹配测试")
            return True
        else:
            print("\nFAIL: 未找到占位符")
            return False

    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    ppt_path = input("请输入PPT文件路径: ").strip()

    if not ppt_path:
        ppt_path = "E:/poster code/poster-again/10月承保贺报(4).pptx"
        print(f"使用默认路径: {ppt_path}")

    success = test_placeholders(ppt_path)

    if success:
        print("\n✓ 占位符识别正常")
        print("您现在可以在程序中正常使用这个模板了！")
    else:
        print("\n✗ 占位符识别失败")
        print("请检查模板中的占位符设置")