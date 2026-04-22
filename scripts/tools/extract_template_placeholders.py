#!/usr/bin/env python3
"""
提取PPT模板中的占位符
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def extract_placeholders():
    """提取PPT模板占位符"""
    try:
        from src.core.template_processor import TemplateProcessor

        template_file = "10月递交贺报(3).pptx"

        if not os.path.exists(template_file):
            print(f"模板文件不存在: {template_file}")
            return

        print(f"正在分析模板: {template_file}")

        processor = TemplateProcessor()
        processor.load_template(template_file)

        placeholders = processor.find_placeholders()

        print(f"\n找到 {len(placeholders)} 个占位符:")
        for i, placeholder in enumerate(placeholders, 1):
            print(f"  {i}. {placeholder}")

        # 检查是否包含配置中的占位符
        config_placeholders = ["ph_中心支公司", "ph_保险种类", "ph_所属家庭"]

        print(f"\n配置文件中的占位符:")
        for config_ph in config_placeholders:
            if config_ph in placeholders:
                print(f"  [OK] {config_ph} - 在模板中找到")
            else:
                # 检查是否有相似的占位符
                similar = [p for p in placeholders if "中心支" in p or "中支" in p]
                if similar:
                    print(f"  [MISMATCH] {config_ph} - 未找到，但找到相似的: {similar}")
                else:
                    print(f"  [NOT FOUND] {config_ph} - 未找到")

        return placeholders

    except Exception as e:
        print(f"提取占位符失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    placeholders = extract_placeholders()