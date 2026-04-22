#!/usr/bin/env python3
"""
检查生成的PPT文件内容
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def check_generated_ppt():
    """检查生成的PPT文件"""
    try:
        from src.core.template_processor import TemplateProcessor

        output_file = "output/output_1.pptx"

        if not os.path.exists(output_file):
            print(f"生成的文件不存在: {output_file}")
            return

        print(f"正在检查生成的文件: {output_file}")

        processor = TemplateProcessor()
        processor.load_template(output_file)

        # 获取所有形状和文本内容
        shapes_info = processor.get_shapes_info()

        print(f"\n文件中的文本内容:")

        found_custom_text = False

        for slide_num, shapes in shapes_info.items():
            print(f"\n幻灯片 {slide_num}:")
            for shape in shapes:
                if shape['text'] and shape['text'].strip():
                    text = shape['text']
                    print(f"  {text}")

                    # 检查是否包含自定义文本
                    if "中支" in text or "1件" in text or "家族" in text:
                        print(f"    [发现自定义文本!] 包含: {text}")
                        found_custom_text = True

                    # 检查占位符相关的文本
                    if any(keyword in text for keyword in ["中心支公司", "保险公司", "所属家庭"]):
                        print(f"    [相关数据] 包含: {text}")

        print(f"\n检查结果:")
        if found_custom_text:
            print(f"  [问题确认] 生成的PPT包含自定义文本，这说明配置被错误应用了")
        else:
            print(f"  [正常] 生成的PPT不包含自定义文本")

        return found_custom_text

    except Exception as e:
        print(f"检查PPT文件失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    found_text = check_generated_ppt()