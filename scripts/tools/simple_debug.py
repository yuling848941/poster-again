#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化版PPT占位符诊断工具
不需要启动PowerPointGUI，仅检查文件
"""

import os
import zipfile
import xml.etree.ElementTree as ET

def check_ppt_placeholders(ppt_path):
    """
    检查PPT文件中的占位符
    通过解析PPTX的XML结构
    """
    try:
        print(f"检查文件: {ppt_path}")
        if not os.path.exists(ppt_path):
            print("文件不存在")
            return

        # PPTX是ZIP文件
        with zipfile.ZipFile(ppt_path, 'r') as zip_file:
            # 获取所有幻灯片文件
            slide_files = [f for f in zip_file.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]

            print(f"\\n总幻灯片数: {len(slide_files)}")
            print("=" * 60)

            total_placeholders = 0

            for slide_file in sorted(slide_files):
                slide_num = slide_file.split('slide')[1].split('.xml')[0]
                print(f"\\n【幻灯片 {slide_num}】")
                print("-" * 40)

                try:
                    # 读取幻灯片XML
                    xml_content = zip_file.read(slide_file).decode('utf-8')
                    root = ET.fromstring(xml_content)

                    # 定义命名空间
                    namespaces = {
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
                    }

                    # 查找所有形状元素
                    shapes = root.findall('.//a:pic', namespaces)
                    text_boxes = root.findall('.//a:txBody', namespaces)
                    shapes_all = root.findall('.//p:sp', namespaces)

                    slide_placeholders = []

                    # 检查每个形状
                    for shape in shapes_all:
                        # 获取形状属性
                        name_attr = shape.get('{http://schemas.openxmlformats.org/presentationml/2006/main}name')
                        if name_attr and name_attr.startswith('ph_'):
                            slide_placeholders.append(name_attr)
                            total_placeholders += 1

                        # 检查是否包含文本
                        text_content = ""
                        text_elements = shape.findall('.//a:t', namespaces)
                        if text_elements:
                            text_content = ''.join([t.text for t in text_elements if t.text])

                        if name_attr and name_attr.startswith('ph_'):
                            print(f"  ✓ 占位符: {name_attr}")
                            if text_content:
                                print(f"    内容: {text_content[:30]}")

                    if not slide_placeholders:
                        print("  - 未找到以ph_开头的占位符")

                except Exception as e:
                    print(f"  解析错误: {e}")

            print("\\n" + "=" * 60)
            print(f"总共找到 {total_placeholders} 个占位符")
            print("=" * 60)

            if total_placeholders == 0:
                print("\\n[问题分析]")
                print("1. 您的PPT中可能没有以'ph_'开头的形状名称")
                print("2. 或者选择的形状类型不正确")
                print("3. 名称中可能包含特殊字符或空格")
                print("\\n[建议解决方案]")
                print("1. 打开PPT文件")
                print("2. 按Alt+F10打开选择窗格")
                print("3. 选中需要作为占位符的形状")
                print("4. 重命名为 'ph_xxx' 格式（如：ph_姓名）")
                print("5. 确保是文本框、占位符等可编辑形状")
            else:
                print(f"\\n✓ 成功找到 {total_placeholders} 个占位符！")
                print("如果程序仍提示未找到，请检查：")
                print("1. 形状必须是可编辑的文本形状")
                print("2. 名称中不能有隐藏字符")
                print("3. 重新保存PPT文件后重试")

    except Exception as e:
        print(f"检查过程中出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        ppt_path = sys.argv[1]
    else:
        ppt_path = input("请输入PPT文件路径: ").strip()

    check_ppt_placeholders(ppt_path)