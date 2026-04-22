#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
精确分析PPT占位符
深度解析PPTX的XML结构
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom

def analyze_ppt(ppt_path):
    """
    深度分析PPT文件中的形状和占位符
    """
    try:
        print(f"分析文件: {ppt_path}")
        print(f"文件大小: {os.path.getsize(ppt_path)} 字节")
        print("=" * 70)

        if not os.path.exists(ppt_path):
            print("文件不存在!")
            return

        with zipfile.ZipFile(ppt_path, 'r') as zip_file:
            # 获取所有幻灯片
            slide_files = [f for f in zip_file.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            slide_files.sort()

            print(f"\n找到 {len(slide_files)} 张幻灯片")

            all_shapes = []
            all_placeholders = []

            for i, slide_file in enumerate(slide_files, 1):
                if i > 5:  # 只分析前5张幻灯片
                    print(f"\n... 跳过其余 {len(slide_files) - 5} 张幻灯片")
                    break

                print(f"\n{'='*70}")
                print(f"【幻灯片 {i}】- {slide_file}")
                print(f"{'='*70}")

                try:
                    # 读取原始XML
                    xml_content = zip_file.read(slide_file)
                    root = ET.fromstring(xml_content)

                    # 查找所有形状
                    # p:sp = presentation shape
                    shapes = root.findall('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sp')

                    print(f"\n找到 {len(shapes)} 个形状")

                    for j, shape in enumerate(shapes, 1):
                        # 获取形状名称
                        name = None
                        nv_pr = shape.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}nvSpPr')
                        if nv_pr is not None:
                            c_nv_pr = nv_pr.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr')
                            if c_nv_pr is not None:
                                name = c_nv_pr.get('name')

                        # 获取形状类型
                        shape_type = "未知"
                        sp_type = shape.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}spType')
                        if sp_type is not None:
                            type_map = {
                                '1': '标题',
                                '2': '副标题',
                                '3': '正文',
                                '4': '日期',
                                '5': '页脚',
                                '6': '页码',
                                '7': '内容',
                                '13': '图片',
                                '14': '图表',
                                '15': '表格',
                                '16': 'SmartArt',
                                '17': '媒体',
                                '18': '应用',
                                '19': '对话',
                                '20': '实心文字',
                            }
                            shape_type = type_map.get(sp_type.get('val', '0'), f'类型{sp_type.get("val")}')

                        # 获取文本内容
                        text_content = ""
                        text_elements = shape.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t')
                        if text_elements:
                            text_content = ''.join([t.text or '' for t in text_elements[:3]]).strip()

                        # 记录形状信息
                        shape_info = {
                            'slide': i,
                            'shape_index': j,
                            'name': name,
                            'type': shape_type,
                            'text': text_content[:50] if text_content else '',
                            'is_placeholder': name.startswith('ph_') if name else False
                        }
                        all_shapes.append(shape_info)

                        # 显示形状信息
                        if name:
                            if name.startswith('ph_'):
                                print(f"\n  ✓ 占位符 #{j}: {name}")
                                all_placeholders.append(name)
                                if text_content:
                                    print(f"    类型: {shape_type}")
                                    print(f"    文本: {text_content[:30]}")
                            else:
                                print(f"\n  - 形状 #{j}: {name}")
                                print(f"    类型: {shape_type}")
                                if text_content:
                                    print(f"    文本: {text_content[:30]}")
                        else:
                            print(f"\n  ? 形状 #{j}: (无名称)")
                            print(f"    类型: {shape_type}")

                    # 特别检查图片和其他非文本形状
                    pictures = root.findall('.//{http://schemas.openxmlformats.org/presentationml/2006/main}pic')
                    if pictures:
                        print(f"\n  额外找到 {len(pictures)} 个图片形状")

                except Exception as e:
                    print(f"  解析幻灯片 {i} 时出错: {e}")
                    import traceback
                    traceback.print_exc()

            # 总结
            print(f"\n{'='*70}")
            print("分析总结")
            print(f"{'='*70}")
            print(f"总幻灯片数: {len(slide_files)}")
            print(f"总形状数: {len(all_shapes)}")
            print(f"有名称的形状数: {len([s for s in all_shapes if s['name']])}")
            print(f"占位符数量: {len(all_placeholders)}")
            print(f"{'='*70}")

            if all_placeholders:
                print("\n✓ 找到以下占位符:")
                for ph in all_placeholders:
                    print(f"  - {ph}")
                print("\n建议: 您的占位符已正确设置，请重新保存PPT并重启程序测试")
            else:
                print("\n✗ 未找到任何以 'ph_' 开头的占位符")
                print("\n可能的原因:")
                print("1. 形状的Name属性没有设置")
                print("2. 名称中包含特殊字符")
                print("3. PowerPoint版本兼容性问题")
                print("\n解决方案:")
                print("1. 手动在选择窗格中重命名每个形状")
                print("2. 确保名称格式为 'ph_xxx'")
                print("3. 保存为新文件名后重试")

    except Exception as e:
        print(f"\n分析过程中出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        ppt_path = sys.argv[1]
    else:
        ppt_path = input("请输入PPT文件路径: ").strip()

    analyze_ppt(ppt_path)