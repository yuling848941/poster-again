#!/usr/bin/env python3
"""
调试PPT占位符
帮助用户查看PPT模板中实际的形状名称
"""

import sys
import os
import win32com.client

def debug_ppt_placeholders(ppt_path):
    """
    调试PPT模板中的占位符

    Args:
        ppt_path: PPT模板文件路径
    """
    try:
        # 连接到PowerPoint
        print("正在连接到PowerPoint...")
        powerpoint_app = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint_app.Visible = False

        print(f"加载模板: {ppt_path}")
        presentation = powerpoint_app.Presentations.Open(os.path.abspath(ppt_path))

        print("\n" + "=" * 60)
        print("PPT模板中的所有形状详情")
        print("=" * 60)

        total_shapes = 0
        total_placeholders = 0

        # 遍历所有幻灯片
        for slide_idx, slide in enumerate(presentation.Slides, 1):
            print(f"\n【幻灯片 {slide_idx}】")
            print("-" * 40)

            slide_placeholders = []
            slide_shapes = 0

            # 遍历幻灯片中的所有形状
            for shape_idx, shape in enumerate(slide.Shapes, 1):
                slide_shapes += 1
                shape_name = shape.Name

                # 检查是否是占位符
                is_placeholder = shape_name.startswith("ph_")

                if is_placeholder:
                    total_placeholders += 1
                    slide_placeholders.append(shape_name)

                # 显示形状信息
                status = "[占位符]" if is_placeholder else "[普通形状]"
                print(f"  {shape_idx}. 形状名称: {shape_name} {status}")

                # 额外信息
                try:
                    shape_type = shape.Type
                    type_name = "未知"
                    if shape_type == 1:  # msoAutoShape
                        type_name = "自动形状"
                    elif shape_type == 13:  # msoCallout
                        type_name = "标注"
                    elif shape_type == 14:  # msoChart
                        type_name = "图表"
                    elif shape_type == 15:  # msoComment
                        type_name = "批注"
                    elif shape_type == 17:  # msoFreeform
                        type_name = "自由形状"
                    elif shape_type == 19:  # msoGroup
                        type_name = "组合"
                    elif shape_type == 22:  # msoLine
                        type_name = "线条"
                    elif shape_type == 24:  # msoOLEControlObject
                        type_name = "OLE控件"
                    elif shape_type == 25:  # msoOLEControlObject
                        type_name = "OLE对象"
                    elif shape_type == 28:  # msoPicture
                        type_name = "图片"
                    elif shape_type == 30:  # msoPlaceholder
                        type_name = "占位符"
                    elif shape_type == 31:  # msoTextEffect
                        type_name = "文本效果"
                    elif shape_type == 32:  # msoTextBox
                        type_name = "文本框"
                    elif shape_type == 33:  # msoPlaceholder
                        type_name = "占位符"

                    print(f"     类型: {type_name}")

                    # 如果是文本框或占位符，尝试获取文本内容
                    if shape_type in [1, 17, 30, 32, 33]:  # 可能的文本形状
                        try:
                            if hasattr(shape, "TextFrame") and shape.TextFrame:
                                text_content = shape.TextFrame.TextRange.Text
                                if text_content.strip():
                                    print(f"     文本内容: {repr(text_content[:50])}")
                        except:
                            pass

                except Exception as e:
                    print(f"     (无法获取类型信息)")

            total_shapes += slide_shapes

            if slide_placeholders:
                print(f"  -> 找到 {len(slide_placeholders)} 个占位符:")
                for ph in slide_placeholders:
                    print(f"     - {ph}")
            else:
                print(f"  -> 未找到占位符")

        print("\n" + "=" * 60)
        print("总结")
        print("=" * 60)
        print(f"总幻灯片数: {presentation.Slides.Count}")
        print(f"总形状数: {total_shapes}")
        print(f"占位符数量: {total_placeholders}")
        print("=" * 60)

        # 关闭PowerPoint
        presentation.Close()
        powerpoint_app.Quit()

        if total_placeholders == 0:
            print("\n❌ 未找到任何以 'ph_' 开头的占位符！")
            print("\n可能的解决方案:")
            print("1. 检查形状名称是否确实以 'ph_' 开头")
            print("2. 重新命名形状：")
            print("   - 选中形状")
            print("   - 右键 -> '编辑 Alt 文本' 或在选择窗格中修改名称")
            print("   - 确保名称以 'ph_' 开头，如：ph_姓名、ph_日期等")
            print("3. 确保选择的是可编辑的形状（文本框、占位符等）")
        else:
            print(f"\n✅ 找到 {total_placeholders} 个占位符！")

        return total_placeholders

    except Exception as e:
        print(f"\n❌ 调试过程中出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return 0

def main():
    """主函数"""
    print("PPT占位符调试工具")
    print("=" * 60)
    print("此工具将帮助您查看PPT模板中的实际形状名称")
    print("=" * 60)

    # 获取PPT文件路径
    ppt_path = input("\n请输入PPT模板文件路径: ").strip()

    if not os.path.exists(ppt_path):
        print(f"❌ 文件不存在: {ppt_path}")
        return

    if not ppt_path.lower().endswith('.pptx'):
        print("❌ 请选择.pptx文件")
        return

    # 调试占位符
    result = debug_ppt_placeholders(ppt_path)

    print("\n调试完成！")
    print("=" * 60)

if __name__ == "__main__":
    main()