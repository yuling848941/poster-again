#!/usr/bin/env python
"""
生成Poster Again程序的图标文件
"""

from PIL import Image, ImageDraw
import os

def create_icon():
    """创建图标"""
    # 创建多尺寸图标
    sizes = [16, 32, 48, 64, 128, 256]
    images = []

    for size in sizes:
        # 创建透明背景
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # 计算尺寸
        padding = size // 8
        icon_size = size - 2 * padding

        # 绘制圆形背景（蓝色渐变效果，用纯色代替）
        circle_bbox = [padding, padding, size - padding, size - padding]
        draw.ellipse(circle_bbox, fill=(37, 99, 235, 255))  # #2563EB

        # 绘制内圆（白色）
        inner_padding = padding + icon_size // 6
        inner_bbox = [inner_padding, inner_padding, size - inner_padding, size - inner_padding]
        draw.ellipse(inner_bbox, fill=(255, 255, 255, 255))

        # 绘制"P"字母（使用简单几何图形代替）
        # P的上半部分
        p_left = padding + icon_size // 4
        p_top = padding + icon_size // 6
        p_right = padding + icon_size * 3 // 5
        p_bottom = padding + icon_size // 2
        draw.rectangle([p_left, p_top, p_right, p_bottom], fill=(37, 99, 235, 255))

        # P的右侧弧形
        arc_bbox = [p_right - icon_size // 6, p_top, p_right, p_bottom]
        draw.pieslice(arc_bbox, start=270, end=90, fill=(37, 99, 235, 255))

        # 绘制竖线
        line_left = padding + icon_size // 8
        line_right = padding + icon_size // 4
        line_top = padding + icon_size // 3
        line_bottom = size - padding
        draw.rectangle([line_left, line_top, line_right, line_bottom], fill=(37, 99, 235, 255))

        images.append(img)

    # 保存为ICO文件
    icon_path = "resources/icon.ico"
    os.makedirs(os.path.dirname(icon_path), exist_ok=True)
    images[0].save(icon_path, format='ICO', sizes=[(img.width, img.height) for img in images])

    print(f"[OK] Icon generated: {icon_path}")

    # 生成PNG预览
    png_path = "resources/icon.png"
    images[-1].save(png_path, format='PNG')
    print(f"[OK] PNG preview generated: {png_path}")

if __name__ == "__main__":
    create_icon()