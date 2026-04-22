#!/usr/bin/env python3
"""
快速WPS COM组件注册
"""

import os
import sys

def main():
    print("WPS Office COM组件快速注册")
    print("=" * 40)

    wps_path = r"C:\Users\Administrator\AppData\Local\kingsoft\WPS Office\12.1.0.23542\office6"
    wpp_exe = os.path.join(wps_path, "wpp.exe")

    print(f"WPS安装路径: {wps_path}")
    print(f"WPS演示程序: {wpp_exe}")
    print()

    if not os.path.exists(wpp_exe):
        print("ERROR: 未找到WPS演示程序")
        input("按回车键退出...")
        return

    print("正在启动WPS演示程序进行COM注册...")
    print("如果出现UAC提示，请点击'是'")
    print()

    try:
        # 直接启动WPS，让它自动注册COM组件
        os.startfile(wpp_exe)
        print("SUCCESS: WPS演示程序已启动")
        print()
        print("请按照以下步骤操作:")
        print("1. 让WPS演示程序完全启动")
        print("2. 然后关闭WPS演示程序")
        print("3. 重新启动此PPT应用程序")
        print("4. 在下拉菜单中选择WPS演示")
        print()
        print("COM组件应该在程序关闭时自动注册")

    except Exception as e:
        print(f"ERROR: 启动WPS失败: {e}")
        print("请手动启动WPS演示程序")

    input("\n按回车键退出...")

if __name__ == "__main__":
    main()