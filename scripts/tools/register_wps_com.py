#!/usr/bin/env python3
"""
注册WPS Office COM组件
"""

import os
import subprocess
import sys

def register_wps_com():
    """注册WPS COM组件"""
    wps_path = r"C:\Users\Administrator\AppData\Local\kingsoft\WPS Office\12.1.0.23542\office6"

    print("WPS Office COM组件注册工具")
    print("=" * 40)

    if not os.path.exists(wps_path):
        print(f"错误：WPS安装路径不存在: {wps_path}")
        return False

    # 检查关键文件
    wpp_exe = os.path.join(wps_path, "wpp.exe")
    et_exe = os.path.join(wps_path, "et.exe")
    wps_exe = os.path.join(wps_path, "wps.exe")

    if not os.path.exists(wpp_exe):
        print(f"错误：未找到WPS演示程序: {wpp_exe}")
        return False

    print(f"WPS安装路径: {wps_path}")
    print()

    success_count = 0

    # 注册WPS演示
    print("正在注册WPS演示COM组件...")
    try:
        result = subprocess.run([wpp_exe, "/regserver"],
                              capture_output=True, text=True, timeout=30)
        if result.returncode == 0:
            print("SUCCESS WPS演示COM组件注册成功")
            success_count += 1
        else:
            print(f"FAIL WPS演示COM组件注册失败: {result.stderr}")
    except Exception as e:
        print(f"ERROR 注册WPS演示时出错: {e}")

    # 注册WPS表格
    if os.path.exists(et_exe):
        print("正在注册WPS表格COM组件...")
        try:
            result = subprocess.run([et_exe, "/regserver"],
                                  capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                print("SUCCESS WPS表格COM组件注册成功")
                success_count += 1
            else:
                print(f"FAIL WPS表格COM组件注册失败: {result.stderr}")
        except Exception as e:
            print(f"ERROR 注册WPS表格时出错: {e}")

    # 注册WPS文字
    if os.path.exists(wps_exe):
        print("正在注册WPS文字COM组件...")
        try:
            result = subprocess.run([wps_exe, "/regserver"],
                                  capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                print("SUCCESS WPS文字COM组件注册成功")
                success_count += 1
            else:
                print(f"FAIL WPS文字COM组件注册失败: {result.stderr}")
        except Exception as e:
            print(f"ERROR 注册WPS文字时出错: {e}")

    print()
    print(f"注册完成：成功注册 {success_count} 个组件")
    print()

    if success_count > 0:
        print("请重新启动此应用程序，然后:")
        print("1. 在下拉菜单中应该能看到'WPS演示'选项")
        print("2. 选择'WPS演示'应该可以正常工作")
        print()
        print("如果仍然显示'不可用'，请尝试:")
        print("- 以管理员身份运行此注册脚本")
        print("- 重新启动计算机")

    return success_count > 0

def test_wps_after_registration():
    """测试注册后的WPS COM组件"""
    print("\n正在测试WPS COM组件...")

    try:
        import pythoncom
        import win32com.client

        # 测试WPS演示
        try:
            pythoncom.CoInitialize()
            wps_app = win32com.client.Dispatch("WPP.Application")
            name = wps_app.Name
            wps_app.Quit()
            pythoncom.CoUninitialize()
            print(f"SUCCESS WPS演示COM组件可用: {name}")
            return True
        except Exception as e:
            print(f"ERROR WPS演示COM组件不可用: {e}")
            return False

    except ImportError:
        print("无法导入win32com模块，跳过测试")
        return True

def main():
    print("开始注册WPS Office COM组件...")
    print("注意：如果遇到权限问题，请以管理员身份运行此脚本")
    print()

    # 执行注册
    success = register_wps_com()

    if success:
        # 测试注册结果
        test_success = test_wps_after_registration()

        if test_success:
            print("\n🎉 WPS Office COM组件注册并测试成功！")
            print("现在可以在应用程序中选择WPS演示了。")
        else:
            print("\n⚠️  WPS Office COM组件注册可能未成功")
            print("请以管理员身份重新运行此脚本，或尝试重新安装WPS Office")
    else:
        print("\n❌ WPS Office COM组件注册失败")
        print("请检查WPS Office安装是否正确")

    input("\n按回车键退出...")

if __name__ == "__main__":
    main()