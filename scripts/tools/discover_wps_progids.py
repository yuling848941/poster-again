#!/usr/bin/env python3
"""
发现WPS Office的正确ProgID
"""

import os
import sys
import pythoncom
import win32com.client
import winreg

def find_wps_progids_from_registry():
    """从注册表中查找所有WPS相关的ProgID"""
    wps_progids = []

    try:
        # 搜索HKEY_CLASSES_ROOT
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "", 0, winreg.KEY_READ) as root_key:
            i = 0
            while True:
                try:
                    prog_id = winreg.EnumKey(root_key, i)
                    i += 1

                    # 查找WPS相关的ProgID
                    prog_id_lower = prog_id.lower()
                    if any(keyword in prog_id_lower for keyword in ['wps', 'kingsoft', 'wpp', 'wpt']):
                        # 检查是否有CLSID
                        try:
                            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID", 0, winreg.KEY_READ) as clsid_key:
                                clsid = winreg.QueryValueEx(clsid_key, "")[0]
                                wps_progids.append((prog_id, clsid))
                        except FileNotFoundError:
                            pass

                except OSError:
                    break

    except Exception as e:
        print(f"搜索注册表时出错: {e}")

    return wps_progids

def test_prog_id(prog_id):
    """测试特定ProgID是否可用"""
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch(prog_id)

        # 获取应用信息
        name = getattr(app, 'Name', '未知应用')
        version = getattr(app, 'Version', '未知版本')

        # 检查是否有演示功能
        has_presentations = hasattr(app, 'Presentations')
        has_slides = hasattr(app, 'Slides') if has_presentations else False

        # 检查应用类型
        app_type = "未知"
        if any(keyword in name.lower() for keyword in ['word', '文字', 'wps']):
            if has_presentations:
                app_type = "WPS文字（误报）"
            else:
                app_type = "WPS文字"
        elif any(keyword in name.lower() for keyword in ['presentation', 'ppt', '演示', 'wpp']):
            app_type = "WPS演示"
        elif any(keyword in name.lower() for keyword in ['excel', '表格', 'et']):
            app_type = "WPS表格"
        elif 'wps' in name.lower():
            app_type = "WPS通用"

        # 关闭应用
        if hasattr(app, 'Quit'):
            app.Quit()

        pythoncom.CoUninitialize()

        return {
            "prog_id": prog_id,
            "name": name,
            "version": version,
            "app_type": app_type,
            "has_presentations": has_presentations,
            "has_slides": has_slides,
            "can_create": True
        }

    except Exception as e:
        return {
            "prog_id": prog_id,
            "error": str(e),
            "can_create": False
        }

def main():
    print("WPS Office ProgID 发现工具")
    print("=" * 50)

    # 1. 从注册表查找
    print("\n1. 从注册表中查找WPS相关的ProgID:")
    registry_progids = find_wps_progids_from_registry()

    if registry_progids:
        print(f"找到 {len(registry_progids)} 个WPS相关的ProgID:")
        for prog_id, clsid in registry_progids:
            print(f"  - {prog_id} (CLSID: {clsid})")
    else:
        print("未在注册表中找到WPS相关的ProgID")

    # 2. 测试常见ProgID
    print("\n2. 测试常见的WPS ProgID:")
    common_progids = [
        "WPP.Application",
        "WPS.Presentation",
        "Kingsoft.Presentation",
        "WPS.PresentationApplication",
        "WPP.Application.1",
        "WPS.Application",
        "KWps.Application",  # 这个是Word，不是演示
        "ET.Application",    # 表格
    ]

    valid_progids = []

    for prog_id in common_progids:
        print(f"\n测试 {prog_id}:")
        result = test_prog_id(prog_id)

        if result["can_create"]:
            print(f"  SUCCESS 可以创建")
            print(f"    应用名: {result['name']}")
            print(f"    版本: {result['version']}")
            print(f"    应用类型: {result['app_type']}")
            print(f"    有演示功能: {result['has_presentations']}")

            if result['app_type'] == "WPS演示" or (result['has_presentations'] and result['has_slides']):
                valid_progids.append(prog_id)
                print("    *** 这是WPS演示程序! ***")
            elif result['app_type'] == "WPS通用" and result['has_presentations']:
                valid_progids.append(prog_id)
                print("    *** 这可能是WPS演示程序! ***")
        else:
            print(f"  FAIL 无法创建: {result['error']}")

    # 3. 总结
    print("\n\n3. 总结:")
    print("=" * 30)

    if valid_progids:
        print(f"找到 {len(valid_progids)} 个可用的WPS演示ProgID:")
        for prog_id in valid_progids:
            print(f"  - {prog_id}")

        print("\n建议:")
        print("1. 更新 WPSProcessor 类中的 prog_ids 列表")
        print("2. 将有效的ProgID放在列表前面，优先尝试")

    else:
        print("未找到可用的WPS演示ProgID")
        print("\n可能的解决方案:")
        print("1. 检查WPS是否支持COM接口")
        print("2. 尝试启动WPS演示程序，然后以管理员身份运行:")
        print("   wpp.exe /regserver")
        print("3. 考虑使用WPS的文档API而不是COM接口")

if __name__ == "__main__":
    main()