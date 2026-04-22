#!/usr/bin/env python3
"""
WPS Office检测工具
详细诊断WPS Office安装和COM注册状态
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client
import winreg
from typing import List, Dict, Optional

def check_com_registration(prog_id: str) -> Dict[str, any]:
    """检查特定ProgID的COM注册情况"""
    result = {
        "prog_id": prog_id,
        "registered": False,
        "clsid": None,
        "path": None,
        "error": None,
        "can_create": False,
        "app_name": None,
        "app_version": None
    }

    try:
        # 检查注册表中的CLSID
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id + "\\CLSID", 0, winreg.KEY_READ) as key:
                clsid = winreg.QueryValueEx(key, "")[0]
                result["clsid"] = clsid
                result["registered"] = True

                # 尝试获取程序路径
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32", 0, winreg.KEY_READ) as path_key:
                        path = winreg.QueryValueEx(path_key, "")[0]
                        result["path"] = path.strip('"')  # 移除可能的引号
                except FileNotFoundError:
                    pass

        except FileNotFoundError:
            result["error"] = "注册表中未找到此ProgID"
            return result

        # 尝试创建COM对象
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch(prog_id)
            result["can_create"] = True

            if hasattr(app, 'Name'):
                result["app_name"] = app.Name
            if hasattr(app, 'Version'):
                result["app_version"] = app.Version

            # 关闭应用程序
            if hasattr(app, 'Quit'):
                app.Quit()

        except Exception as e:
            result["error"] = f"COM创建失败: {str(e)}"

        finally:
            pythoncom.CoUninitialize()

    except Exception as e:
        result["error"] = f"检查失败: {str(e)}"

    return result

def find_wps_installations() -> List[Dict[str, str]]:
    """查找系统中安装的WPS Office"""
    installations = []

    # 检查常见的注册表位置
    registry_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]

    wps_keywords = ["wps", "kingsoft", "wps office"]

    for hkey, path in registry_paths:
        try:
            with winreg.OpenKey(hkey, path, 0, winreg.KEY_READ) as key:
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        i += 1

                        # 检查是否是WPS相关
                        subkey_lower = subkey_name.lower()
                        if any(keyword in subkey_lower for keyword in wps_keywords):
                            try:
                                with winreg.OpenKey(key, subkey_name, 0, winreg.KEY_READ) as subkey:
                                    try:
                                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                    except FileNotFoundError:
                                        display_name = subkey_name

                                    try:
                                        install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                    except FileNotFoundError:
                                        install_location = "未知"

                                    try:
                                        version = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                                    except FileNotFoundError:
                                        version = "未知"

                                    installations.append({
                                        "name": display_name,
                                        "path": install_location,
                                        "version": version,
                                        "registry_key": subkey_name
                                    })
                            except Exception:
                                continue

                    except OSError:
                        break

        except Exception:
            continue

    return installations

def main():
    """主检测函数"""
    print("WPS Office 详细检测工具")
    print("=" * 60)

    # 1. 检查WPS安装情况
    print("\n1. 检查WPS Office安装情况:")
    installations = find_wps_installations()

    if installations:
        print(f"找到 {len(installations)} 个WPS相关安装:")
        for i, inst in enumerate(installations, 1):
            print(f"  {i}. {inst['name']}")
            print(f"     版本: {inst['version']}")
            print(f"     路径: {inst['path']}")
    else:
        print("  未找到WPS Office安装")

    # 2. 检查WPS相关ProgID
    print("\n2. 检查WPS COM ProgID:")
    wps_prog_ids = [
        "KWPP.Application",          # WPS演示（正确）
        "KWPP.Presentation",         # WPS演示
        "KWPP.Presentation.12",      # WPS演示v12
        "WPP.Application",           # WPS演示（标准）
        "WPS.Presentation",          # WPS演示
        "KWps.Application",          # 金山WPS通用（Word）
        "WPS.Application",           # WPS通用
        "Kingsoft.Presentation",     # 金山演示
        "WPS.PresentationApplication" # WPS演示应用
    ]

    available_prog_ids = []
    for prog_id in wps_prog_ids:
        print(f"\n检查 {prog_id}:")
        result = check_com_registration(prog_id)

        if result["registered"]:
            print(f"  SUCCESS 已注册 (CLSID: {result['clsid']})")
        else:
            print(f"  FAIL 未注册")

        if result["path"]:
            print(f"  路径: {result['path']}")

        if result["can_create"]:
            print(f"  SUCCESS COM对象创建成功")
            print(f"    应用名: {result['app_name']}")
            print(f"    版本: {result['app_version']}")
            available_prog_ids.append(prog_id)
        else:
            print(f"  FAIL COM对象创建失败: {result['error']}")

    # 3. 检查PowerPoint作为对比
    print("\n\n3. 检查Microsoft PowerPoint (对比):")
    ppt_result = check_com_registration("PowerPoint.Application")
    print(f"检查 PowerPoint.Application:")

    if ppt_result["registered"]:
        print(f"  SUCCESS 已注册 (CLSID: {ppt_result['clsid']})")
    else:
        print(f"  FAIL 未注册")

    if ppt_result["can_create"]:
        print(f"  SUCCESS COM对象创建成功")
        print(f"    应用名: {ppt_result['app_name']}")
        print(f"    版本: {ppt_result['app_version']}")
    else:
        print(f"  FAIL COM对象创建失败: {ppt_result['error']}")

    # 4. 总结和建议
    print("\n\n4. 总结和建议:")
    print("=" * 40)

    if available_prog_ids:
        print(f"SUCCESS 发现 {len(available_prog_ids)} 个可用的WPS ProgID:")
        for prog_id in available_prog_ids:
            print(f"  - {prog_id}")
        print("\n建议: WPS Office COM接口可用，应该可以在下拉菜单中选择")
    else:
        print("FAIL 没有发现可用的WPS ProgID")

        if installations:
            print("\n可能的原因:")
            print("1. WPS Office已安装但COM组件未注册")
            print("2. 需要以管理员身份运行WPS Office进行注册")
            print("3. WPS Office版本不支持COM接口")
            print("\n建议的解决方案:")
            print("1. 重新安装WPS Office（选择完整安装）")
            print("2. 以管理员身份运行以下命令（在WPS安装目录下）:")
            print("   wps.exe /regserver")
            print("   et.exe /regserver")
            print("   wpp.exe /regserver")
        else:
            print("\n原因: 系统中未安装WPS Office")
            print("建议: 从官方网站下载安装WPS Office")

if __name__ == "__main__":
    main()