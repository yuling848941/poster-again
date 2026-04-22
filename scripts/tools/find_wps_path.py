#!/usr/bin/env python3
"""
查找WPS Office安装路径并提供COM注册指导
"""

import os
import winreg
import sys

def find_wps_installation_path():
    """查找WPS Office的安装路径"""
    registry_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\wps.exe"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Kingsoft\Office"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Kingsoft\Office"),
    ]

    wps_paths = []

    for hkey, path in registry_paths:
        try:
            with winreg.OpenKey(hkey, path, 0, winreg.KEY_READ) as key:
                try:
                    # 尝试直接获取路径
                    wps_path = winreg.QueryValueEx(key, "")[0]
                    if wps_path and os.path.exists(wps_path):
                        wps_paths.append(wps_path)
                except FileNotFoundError:
                    # 如果直接路径不存在，遍历子项
                    i = 0
                    while True:
                        try:
                            subkey_name = winreg.EnumKey(key, i)
                            i += 1
                            try:
                                with winreg.OpenKey(key, subkey_name, 0, winreg.KEY_READ) as subkey:
                                    try:
                                        install_path = winreg.QueryValueEx(subkey, "InstallPath")[0]
                                        if install_path:
                                            wps_exe = os.path.join(install_path, "wps.exe")
                                            if os.path.exists(wps_exe):
                                                wps_paths.append(wps_exe)

                                            wpp_exe = os.path.join(install_path, "wpp.exe")
                                            if os.path.exists(wpp_exe):
                                                wps_paths.append(wpp_exe)
                                    except FileNotFoundError:
                                        pass
                            except OSError:
                                break
                        except OSError:
                            break
        except FileNotFoundError:
            continue

    # 去重
    return list(set(wps_paths))

def check_common_wps_locations():
    """检查常见的WPS安装位置"""
    common_paths = [
        r"C:\Program Files (x86)\Kingsoft\WPS Office",
        r"C:\Program Files\Kingsoft\WPS Office",
        r"C:\Program Files\WPS Office",
        r"C:\Program Files (x86)\WPS Office",
        r"D:\Program Files (x86)\Kingsoft\WPS Office",
        r"D:\Program Files\Kingsoft\WPS Office",
        r"D:\Program Files\WPS Office",
        r"D:\Program Files (x86)\WPS Office",
    ]

    found_paths = []

    for base_path in common_paths:
        if os.path.exists(base_path):
            # 查找所有版本
            for item in os.listdir(base_path):
                item_path = os.path.join(base_path, item)
                if os.path.isdir(item_path):
                    # 检查是否有wps.exe或wpp.exe
                    wps_exe = os.path.join(item_path, "wps.exe")
                    wpp_exe = os.path.join(item_path, "wpp.exe")
                    et_exe = os.path.join(item_path, "et.exe")

                    if os.path.exists(wps_exe):
                        found_paths.append(wps_exe)
                    if os.path.exists(wpp_exe):
                        found_paths.append(wpp_exe)
                    if os.path.exists(et_exe):
                        found_paths.append(et_exe)

    return found_paths

def main():
    print("WPS Office 安装路径查找工具")
    print("=" * 50)

    print("\n1. 从注册表查找WPS安装路径:")
    registry_paths = find_wps_installation_path()

    if registry_paths:
        print("  找到以下WPS程序:")
        for path in registry_paths:
            print(f"    - {path}")
    else:
        print("  未在注册表中找到WPS安装路径")

    print("\n2. 检查常见安装位置:")
    common_paths = check_common_wps_locations()

    if common_paths:
        print("  找到以下WPS程序:")
        for path in common_paths:
            print(f"    - {path}")
    else:
        print("  未在常见位置找到WPS")

    all_paths = registry_paths + common_paths

    if all_paths:
        print("\n3. COM注册指导:")
        print("=" * 30)

        wps_exe = None
        wpp_exe = None
        et_exe = None

        for path in all_paths:
            if path.endswith("wps.exe"):
                wps_exe = path
            elif path.endswith("wpp.exe"):
                wpp_exe = path
            elif path.endswith("et.exe"):
                et_exe = path

        print("请按以下步骤注册WPS COM组件:")
        print("1. 以管理员身份打开命令提示符")
        print("2. 运行以下命令:")

        if wpp_exe:
            wpp_dir = os.path.dirname(wpp_exe)
            print(f'   cd "{wpp_dir}"')
            print("   wpp.exe /regserver")
            print()

        if et_exe:
            et_dir = os.path.dirname(et_exe)
            print(f'   cd "{et_dir}"')
            print("   et.exe /regserver")
            print()

        if wps_exe:
            wps_dir = os.path.dirname(wps_exe)
            print(f'   cd "{wps_dir}"')
            print("   wps.exe /regserver")
            print()

        print("3. 重新启动此应用程序")
        print("4. 在下拉菜单中应该就能看到可用的WPS选项了")

    else:
        print("\nERROR: 未找到WPS Office安装")
        print("请确认WPS Office已正确安装")

if __name__ == "__main__":
    main()