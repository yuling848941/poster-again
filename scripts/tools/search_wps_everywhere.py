#!/usr/bin/env python3
"""
全面搜索WPS Office安装
"""

import os
import winreg
import sys
from pathlib import Path

def search_drives_for_wps():
    """在所有驱动器中搜索WPS"""
    # 获取所有可用驱动器
    drives = []
    for letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        drive = f"{letter}:\\"
        if os.path.exists(drive):
            drives.append(drive)

    wps_found = []

    print("正在搜索以下驱动器:", ", ".join(drives))

    for drive in drives:
        print(f"\n搜索驱动器 {drive}...")
        try:
            # 搜索常见安装目录
            search_paths = [
                drive + "Program Files\\",
                drive + "Program Files (x86)\\",
                drive + "Program Files\\Kingsoft\\",
                drive + "Program Files (x86)\\Kingsoft\\",
                drive + "Kingsoft\\",
            ]

            for search_path in search_paths:
                if os.path.exists(search_path):
                    try:
                        for item in os.listdir(search_path):
                            item_path = os.path.join(search_path, item)
                            if os.path.isdir(item_path):
                                # 检查是否是WPS相关目录
                                item_lower = item.lower()
                                if "wps" in item_lower or "kingsoft" in item_lower:
                                    print(f"  找到相关目录: {item_path}")

                                    # 在该目录中搜索可执行文件
                                    for root, dirs, files in os.walk(item_path):
                                        for file in files:
                                            if file.lower() in ['wps.exe', 'wpp.exe', 'et.exe']:
                                                full_path = os.path.join(root, file)
                                                wps_found.append(full_path)
                                                print(f"    找到: {full_path}")
                    except PermissionError:
                        print(f"  权限不足，无法访问: {search_path}")
                        continue
        except Exception as e:
            print(f"  搜索 {drive} 时出错: {e}")

    return wps_found

def check_wps_running_processes():
    """检查正在运行的WPS进程"""
    try:
        import psutil
        wps_processes = []

        for proc in psutil.process_iter(['name', 'exe']):
            try:
                proc_info = proc.info
                if proc_info['name'] and 'wps' in proc_info['name'].lower():
                    wps_processes.append(proc_info)
                    print(f"找到运行中的WPS进程: {proc_info['name']} - {proc_info['exe']}")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        return wps_processes
    except ImportError:
        print("需要安装psutil模块来检查运行进程")
        print("请运行: pip install psutil")
        return []

def main():
    print("WPS Office 全面搜索工具")
    print("=" * 50)

    print("\n1. 检查正在运行的WPS进程:")
    running_processes = check_wps_running_processes()

    print("\n2. 在所有驱动器中搜索WPS:")
    wps_executables = search_drives_for_wps()

    print("\n3. 搜索结果总结:")
    print("=" * 30)

    if wps_executables:
        print("找到以下WPS可执行文件:")
        wps_exe = None
        wpp_exe = None
        et_exe = None

        for exe in wps_executables:
            print(f"  - {exe}")
            if exe.lower().endswith('wps.exe'):
                wps_exe = exe
            elif exe.lower().endswith('wpp.exe'):
                wpp_exe = exe
            elif exe.lower().endswith('et.exe'):
                et_exe = exe

        print("\nCOM注册步骤:")
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

        print("3. 重启应用程序后WPS选项应该可用")

    elif running_processes:
        print("虽然找到了运行中的WPS进程，但没有找到安装文件")
        print("可能WPS是从Windows Store安装的，或者安装在其他位置")

        # 尝试从进程路径推断安装位置
        for proc in running_processes:
            if proc['exe']:
                exe_path = proc['exe']
                if os.path.exists(exe_path):
                    exe_dir = os.path.dirname(exe_path)
                    print(f"\n根据进程推断的安装目录: {exe_dir}")
                    print(f"请尝试在该目录下查找 wpp.exe 并运行:")
                    print(f'cd "{exe_dir}"')
                    print("wpp.exe /regserver")

    else:
        print("未找到WPS Office安装")
        print("建议:")
        print("1. 确认WPS Office已正确安装")
        print("2. 尝试重新安装WPS Office（选择完整安装）")
        print("3. 如果是从Windows Store安装的，建议下载完整版重新安装")

if __name__ == "__main__":
    main()