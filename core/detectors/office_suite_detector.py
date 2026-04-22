"""
Office 套件检测器
检测系统中安装的 Microsoft Office 或 WPS Office
"""

import os
import logging
from typing import Optional, Dict, Any
from pathlib import Path

logger = logging.getLogger(__name__)


class OfficeSuiteDetector:
    """检测系统中安装的 Office 套件"""

    # Microsoft Office 可能的安装路径
    MS_OFFICE_PATHS = [
        r"C:\Program Files\Microsoft Office\root\Office16",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16",
        r"C:\Program Files\Microsoft Office\Office16",
        r"C:\Program Files (x86)\Microsoft Office\Office16",
        r"C:\Program Files\Microsoft Office\root\Office15",
        r"C:\Program Files (x86)\Microsoft Office\root\Office15",
        r"C:\Program Files\Microsoft Office\Office15",
        r"C:\Program Files (x86)\Microsoft Office\Office15",
    ]

    # WPS Office 可能的安装路径
    WPS_PATHS = [
        r"C:\Program Files\Kingsoft\WPS Office",
        r"C:\Program Files (x86)\Kingsoft\WPS Office",
        r"C:\Program Files\WPS Office",
        r"C:\Program Files (x86)\WPS Office",
    ]

    @classmethod
    def detect_office_suite(cls) -> Optional[Dict[str, Any]]:
        """
        检测系统中安装的 Office 套件

        Returns:
            Optional[Dict]: 检测到的 Office 套件信息，如果没有安装则返回 None
            {
                'type': 'microsoft' | 'wps',
                'name': str,
                'path': str,
                'version': str
            }
        """
        # 优先检测 Microsoft Office
        ms_office = cls._detect_microsoft_office()
        if ms_office:
            logger.info(f"检测到 Microsoft Office: {ms_office['name']}")
            return ms_office

        # 然后检测 WPS Office
        wps_office = cls._detect_wps_office()
        if wps_office:
            logger.info(f"检测到 WPS Office: {wps_office['name']}")
            return wps_office

        logger.warning("未检测到 Microsoft Office 或 WPS Office")
        return None

    @classmethod
    def _detect_microsoft_office(cls) -> Optional[Dict[str, Any]]:
        """检测 Microsoft Office"""
        try:
            # 检查 PowerPoint 可执行文件
            for office_path in cls.MS_OFFICE_PATHS:
                ppt_exe = os.path.join(office_path, "POWERPNT.EXE")
                if os.path.exists(ppt_exe):
                    # 尝试获取版本信息
                    version = cls._get_file_version(ppt_exe)
                    return {
                        'type': 'microsoft',
                        'name': f'Microsoft Office PowerPoint',
                        'path': office_path,
                        'exe_path': ppt_exe,
                        'version': version or '未知版本'
                    }

            # 尝试通过注册表检测
            return cls._detect_ms_office_from_registry()

        except Exception as e:
            logger.debug(f"检测 Microsoft Office 时出错: {e}")
            return None

    @classmethod
    def _detect_wps_office(cls) -> Optional[Dict[str, Any]]:
        """检测 WPS Office"""
        try:
            # 检查 WPS 可执行文件
            for wps_path in cls.WPS_PATHS:
                # WPS 可能有多个版本
                for root, dirs, files in os.walk(wps_path):
                    for file in files:
                        if file.lower() in ['wps.exe', 'wpp.exe', 'et.exe']:
                            exe_path = os.path.join(root, file)
                            version = cls._get_file_version(exe_path)
                            return {
                                'type': 'wps',
                                'name': 'WPS Office',
                                'path': root,
                                'exe_path': exe_path,
                                'version': version or '未知版本'
                            }

            # 尝试通过注册表检测
            return cls._detect_wps_from_registry()

        except Exception as e:
            logger.debug(f"检测 WPS Office 时出错: {e}")
            return None

    @classmethod
    def _detect_ms_office_from_registry(cls) -> Optional[Dict[str, Any]]:
        """从注册表检测 Microsoft Office"""
        try:
            import winreg

            # 尝试打开 Office 注册表项
            key_paths = [
                r"SOFTWARE\Microsoft\Office\16.0\PowerPoint\InstallRoot",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\PowerPoint\InstallRoot",
                r"SOFTWARE\Microsoft\Office\15.0\PowerPoint\InstallRoot",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\15.0\PowerPoint\InstallRoot",
            ]

            for key_path in key_paths:
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    path, _ = winreg.QueryValueEx(key, "Path")
                    winreg.CloseKey(key)

                    if path and os.path.exists(path):
                        ppt_exe = os.path.join(path, "POWERPNT.EXE")
                        if os.path.exists(ppt_exe):
                            return {
                                'type': 'microsoft',
                                'name': 'Microsoft Office PowerPoint',
                                'path': path,
                                'exe_path': ppt_exe,
                                'version': '16.0' if '16.0' in key_path else '15.0'
                            }
                except:
                    continue

        except Exception as e:
            logger.debug(f"从注册表检测 Microsoft Office 时出错: {e}")

        return None

    @classmethod
    def _detect_wps_from_registry(cls) -> Optional[Dict[str, Any]]:
        """从注册表检测 WPS Office"""
        try:
            import winreg

            # 尝试打开 WPS 注册表项
            key_paths = [
                r"SOFTWARE\Kingsoft\Office",
                r"SOFTWARE\WOW6432Node\Kingsoft\Office",
            ]

            for key_path in key_paths:
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    path, _ = winreg.QueryValueEx(key, "InstallPath")
                    winreg.CloseKey(key)

                    if path and os.path.exists(path):
                        return {
                            'type': 'wps',
                            'name': 'WPS Office',
                            'path': path,
                            'exe_path': path,
                            'version': '未知版本'
                        }
                except:
                    continue

        except Exception as e:
            logger.debug(f"从注册表检测 WPS Office 时出错: {e}")

        return None

    @classmethod
    def _get_file_version(cls, file_path: str) -> Optional[str]:
        """获取文件版本信息"""
        try:
            import win32api
            info = win32api.GetFileVersionInfo(file_path, '\\')
            ms = info['FileVersionMS']
            ls = info['FileVersionLS']
            version = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
            return version
        except:
            return None

    @classmethod
    def is_office_available(cls) -> bool:
        """检查是否有可用的 Office 套件"""
        return cls.detect_office_suite() is not None

    @classmethod
    def get_office_info(cls) -> str:
        """获取 Office 信息字符串"""
        office = cls.detect_office_suite()
        if office:
            return f"{office['name']} ({office['version']})"
        return "未安装 Office 套件"
