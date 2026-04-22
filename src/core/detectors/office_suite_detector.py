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
        # Office 365 / 2016+ Click-to-Run
        r"C:\Program Files\Microsoft Office\root\Office16",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16",
        r"C:\Program Files\Microsoft Office\root\Office15",
        r"C:\Program Files (x86)\Microsoft Office\root\Office15",
        # Office 2016/2019/2021 永久版本
        r"C:\Program Files\Microsoft Office\Office16",
        r"C:\Program Files (x86)\Microsoft Office\Office16",
        r"C:\Program Files\Microsoft Office\Office15",
        r"C:\Program Files (x86)\Microsoft Office\Office15",
        # Office 365 用户安装路径
        os.path.expanduser(r"~\AppData\Local\Microsoft\Office\root\Office16"),
        os.path.expanduser(r"~\AppData\Local\Microsoft\Office\root\Office15"),
        # 其他可能的安装路径
        r"D:\Program Files\Microsoft Office\root\Office16",
        r"D:\Program Files (x86)\Microsoft Office\root\Office16",
        r"E:\Program Files\Microsoft Office\root\Office16",
        r"E:\Program Files (x86)\Microsoft Office\root\Office16",
    ]

    # WPS Office 可能的安装路径
    WPS_PATHS = [
        r"C:\Program Files\Kingsoft\WPS Office",
        r"C:\Program Files (x86)\Kingsoft\WPS Office",
        r"C:\Program Files\WPS Office",
        r"C:\Program Files (x86)\WPS Office",
        # WPS 个人版可能路径
        os.path.expanduser(r"~\AppData\Local\Kingsoft\WPS Office"),
        # 其他盘符安装路径
        r"D:\Program Files\Kingsoft\WPS Office",
        r"D:\Program Files (x86)\Kingsoft\WPS Office",
        r"D:\Program Files\WPS Office",
        r"E:\Program Files\Kingsoft\WPS Office",
        r"E:\Program Files\WPS Office",
    ]

    @classmethod
    def detect_office_suite(cls, use_cache: bool = True) -> Optional[Dict[str, Any]]:
        """
        检测系统中安装的 Office 套件

        Args:
            use_cache: 是否使用缓存（如果之前检测成功过）

        Returns:
            Optional[Dict]: 检测到的 Office 套件信息，如果没有安装则返回 None
            {
                'type': 'microsoft' | 'wps',
                'name': str,
                'path': str,
                'version': str,
                'progid': str  # COM ProgID
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

        # 文件系统和注册表检测都失败，尝试 COM 接口直接检测
        logger.debug("文件系统检测失败，尝试 COM 接口检测...")
        com_office = cls._detect_via_com()
        if com_office:
            logger.info(f"通过 COM 接口检测到 Office: {com_office['name']}")
            return com_office

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
                except (OSError, PermissionError):
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
                except (OSError, PermissionError):
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
        except Exception as e:
            logger.debug(f"获取文件版本失败：{e}")
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

    @classmethod
    def test_com_availability(cls, office_type: str = "microsoft") -> bool:
        """
        测试 COM 接口可用性

        Args:
            office_type: Office 类型，"microsoft" 或 "wps"

        Returns:
            bool: COM 接口是否可用
        """
        try:
            import win32com.client

            if office_type == "microsoft":
                # 尝试创建 PowerPoint 实例
                try:
                    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                    # 检查是否成功创建
                    if powerpoint:
                        powerpoint.Quit()
                        return True
                except Exception:
                    pass

                # 尝试获取活动实例
                try:
                    powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
                    if powerpoint:
                        return True
                except Exception:
                    pass

            elif office_type == "wps":
                # 尝试创建 WPS 实例
                wps_progids = ["Kwpp.Application", "WPP.Application"]
                for progid in wps_progids:
                    try:
                        wps = win32com.client.Dispatch(progid)
                        if wps:
                            wps.Quit()
                            return True
                    except Exception:
                        pass

                    try:
                        wps = win32com.client.GetActiveObject(progid)
                        if wps:
                            return True
                    except Exception:
                        pass

            return False

        except Exception:
            return False

    @classmethod
    def _detect_via_com(cls) -> Optional[Dict[str, Any]]:
        """
        通过 COM 接口检测 Office 套件（后备检测方式）

        Returns:
            Optional[Dict]: 检测到的 Office 套件信息
        """
        try:
            import win32com.client

            # 优先尝试 Microsoft Office
            try:
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                version = powerpoint.Version
                powerpoint.Quit()
                return {
                    'type': 'microsoft',
                    'name': 'Microsoft Office PowerPoint (COM 检测)',
                    'path': 'COM 接口',
                    'exe_path': 'POWERPNT.EXE',
                    'version': version or '未知版本',
                    'progid': 'PowerPoint.Application'
                }
            except Exception:
                pass

            # 尝试获取活动 PowerPoint 实例
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
                version = powerpoint.Version
                return {
                    'type': 'microsoft',
                    'name': 'Microsoft Office PowerPoint (活动实例)',
                    'path': 'COM 接口',
                    'exe_path': 'POWERPNT.EXE',
                    'version': version or '未知版本',
                    'progid': 'PowerPoint.Application'
                }
            except Exception:
                pass

            # 尝试 WPS Office
            wps_progids = ["Kwpp.Application", "WPP.Application"]
            for progid in wps_progids:
                try:
                    wps = win32com.client.Dispatch(progid)
                    wps.Quit()
                    return {
                        'type': 'wps',
                        'name': 'WPS Office (COM 检测)',
                        'path': 'COM 接口',
                        'exe_path': 'WPS.exe',
                        'version': '未知版本',
                        'progid': progid
                    }
                except Exception:
                    pass

                try:
                    wps = win32com.client.GetActiveObject(progid)
                    return {
                        'type': 'wps',
                        'name': 'WPS Office (活动实例)',
                        'path': 'COM 接口',
                        'exe_path': 'WPS.exe',
                        'version': '未知版本',
                        'progid': progid
                    }
                except Exception:
                    pass

            return None

        except Exception as e:
            logger.debug(f"通过 COM 接口检测 Office 时出错：{e}")
            return None
