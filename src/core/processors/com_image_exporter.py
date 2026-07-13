"""
COM 接口图片导出器
使用 Microsoft Office 或 WPS Office 导出 PPT 为图片
"""

import os
import logging
import tempfile
import shutil
import time
from typing import Optional, List
from pathlib import Path

try:
    import pywintypes
except ImportError:
    pywintypes = None

try:
    from src.core.detectors.office_suite_detector import OfficeSuiteDetector
except ImportError:
    from core.detectors.office_suite_detector import OfficeSuiteDetector

logger = logging.getLogger(__name__)


class COMImageExporter:
    """
    使用 Office COM 接口导出 PPT 为图片

    优先使用 Microsoft Office，如果没有则使用 WPS Office
    """

    def __init__(self):
        self.office_info = None
        self._check_office()

    def _check_office(self) -> bool:
        """检查是否有可用的 Office 套件"""
        self.office_info = OfficeSuiteDetector.detect_office_suite()
        if self.office_info:
            logger.info(f"使用 Office 套件：{self.office_info['name']}")
            return True
        logger.warning("未检测到 Office 套件，图片导出功能将不可用")
        return False

    def is_available(self) -> bool:
        """检查导出器是否可用"""
        return self.office_info is not None

    def export_to_images(self, pptx_path: str, output_dir: str,
                        image_format: str = "PNG", quality: float = 1.0) -> List[str]:
        """
        将 PPTX 导出为图片

        Args:
            pptx_path: PPTX 文件路径
            output_dir: 输出目录
            image_format: 图片格式 (PNG, JPG)
            quality: 缩放倍数 (1.0=原始, 2.0=增强, 3.0=高质量)

        Returns:
            List[str]: 生成的图片路径列表
        """
        if not self.is_available():
            raise RuntimeError("未安装 Microsoft Office 或 WPS Office，无法导出图片")

        if not os.path.exists(pptx_path):
            raise FileNotFoundError(f"文件不存在：{pptx_path}")

        # 确保使用绝对路径
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        # 获取幻灯片尺寸（用于日志）
        from pptx import Presentation
        prs = Presentation(pptx_path)
        slide_width_emu = prs.slide_width
        slide_height_emu = prs.slide_height
        slide_width_px = int(slide_width_emu / 914400 * 96)
        slide_height_px = int(slide_height_emu / 914400 * 96)

        logger.info(f"幻灯片尺寸：{slide_width_px}x{slide_height_px} 像素")

        if self.office_info['type'] == 'microsoft':
            return self._export_with_microsoft(pptx_path, output_dir, image_format, quality,
                                               slide_width_px, slide_height_px)
        else:
            return self._export_with_wps(pptx_path, output_dir, image_format, quality,
                                         slide_width_px, slide_height_px)

    def _export_with_microsoft(self, pptx_path: str, output_dir: str,
                               image_format: str, quality: float,
                               slide_width: int, slide_height: int) -> List[str]:
        """
        使用 Microsoft Office 导出图片

        实例获取策略：优先复用用户已打开的 PowerPoint（GetActiveObject），
        避免导出结束后 Quit() 误杀用户进程。只有用户没开 PowerPoint 时，
        才 Dispatch 启动新实例并由我们负责清理（we_started=True）。

        Args:
            quality: 缩放倍数 (1.0/2.0/3.0)，作用于 Export 的 width/height 参数
        """
        powerpoint = None
        we_started = False  # 标记我们是否新建了这个进程（决定能否 Quit）
        try:
            import win32com.client

            # 优先复用用户已打开的 PowerPoint 实例，避免误杀
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
                logger.debug("复用已运行的 PowerPoint 实例")
            except Exception:
                # 用户没有打开 PowerPoint，启动新实例（这种情况我们拥有它，可负责清理）
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                we_started = True
                logger.debug("启动新的 PowerPoint 实例（我们将负责清理）")

            try:
                # 打开演示文稿（WithWindow=False 避免弹出窗口）
                presentation = powerpoint.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)

                # 导出图片
                format_map = {
                    'PNG': 'PNG',
                    'JPG': 'JPG',
                    'JPEG': 'JPG',
                    'BMP': 'BMP',
                    'GIF': 'GIF'
                }
                export_format = format_map.get(image_format.upper(), 'JPG')

                generated_files = []

                # 按缩放倍数计算导出分辨率（单位：像素）
                export_width = int(slide_width * quality)
                export_height = int(slide_height * quality)

                for i in range(1, presentation.Slides.Count + 1):
                    slide = presentation.Slides(i)
                    filename = f"slide_{i}.{image_format.lower()}"
                    output_path = os.path.join(output_dir, filename)
                    slide.Export(output_path, export_format, export_width, export_height)
                    generated_files.append(output_path)
                    logger.info(f"导出图片：{output_path} ({export_width}x{export_height})")

                presentation.Close()
                return generated_files

            finally:
                # 只关闭我们自己打开的演示文稿（已在上面 presentation.Close() 完成）
                # 绝不关闭用户已打开的 PowerPoint 实例
                if powerpoint is not None:
                    if we_started:
                        # 我们启动的实例：若没有其他文档残留，才退出进程
                        try:
                            if powerpoint.Presentations.Count == 0:
                                powerpoint.Quit()
                                logger.debug("退出我们启动的 PowerPoint 实例")
                        except Exception as e:
                            logger.debug(f"退出 PowerPoint 实例失败（可忽略）: {e}")
                    else:
                        # 复用的用户实例：只释放 COM 引用，绝不 Quit
                        try:
                            del powerpoint
                        except Exception:
                            pass

        except Exception as e:
            logger.error(f"使用 Microsoft Office 导出图片时出错：{e}")
            raise

    def _export_with_wps(self, pptx_path: str, output_dir: str,
                        image_format: str, quality: float,
                        slide_width: int, slide_height: int) -> List[str]:
        """
        使用 WPS Office 导出

        Args:
            quality: 缩放倍数 (1.0/2.0/3.0)，作用于 Export 的 width/height 参数
        """
        try:
            import win32com.client

            wps = None
            presentation = None

            # 尝试获取已有的 WPS 实例
            try:
                wps = win32com.client.GetActiveObject("Kwpp.Application")
            except (pywintypes.com_error, OSError) if pywintypes else Exception:
                try:
                    wps = win32com.client.GetActiveObject("WPP.Application")
                except (pywintypes.com_error, OSError) if pywintypes else Exception:
                    try:
                        wps = win32com.client.Dispatch("Kwpp.Application")
                    except (pywintypes.com_error, OSError) if pywintypes else Exception:
                        wps = win32com.client.Dispatch("WPP.Application")

            try:
                # 打开演示文稿（WithWindow=False 避免弹出窗口）
                presentation = wps.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)

                # 设置导出格式
                format_map = {
                    'PNG': 'PNG',
                    'JPG': 'JPG',
                    'JPEG': 'JPG',
                    'BMP': 'BMP',
                    'GIF': 'GIF'
                }
                export_format = format_map.get(image_format.upper(), 'JPG')

                # quality 即缩放倍数（1.0/2.0/3.0），直接使用
                generated_files = []

                # 导出每一页
                for i in range(1, presentation.Slides.Count + 1):
                    slide = presentation.Slides(i)

                    filename = f"slide_{i}.{image_format.lower()}"
                    output_path = os.path.join(output_dir, filename)

                    # 导出幻灯片
                    width = int(slide_width * quality)
                    height = int(slide_height * quality)
                    slide.Export(output_path, export_format, width, height)
                    generated_files.append(output_path)
                    logger.info(f"导出图片：{output_path} ({width}x{height})")

                presentation.Close()
                return generated_files

            finally:
                # 释放 COM 引用，但不调用 Quit() 以免影响用户已打开的 WPS
                if presentation is not None:
                    try:
                        del presentation
                    except Exception:
                        pass
                if wps is not None:
                    try:
                        del wps
                    except Exception:
                        pass

        except Exception as e:
            logger.error(f"使用 WPS Office 导出图片时出错：{e}")
            # 如果 WPS COM 接口失败，尝试使用 Microsoft Office 的方法
            logger.info("尝试使用 Microsoft Office 兼容模式...")
            return self._export_with_microsoft(pptx_path, output_dir, image_format, quality,
                                               slide_width, slide_height)

    def get_office_info(self) -> str:
        """获取当前使用的 Office 信息"""
        if self.office_info:
            return f"{self.office_info['name']} ({self.office_info['version']})"
        return "未安装 Office 套件"
