"""
COM 接口图片导出器
使用 Microsoft Office 或 WPS Office 导出 PPT 为图片
"""

import os
import logging
import subprocess
import tempfile
from typing import Optional, List
from pathlib import Path

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
            logger.info(f"使用 Office 套件: {self.office_info['name']}")
            return True
        logger.warning("未检测到 Office 套件，图片导出功能将不可用")
        return False

    def is_available(self) -> bool:
        """检查导出器是否可用"""
        return self.office_info is not None

    def export_to_images(self, pptx_path: str, output_dir: str,
                        image_format: str = "PNG", quality: str = "原始大小") -> List[str]:
        """
        将 PPTX 导出为图片

        Args:
            pptx_path: PPTX 文件路径
            output_dir: 输出目录
            image_format: 图片格式 (PNG, JPG)
            quality: 图片质量

        Returns:
            List[str]: 生成的图片路径列表
        """
        if not self.is_available():
            raise RuntimeError("未安装 Microsoft Office 或 WPS Office，无法导出图片")

        if not os.path.exists(pptx_path):
            raise FileNotFoundError(f"文件不存在: {pptx_path}")

        # 确保使用绝对路径
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        # 使用 python-pptx 读取幻灯片实际尺寸
        from pptx import Presentation
        prs = Presentation(pptx_path)
        # 幻灯片尺寸（EMU）
        slide_width_emu = prs.slide_width
        slide_height_emu = prs.slide_height
        # 转换为像素（96 DPI）
        slide_width_px = int(slide_width_emu / 914400 * 96)
        slide_height_px = int(slide_height_emu / 914400 * 96)

        logger.info(f"幻灯片尺寸: {slide_width_px}x{slide_height_px} 像素")

        if self.office_info['type'] == 'microsoft':
            return self._export_with_microsoft(pptx_path, output_dir, image_format, quality,
                                               slide_width_px, slide_height_px)
        else:
            return self._export_with_wps(pptx_path, output_dir, image_format, quality,
                                         slide_width_px, slide_height_px)

    def _export_with_microsoft(self, pptx_path: str, output_dir: str,
                               image_format: str, quality: str,
                               slide_width: int, slide_height: int) -> List[str]:
        """使用 Microsoft Office 导出"""
        try:
            import win32com.client

            # 创建 PowerPoint 应用
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # 某些版本的 Office 不允许隐藏窗口，所以不设置 Visible
            # powerpoint.Visible = False

            try:
                # 打开演示文稿
                presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)

                # 设置导出格式
                # 注意：某些 Office 版本可能不支持 PNG，使用 JPG 更兼容
                format_map = {
                    'PNG': 'PNG',
                    'JPG': 'JPG',
                    'JPEG': 'JPG',
                    'BMP': 'BMP',
                    'GIF': 'GIF'
                }
                export_format = format_map.get(image_format.upper(), 'JPG')

                # 设置分辨率
                quality_map = {
                    '原始大小': 1.0,
                    '2倍': 2.0,
                    '3倍': 3.0
                }
                scale = quality_map.get(quality, 1.0)

                generated_files = []

                # 导出每一页
                for i in range(1, presentation.Slides.Count + 1):
                    slide = presentation.Slides(i)

                    # 生成文件名
                    filename = f"slide_{i}.{image_format.lower()}"
                    output_path = os.path.join(output_dir, filename)

                    # 导出幻灯片
                    # PowerPoint Export 方法参数：FileName, FilterName, ScaleWidth, ScaleHeight
                    # 使用实际幻灯片尺寸，根据质量设置缩放比例
                    width = int(slide_width * scale)
                    height = int(slide_height * scale)
                    slide.Export(output_path, export_format, width, height)
                    generated_files.append(output_path)
                    logger.info(f"导出图片: {output_path} ({width}x{height})")

                presentation.Close()
                return generated_files

            finally:
                powerpoint.Quit()

        except Exception as e:
            logger.error(f"使用 Microsoft Office 导出图片时出错: {e}")
            raise

    def _export_with_wps(self, pptx_path: str, output_dir: str,
                        image_format: str, quality: str,
                        slide_width: int, slide_height: int) -> List[str]:
        """使用 WPS Office 导出"""
        try:
            import win32com.client

            # 尝试使用 WPS 的 COM 接口
            # WPS 通常兼容 Microsoft Office 的 COM 接口
            try:
                wps = win32com.client.Dispatch("Kwpp.Application")
            except:
                # 尝试其他 WPS COM 名称
                wps = win32com.client.Dispatch("WPP.Application")

            wps.Visible = False

            try:
                # 打开演示文稿
                presentation = wps.Presentations.Open(pptx_path, WithWindow=False)

                # 设置导出格式
                format_map = {
                    'PNG': 'PNG',
                    'JPG': 'JPG',
                    'JPEG': 'JPG',
                    'BMP': 'BMP',
                    'GIF': 'GIF'
                }
                export_format = format_map.get(image_format.upper(), 'JPG')

                quality_map = {
                    '原始大小': 1.0,
                    '2倍': 2.0,
                    '3倍': 3.0
                }
                scale = quality_map.get(quality, 1.0)

                generated_files = []

                # 导出每一页
                for i in range(1, presentation.Slides.Count + 1):
                    slide = presentation.Slides(i)

                    filename = f"slide_{i}.{image_format.lower()}"
                    output_path = os.path.join(output_dir, filename)

                    # PowerPoint Export 方法参数：FileName, FilterName, ScaleWidth, ScaleHeight
                    # 使用实际幻灯片尺寸，根据质量设置缩放比例
                    width = int(slide_width * scale)
                    height = int(slide_height * scale)
                    slide.Export(output_path, export_format, width, height)
                    generated_files.append(output_path)
                    logger.info(f"导出图片: {output_path} ({width}x{height})")

                presentation.Close()
                return generated_files

            finally:
                wps.Quit()

        except Exception as e:
            logger.error(f"使用 WPS Office 导出图片时出错: {e}")
            # 如果 WPS COM 接口失败，尝试使用 Microsoft Office 的方法
            logger.info("尝试使用 Microsoft Office 兼容模式...")
            return self._export_with_microsoft(pptx_path, output_dir, image_format, quality,
                                               slide_width, slide_height)

    def get_office_info(self) -> str:
        """获取当前使用的 Office 信息"""
        if self.office_info:
            return f"{self.office_info['name']} ({self.office_info['version']})"
        return "未安装 Office 套件"
