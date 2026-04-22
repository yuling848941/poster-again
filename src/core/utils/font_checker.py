"""
字体检查器 - 检测系统已安装字体和 PPTX 模板使用的字体
"""

import os
import logging
import zipfile
import xml.etree.ElementTree as ET
import re
from typing import Set, List, Dict, Optional
from pathlib import Path

logger = logging.getLogger(__name__)


class FontChecker:
    """检查 PPTX 模板使用的字体是否已安装到系统"""

    # Windows 字体目录
    SYSTEM_FONTS_DIR = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')

    # 系统默认字体（主题字体，通常不需要提醒）
    SYSTEM_THEME_FONTS = {
        'arial', 'times new roman', 'calibri', 'cambria', 'consolas',
        'georgia', 'impact', 'comic sans ms', 'trebuchet ms', 'verdana',
        'tahoma', 'palatino linotype', 'book antiqua', 'garamond',
        'courier new', 'lucida console', 'microsoft sans serif',
        # 中日韩字体（系统默认）
        'ms pgothic', 'ms ui gothic', 'yu gothic', 'meiryo',
        'simhei', 'simsun', 'fangsong', 'kaiti', 'microsoft yahei',
        'malgun gothic', 'batang', 'dotum', 'gulim',
        '新細明體', '細明體', '標楷體',
        # 其他系统字体
        'dengxian', 'dengxian light', 'yu gothic light', '맑은 고딕',
        # 东南亚字体（Windows 自带）
        'angsana new', 'cordia new', 'daunpenh', 'dokchampa',
        'estrangelo edessa', 'gautami', 'iskoola pota', 'kalinga',
        'kartika', 'latha', 'mangal', 'microsoft himalaya',
        'microsoft yi baiti', 'moolboran', 'mv boli', 'nyala',
        'plantagenet cherokee', 'raavi', 'shruti', 'tunga', 'vrinda',
        'euphemia', 'mongolian baiti', 'microsoft uighur',
        # 微软雅黑变体
        'microsoft yahei & microsoft yahei ui',
        'microsoft yahei bold & microsoft yahei ui bold',
        'microsoft yahei light & microsoft yahei ui light',
    }

    # 常见中文字体名称映射（包含可能的变体）
    COMMON_CHINESE_FONTS = {
        '微软雅黑': ['微软雅黑', 'Microsoft YaHei', 'yahei'],
        '宋体': ['宋体', 'SimSun', 'sunsong'],
        '黑体': ['黑体', 'SimHei'],
        '楷体': ['楷体', 'KaiTi', 'STKaiti'],
        '仿宋': ['仿宋', 'FangSong', 'STFangSong'],
        '华文黑体': ['华文黑体', 'STHeiti'],
        '华文宋体': ['华文宋体', 'STSong'],
        '方正': ['方正', 'FZ'],
        '汉仪': ['汉仪', 'HY'],
        '华康': ['华康', 'DF'],
    }

    def __init__(self):
        self._installed_fonts: Optional[Set[str]] = None

    def get_installed_fonts(self) -> Set[str]:
        """获取系统已安装的字体名称集合"""
        if self._installed_fonts is not None:
            return self._installed_fonts

        fonts = set()

        try:
            # 方法 1: 从注册表读取（最准确）
            try:
                import winreg
                key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
                )
                i = 0
                while True:
                    try:
                        value = winreg.EnumValue(key, i)
                        name = value[0]  # 字体名称（如 "Arial (TrueType)"）
                        # 提取字体名称（去掉文件扩展名和 (TrueType) 等后缀）
                        font_name = name.split('(')[0].strip().lower()
                        fonts.add(font_name)
                        fonts.add(name.lower())
                        i += 1
                    except OSError:
                        break
                winreg.CloseKey(key)
            except Exception:
                pass

            # 方法 2: 从字体文件目录读取文件名
            if os.path.exists(self.SYSTEM_FONTS_DIR):
                for filename in os.listdir(self.SYSTEM_FONTS_DIR):
                    if filename.lower().endswith(('.ttf', '.otf', '.ttc')):
                        # 使用文件名作为字体名称的参考
                        fonts.add(filename.lower())
                        # 添加常见字体文件名映射
                        font_file_map = {
                            'msyh.ttc': '微软雅黑',
                            'msyhbd.ttc': '微软雅黑 bold',
                            'msyhl.ttc': '微软雅黑 light',
                            'simsun.ttc': '宋体',
                            'simhei.ttf': '黑体',
                            'kaiti.ttf': '楷体',
                            'fangsong.ttf': '仿宋',
                        }
                        if filename.lower() in font_file_map:
                            fonts.add(font_file_map[filename.lower()].lower())

        except Exception as e:
            logger.warning(f"获取系统字体失败：{e}")

        self._installed_fonts = fonts
        return fonts

    def get_fonts_from_pptx(self, pptx_path: str) -> Dict[str, List[str]]:
        """
        从 PPTX 文件中提取使用的字体信息

        Returns:
            Dict: {字体名称：[使用位置列表]}
        """
        fonts = {}

        try:
            with zipfile.ZipFile(pptx_path, 'r') as zf:
                # 检查的文件列表
                files_to_check = [
                    'ppt/theme/theme1.xml',
                    'ppt/slideMasters/slideMaster1.xml',
                ]

                # 添加所有幻灯片文件
                for name in zf.namelist():
                    if name.startswith('ppt/slides/') and name.endswith('.xml'):
                        files_to_check.append(name)

                for file_path in files_to_check:
                    try:
                        data = zf.read(file_path)
                        content = data.decode('utf-8')
                        slide_name = file_path.split('/')[-1]

                        # 提取 typeface 属性中的字体
                        for match in re.findall(r'typeface="([^"]+)"', content):
                            if match and not match.startswith('+'):  # 跳过主题字体引用
                                if match not in fonts:
                                    fonts[match] = []
                                fonts[match].append(slide_name)

                        # 提取 font-family CSS 属性
                        for match in re.findall(r'font-family:[\"\']?([^;\"]+)', content):
                            font_name = match.strip()
                            if font_name:
                                if font_name not in fonts:
                                    fonts[font_name] = []
                                fonts[font_name].append(slide_name)

                    except Exception:
                        continue

        except Exception as e:
            logger.error(f"读取 PPTX 文件失败：{e}")

        return fonts

    def check_template_fonts(self, pptx_path: str) -> Dict:
        """
        检查 PPTX 模板使用的字体是否已安装

        Args:
            pptx_path: PPTX 文件路径

        Returns:
            Dict: 检查结果，包含：
                - all_installed: bool - 所有字体是否已安装
                - missing_fonts: List[str] - 缺失的字体列表
                - installed_fonts: List[str] - 已安装的字体列表
                - font_details: Dict - 每个字体的详细信息
        """
        template_fonts = self.get_fonts_from_pptx(pptx_path)
        installed = self.get_installed_fonts()

        result = {
            'all_installed': True,
            'missing_fonts': [],
            'installed_fonts': [],
            'font_details': {}
        }

        for font_name, locations in template_fonts.items():
            # 跳过系统主题字体（不提醒）
            if font_name.lower() in self.SYSTEM_THEME_FONTS:
                continue

            # 检查字体是否已安装（不区分大小写）
            font_lower = font_name.lower()
            is_installed = font_lower in installed or font_name in installed

            # 检查常见字体名称变体
            if not is_installed:
                for common_name, variants in self.COMMON_CHINESE_FONTS.items():
                    if any(v.lower() in font_lower or font_lower in v.lower()
                           for v in variants):
                        # 如果包含常见字体名称，认为已安装
                        if any(v.lower() in installed for v in variants):
                            is_installed = True
                            break

            detail = {
                'name': font_name,
                'installed': is_installed,
                'locations': locations
            }

            if is_installed:
                result['installed_fonts'].append(font_name)
            else:
                # 再次检查：是否是已知系统字体的中文名
                is_system_font = False
                sys_font_names = {
                    '微软雅黑': ['microsoft yahei'],
                    '宋体': ['simsun'],
                    '黑体': ['simhei'],
                    '楷体': ['kaiti', 'stkaiti'],
                    '仿宋': ['fangsong', 'stfangsong'],
                    '新宋体': ['nsimsun'],
                }
                for cn_name, en_names in sys_font_names.items():
                    if font_name == cn_name:
                        # 检查对应的英文名是否在系统中（支持部分匹配）
                        if any(any(en in installed_font for en in en_names)
                               for installed_font in installed):
                            is_system_font = True
                            is_installed = True
                            break

                if not is_system_font:
                    result['all_installed'] = False
                    result['missing_fonts'].append(font_name)

            result['font_details'][font_name] = detail

        return result

    def get_missing_fonts_message(self, check_result: Dict) -> str:
        """生成缺失字体的用户友好提示信息"""
        if check_result['all_installed']:
            return "模板使用的所有字体已安装。"

        missing = check_result['missing_fonts']
        if not missing:
            return "未检测到模板使用的自定义字体。"

        messages = []
        messages.append("⚠️ 模板使用了以下系统未安装的字体：")
        for font in missing:
            locations = check_result['font_details'][font]['locations']
            messages.append(f"  • {font}（用于：{', '.join(locations)}）")

        messages.append("")
        messages.append("建议操作：")
        messages.append("  1. 安装缺失的字体文件后重新运行程序")
        messages.append("  2. 或在 PowerPoint 中将模板字体替换为系统已安装字体")
        messages.append("  3. 或导出图片时，缺失字体的文本将使用微软雅黑显示")

        return "\n".join(messages)
