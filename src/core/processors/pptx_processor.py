"""
PPTX 模板处理器 - 纯 Python 实现
基于 python-pptx 库，无需 Microsoft Office 或 WPS

支持占位符识别方式（按优先级）：
1. Alt Text（替代文本）- 推荐，完全隐藏
2. 形状名称（Shape Name）- 备选
3. 文本内容标记（如 {{字段名}}）- 兼容方案
"""

import os
import logging
from typing import Dict, List, Optional, Tuple, Any
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

from render_core.pillow_image_exporter import PillowImageExporter
from render_core.format_inheritance import load_theme
from render_core.personalize import snapshot_slides, restore_slides
from render_core import font_registry

try:
    from src.core.utils.pptx_fast_save import save_pptx
except ImportError:
    from core.utils.pptx_fast_save import save_pptx

try:
    from src.core.interfaces.template_processor import (
        ITemplateProcessor,
        ConnectionError,
        TemplateLoadError,
        PlaceholderError,
        SaveError,
    )
except ImportError:
    from core.interfaces.template_processor import (
        ITemplateProcessor,
        ConnectionError,
        TemplateLoadError,
        PlaceholderError,
        SaveError,
    )

logger = logging.getLogger(__name__)


class PPTXProcessor(ITemplateProcessor):
    """
    基于 python-pptx 的模板处理器
    
    无需安装 Microsoft Office 或 WPS，纯 Python 实现
    """
    
    # 占位符前缀
    PLACEHOLDER_PREFIX = "ph_"
    # 文本标记格式
    TEXT_MARKER_START = "{{"
    TEXT_MARKER_END = "}}"
    
    def __init__(self):
        """初始化 PPTX 处理器"""
        super().__init__()
        self.presentation: Optional[Presentation] = None
        self.template_path: Optional[str] = None
        # 批量渲染的复用缓存（生命周期 ≈ 一个批次；批次结束后处理器会被重建）：
        # 引擎按策略缓存，避免每行重新检测 Office；渲染上下文按模板缓存
        self._engine_cache: Dict[str, tuple] = {}
        self._engine_strategy: Optional[str] = None     # 本实例解析出的引擎策略（缓存）
        self._python_exporter: Optional[PillowImageExporter] = None
        self._render_theme: Optional[tuple] = None      # (template_path, theme)
        self._precheck_done: Optional[str] = None       # 已跑过预检的模板路径
        self._slide_xml = None                          # 幻灯片原始 XML 快照（恢复用）
        self.last_export_error: Optional[str] = None    # 最近一次导出失败的原因（供上层透出）
        
    def connect_to_application(self) -> bool:
        """
        连接到应用程序
        
        对于 python-pptx，此方法始终返回 True
        因为不需要连接外部应用程序
        
        Returns:
            bool: 始终返回 True
        """
        self.is_connected = True
        logger.info("PPTX 处理器初始化完成（无需外部应用程序）")
        return True
    
    def load_template(self, template_path: str) -> bool:
        """
        加载 PPT 模板文件
        
        Args:
            template_path: 模板文件路径
            
        Returns:
            bool: 加载是否成功
        """
        if not self.validate_template_path(template_path):
            raise TemplateLoadError(f"无效的模板文件路径: {template_path}")
        
        try:
            # 确保路径是绝对路径
            template_path = os.path.abspath(template_path)
            
            # 使用 python-pptx 加载演示文稿
            self.presentation = Presentation(template_path)
            self.template_path = template_path
            # 幻灯片原始 XML 快照：批量生成逐行「替换 -> 恢复」复用同一 Presentation，
            # 不必每行重新解压解析整个包（也让渲染器的图片缓存跨行生效）
            self._slide_xml = snapshot_slides(self.presentation)
            # 模板变了，渲染上下文缓存作废
            self._render_theme = None
            self._precheck_done = None

            logger.info(f"成功加载模板: {template_path}")
            logger.info(f"幻灯片数量: {len(self.presentation.slides)}")
            return True
            
        except Exception as e:
            logger.error(f"加载模板失败: {str(e)}")
            raise TemplateLoadError(f"无法加载模板文件: {str(e)}")

    def restore_template(self) -> bool:
        """把当前 Presentation 恢复到加载时的模板原状（撤销占位符替换）。

        批量生成逐行复用同一 Presentation：替换 -> 导出/保存 -> 恢复。
        与重新 load_template 等价（占位符替换只动幻灯片 spTree 内的文本），
        但无需重新解压解析整个包，渲染器的字体/图片缓存也得以跨行保留。
        """
        if not self.presentation or self._slide_xml is None:
            return False
        restore_slides(self.presentation, self._slide_xml)
        return True

    def _get_alt_text(self, shape) -> str:
        """
        获取形状的Alt Text（替代文本）
        
        支持多种方式读取，包括WPS设置的Alt Text
        
        Args:
            shape: PPTX形状对象
            
        Returns:
            str: Alt Text内容
        """
        # 方法1: 使用标准alt_text属性
        alt_text = getattr(shape, 'alt_text', '') or ''
        if alt_text:
            return alt_text
        
        # 方法2: 从XML中读取descr属性（WPS使用这种方式）
        try:
            from pptx.oxml.ns import qn
            nvSpPr = shape._element.find(qn('p:nvSpPr'))
            if nvSpPr is not None:
                cNvPr = nvSpPr.find(qn('p:cNvPr'))
                if cNvPr is not None:
                    descr = cNvPr.get('descr', '')
                    if descr:
                        return descr
        except Exception as e:
            logger.debug(f"获取形状属性失败：{e}")
            pass
        
        return ''

    def _iterate_all_shapes(self, shapes):
        """
        递归遍历所有形状（包括组合形状内部的子形状）

        Args:
            shapes: 形状集合

        Yields:
            形状对象
        """
        for shape in shapes:
            yield shape

            # 如果是组合形状，递归遍历其子形状（支持嵌套组合）
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    for subshape in self._iterate_all_shapes(shape.shapes):
                        yield subshape
                except Exception as e:
                    logger.debug(f"遍历组合形状失败：{e}")

    def find_placeholders(self) -> List[str]:
        """
        查找所有占位符
        
        按以下优先级识别占位符：
        1. Alt Text（替代文本）以 ph_ 开头
        2. 形状名称以 ph_ 开头
        3. 文本内容中的 {{字段名}} 格式
        
        Returns:
            List[str]: 占位符列表（去重）
        """
        if not self.presentation:
            logger.error("未加载演示文稿")
            return []
        
        placeholders = set()
        
        try:
            for slide_idx, slide in enumerate(self.presentation.slides, 1):
                for shape in self._iterate_all_shapes(slide.shapes):
                    # 1. 检查 Alt Text（支持WPS和PowerPoint）
                    alt_text = self._get_alt_text(shape)
                    if alt_text and alt_text.startswith(self.PLACEHOLDER_PREFIX):
                        placeholders.add(alt_text)
                        logger.debug(f"幻灯片 {slide_idx}: 从 Alt Text 发现占位符 '{alt_text}'")
                        continue
                    
                    # 2. 检查形状名称
                    shape_name = getattr(shape, 'name', '') or ''
                    if shape_name and shape_name.startswith(self.PLACEHOLDER_PREFIX):
                        placeholders.add(shape_name)
                        logger.debug(f"幻灯片 {slide_idx}: 从形状名称发现占位符 '{shape_name}'")
                        continue
                    
                    # 3. 检查文本内容中的标记
                    if shape.has_text_frame:
                        text = shape.text_frame.text
                        markers = self._extract_text_markers(text)
                        for marker in markers:
                            placeholders.add(f"{self.PLACEHOLDER_PREFIX}{marker}")
                            logger.debug(f"幻灯片 {slide_idx}: 从文本内容发现占位符 '{marker}'")
            
            result = sorted(list(placeholders))
            logger.info(f"共找到 {len(result)} 个占位符: {result}")
            return result
            
        except Exception as e:
            logger.error(f"查找占位符失败: {str(e)}")
            raise PlaceholderError(f"查找占位符时出错: {str(e)}")
    
    def _extract_text_markers(self, text: str) -> List[str]:
        """
        从文本中提取标记格式的占位符
        
        例如："姓名: {{name}}" -> ["name"]
        
        Args:
            text: 文本内容
            
        Returns:
            List[str]: 提取的占位符名称列表
        """
        markers = []
        start = self.TEXT_MARKER_START
        end = self.TEXT_MARKER_END
        
        idx = 0
        while True:
            start_idx = text.find(start, idx)
            if start_idx == -1:
                break
            end_idx = text.find(end, start_idx + len(start))
            if end_idx == -1:
                break
            
            marker = text[start_idx + len(start):end_idx].strip()
            if marker:
                markers.append(marker)
            idx = end_idx + len(end)
        
        return markers
    
    def get_placeholder_info(self) -> Dict[str, Dict[str, Any]]:
        """
        获取详细的占位符信息
        
        Returns:
            Dict: 占位符详细信息，包括位置、类型等
        """
        if not self.presentation:
            return {}
        
        info = {}
        
        try:
            for slide_idx, slide in enumerate(self.presentation.slides, 1):
                for shape in self._iterate_all_shapes(slide.shapes):
                    placeholder_key = None
                    source = None
                    
                    # 检查 Alt Text（支持WPS和PowerPoint）
                    alt_text = self._get_alt_text(shape)
                    if alt_text and alt_text.startswith(self.PLACEHOLDER_PREFIX):
                        placeholder_key = alt_text
                        source = 'alt_text'
                    
                    # 检查形状名称
                    elif shape.name and shape.name.startswith(self.PLACEHOLDER_PREFIX):
                        placeholder_key = shape.name
                        source = 'shape_name'
                    
                    if placeholder_key:
                        info[placeholder_key] = {
                            'slide': slide_idx,
                            'shape_name': shape.name,
                            'shape_type': str(shape.shape_type),
                            'source': source,
                            'has_text_frame': shape.has_text_frame,
                            'current_text': shape.text_frame.text if shape.has_text_frame else ''
                        }
            
            return info
            
        except Exception as e:
            logger.error(f"获取占位符信息失败: {str(e)}")
            return {}
    
    def replace_placeholders(self, data: Dict[str, str]) -> bool:
        """
        替换占位符内容
        
        支持三种占位符格式：
        1. Alt Text: ph_字段名
        2. 形状名称: ph_字段名
        3. 文本标记: {{字段名}}
        
        Args:
            data: 占位符和对应值的字典
            
        Returns:
            bool: 替换是否成功
        """
        if not self.presentation:
            logger.error("未加载演示文稿")
            return False
        
        try:
            replaced_count = 0
            
            for slide in self.presentation.slides:
                for shape in self._iterate_all_shapes(slide.shapes):
                    # 获取此形状对应的占位符键
                    placeholder_key = self._get_placeholder_key(shape)
                    
                    if placeholder_key and placeholder_key in data:
                        new_value = str(data[placeholder_key])
                        
                        if shape.has_text_frame:
                            self._replace_text_in_shape(shape, new_value)
                            replaced_count += 1
                            logger.debug(f"替换占位符 '{placeholder_key}' -> '{new_value}'")
                    
                    # 处理文本标记格式 {{字段名}}
                    elif shape.has_text_frame:
                        text_replaced = self._replace_text_markers(shape, data)
                        if text_replaced:
                            replaced_count += 1
            
            logger.info(f"成功替换 {replaced_count} 个占位符")
            return True
            
        except Exception as e:
            logger.error(f"替换占位符失败: {str(e)}")
            raise PlaceholderError(f"替换占位符时出错: {str(e)}")
    
    def _get_placeholder_key(self, shape) -> Optional[str]:
        """
        获取形状的占位符键
        
        Args:
            shape: PPT 形状对象
            
        Returns:
            Optional[str]: 占位符键，如果没有则返回 None
        """
        # 优先检查 Alt Text（支持WPS和PowerPoint）
        alt_text = self._get_alt_text(shape)
        if alt_text and alt_text.startswith(self.PLACEHOLDER_PREFIX):
            return alt_text
        
        # 其次检查形状名称
        shape_name = getattr(shape, 'name', '') or ''
        if shape_name and shape_name.startswith(self.PLACEHOLDER_PREFIX):
            return shape_name
        
        return None
    
    def _replace_text_in_shape(self, shape, new_text: str):
        """
        替换形状中的文本，保持格式

        策略：保留第一个段落的完整格式（字体、大小、颜色等），
        将新文本填入第一个段落，清空其他段落。

        Args:
            shape: PPT 形状对象
            new_text: 新文本内容
        """
        if not shape.has_text_frame:
            return

        text_frame = shape.text_frame

        if len(text_frame.paragraphs) == 0:
            return

        # 获取第一个段落（保留其格式）
        first_paragraph = text_frame.paragraphs[0]

        # 处理第一个段落的 runs
        if len(first_paragraph.runs) == 1:
            # 只有一个 run，直接替换文本
            first_paragraph.runs[0].text = new_text
        elif len(first_paragraph.runs) > 1:
            # 多个 runs，保留第一个 run 的格式，清空其他 runs
            first_paragraph.runs[0].text = new_text
            for run in first_paragraph.runs[1:]:
                run.text = ""
        else:
            # 没有 runs，创建一个新 run（这种情况很少见）
            first_paragraph.add_run().text = new_text

        # 清空其他段落的内容（保留段落结构但不显示）
        for i, paragraph in enumerate(text_frame.paragraphs):
            if i == 0:
                continue  # 跳过第一个段落

            # 清空该段落的所有 runs
            for run in paragraph.runs:
                run.text = ""
    
    def _replace_text_markers(self, shape, data: Dict[str, str]) -> bool:
        """
        替换文本中的标记格式占位符
        
        Args:
            shape: PPT 形状对象
            data: 数据字典
            
        Returns:
            bool: 是否进行了替换
        """
        if not shape.has_text_frame:
            return False
        
        replaced = False
        
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                original_text = run.text
                new_text = original_text
                
                # 替换所有 {{字段名}} 格式的标记
                for key, value in data.items():
                    # 支持 {{字段名}} 和 ph_字段名 两种格式
                    marker1 = f"{self.TEXT_MARKER_START}{key.replace(self.PLACEHOLDER_PREFIX, '')}{self.TEXT_MARKER_END}"
                    marker2 = f"{self.TEXT_MARKER_START}{key}{self.TEXT_MARKER_END}"
                    
                    if marker1 in new_text:
                        new_text = new_text.replace(marker1, str(value))
                        replaced = True
                    if marker2 in new_text:
                        new_text = new_text.replace(marker2, str(value))
                        replaced = True
                
                if new_text != original_text:
                    run.text = new_text
        
        return replaced
    
    def save_presentation(self, output_path: str) -> bool:
        """
        保存演示文稿
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            bool: 保存是否成功
        """
        if not self.presentation:
            logger.error("未加载演示文稿")
            return False
        
        try:
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 保存演示文稿（已压缩媒体不再重复 deflate，大模板提速明显；
            # 首次使用会自动与标准保存逐成员校验，不一致则退回标准保存）
            save_pptx(self.presentation, output_path)
            
            logger.info(f"成功保存演示文稿: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"保存演示文稿失败: {str(e)}")
            raise SaveError(f"保存演示文稿时出错: {str(e)}")
    
    def close_presentation(self):
        """关闭当前演示文稿"""
        try:
            if self.presentation:
                # python-pptx 不需要显式关闭，只需释放引用
                self.presentation = None
                self.template_path = None
                self._slide_xml = None
                logger.info("已关闭演示文稿")
        except Exception as e:
            logger.error(f"关闭演示文稿失败: {str(e)}")
    
    def close_application(self):
        """关闭应用程序

        对于 python-pptx，此方法仅释放资源
        """
        self.close_presentation()
        self.is_connected = False
        logger.info("PPTX 处理器已关闭")

    def close(self):
        """关闭处理器，释放资源（别名）"""
        self.close_application()
    
    def get_application_info(self) -> Dict[str, str]:
        """
        获取应用程序信息
        
        Returns:
            Dict: 应用程序信息
        """
        try:
            import pptx
            version = getattr(pptx, '__version__', 'unknown')
        except Exception:
            version = 'unknown'
        return {
            'name': 'python-pptx',
            'version': version,
            'type': 'pure_python',
            'description': '纯 Python PPT 处理库，无需 Microsoft Office 或 WPS'
        }
    
    def export_to_images(self, output_dir: str, format: str = "JPG", quality: float = 1.0,
                         base_filename: str = None) -> list:
        """
        导出演示文稿为图片

        引擎由 processing.image_export_engine 决定（auto/office/python）：
        内置 Pillow 引擎直接渲染内存中的 Presentation；COM 引擎（Office/WPS）
        读取快速保存的临时文件。auto 策略下 COM 失败会自动降级内置引擎。

        Args:
            output_dir: 输出目录
            format: 图片格式 (JPG, PNG)，默认 JPG
            quality: 图片质量/分辨率缩放倍数 (1.0-4.0)
            base_filename: 基础文件名（不含扩展名），用于重命名生成的图片

        Returns:
            list: 生成的图片路径列表
        """
        try:
            self.last_export_error = None
            engine, exporter = self._pick_image_engine()
            info = (exporter.get_office_info() if engine == 'office'
                    else exporter.get_engine_info())
            logger.info(f"图片导出引擎: {engine}（{info}）")

            if engine == 'python':
                # 内置引擎直接渲染内存中的 Presentation：不必每行另存临时文件、
                # 重新解压打开、重新预检（模板级上下文只准备一次，逐行复用）
                generated_files = self._render_in_memory(
                    exporter, output_dir, format, quality)
            else:
                # COM 引擎需要一个磁盘文件；用快速保存写出（已压缩媒体不重复压缩）
                temp_path = self._save_temp_fast()
                try:
                    generated_files = self._export_with_engine(
                        exporter, temp_path, output_dir, format, quality)
                finally:
                    try:
                        os.remove(temp_path)
                    except Exception as e:
                        logger.debug(f"清理临时文件失败：{e}")

            # 如果指定了基础文件名，重命名生成的图片
            if base_filename and generated_files:
                renamed_files = []
                for i, old_path in enumerate(generated_files, 1):
                    old_name = os.path.basename(old_path)
                    ext = os.path.splitext(old_name)[1]
                    # 命名规则（有意为之，勿改）：单张图片不带序号，
                    # 多张图片带 _i 序号。下游解析需兼容这两种情况。
                    new_name = f"{base_filename}{ext}" if len(generated_files) == 1 else f"{base_filename}_{i}{ext}"
                    new_path = os.path.join(output_dir, new_name)

                    # 如果新文件名已存在，先删除
                    if os.path.exists(new_path) and new_path != old_path:
                        os.remove(new_path)

                    # 重命名文件
                    if old_path != new_path:
                        os.rename(old_path, new_path)
                        logger.info(f"重命名图片: {old_name} -> {new_name}")
                    renamed_files.append(new_path)
                generated_files = renamed_files

            logger.info(f"成功导出图片到: {output_dir}")
            return generated_files

        except RuntimeError:
            raise
        except Exception as e:
            self.last_export_error = f"{type(e).__name__}: {e}"
            logger.error(f"导出图片时出错: {e}")
            return []

    def _get_python_exporter(self) -> PillowImageExporter:
        """复用的内置渲染器：字体缓存与图片缓存跨行保留（海报背景图逐行复用）"""
        if self._python_exporter is None:
            self._python_exporter = PillowImageExporter()
        return self._python_exporter

    def _ensure_render_context(self):
        """模板级渲染上下文（主题 / 内嵌字体 / 预检）只准备一次，逐行复用。"""
        path = self.template_path
        if self._render_theme is None or self._render_theme[0] != path:
            self._render_theme = (path, load_theme(path))
            font_registry.load_embedded_fonts(path)
        if self._precheck_done != path:
            self._precheck_done = path
            issues = self._get_python_exporter().precheck(path)
            for issue in issues:
                logger.warning(f"[预检] 第{issue.slide}页 {issue.shape!r} [{issue.code}] {issue.message}")
            if issues:
                logger.warning(f"[预检] 共 {len(issues)} 项不支持/风险提示，输出可能与 PowerPoint 存在差异")

    def _render_in_memory(self, exporter, output_dir, format, quality):
        """直接渲染当前 Presentation（占位符已替换），不经过临时文件。"""
        self._ensure_render_context()
        return exporter.export_presentation(
            self.presentation, output_dir, format, quality,
            theme=self._render_theme[1])

    def _save_temp_fast(self) -> str:
        """把当前 Presentation 快速保存到一个临时 .pptx（供 COM 引擎读取）。"""
        import tempfile
        fd, temp_path = tempfile.mkstemp(suffix='.pptx')
        os.close(fd)
        save_pptx(self.presentation, temp_path)
        return temp_path

    # ---------------- 图片导出引擎分层（默认内置渲染 / Office 高保真可选）----------------

    @staticmethod
    def _image_engine_config() -> str:
        """导出引擎策略：'python'（内置渲染，默认）| 'office'（COM 高保真）
        | 'auto'（有 Office 用 COM，否则内置）。
        来源：环境变量 POSTERAGAIN_IMAGE_ENGINE > 运行时配置 processing.image_export_engine
        > 旧键 use_com_interface(true->office)。缺省 python。
        运行时配置在开发/打包环境都能正确定位（exe 旁那份），故打包后同样可配置。"""
        import os
        env = os.environ.get('POSTERAGAIN_IMAGE_ENGINE')
        if env and env.strip().lower() in ('auto', 'office', 'python'):
            return env.strip().lower()
        try:
            try:
                from src.config import ConfigManager
            except ImportError:
                from config_manager import ConfigManager
            cm = ConfigManager()
            eng = cm.get('processing.image_export_engine')
            if isinstance(eng, str) and eng.strip().lower() in ('auto', 'office', 'python'):
                return eng.strip().lower()
            if cm.get('processing.use_com_interface') is True:
                return 'office'
        except Exception as e:
            logger.debug(f"读取图片导出引擎配置失败（用默认 python）: {e}")
        return 'python'

    def invalidate_engine_cache(self):
        """引擎配置变更后调用，使下次导出重新解析策略。"""
        self._engine_strategy = None
        self._engine_cache = {}

    def _pick_image_engine(self):
        """返回 (engine_name, exporter)。auto 策略下 COM 优先、纯 Python 兜底。

        策略每实例解析一次（运行时配置），选择结果按策略缓存：
        一个批次里只检测一次 Office，内置渲染器也跨行复用（字体/图片缓存得以保留）。
        """
        if self._engine_strategy is None:
            self._engine_strategy = self._image_engine_config()
        strategy = self._engine_strategy
        cached = self._engine_cache.get(strategy)
        if cached is not None:
            return cached
        if strategy == 'python':
            result = ('python', self._get_python_exporter())
        elif strategy == 'office':
            from src.core.processors.com_image_exporter import COMImageExporter
            com = COMImageExporter()
            if not com.is_available():
                raise RuntimeError(
                    "导出引擎配置为 office，但未检测到 Microsoft Office/WPS——"
                    "请安装 Office 或把 processing.image_export_engine 改为 auto/python")
            result = ('office', com)
        else:
            # auto：COM 优先
            from src.core.processors.com_image_exporter import COMImageExporter
            com = COMImageExporter()
            if com.is_available():
                result = ('office', com)
            else:
                logger.info("auto 策略：未检测到 Office，降级纯 Python 渲染引擎")
                result = ('python', self._get_python_exporter())
        self._engine_cache[strategy] = result
        return result

    def _export_with_engine(self, exporter, temp_path, output_dir, format, quality):
        """按所选引擎导出；auto 策略下 COM 失败自动降级纯 Python（保持总能出图）"""
        try:
            return exporter.export_to_images(temp_path, output_dir, format, quality)
        except Exception as e:
            if isinstance(exporter, PillowImageExporter):
                raise          # 已经是内置引擎，再"降级"一次只是把同一个失败重跑一遍
            if self._engine_strategy == 'office':
                raise
            logger.warning(f"COM 导出失败（{e}），降级纯 Python 渲染引擎重试")
            return self._render_in_memory(
                self._get_python_exporter(), output_dir, format, quality)

    def get_slide_count(self) -> int:
        """
        获取幻灯片数量

        Returns:
            int: 幻灯片数量
        """
        if self.presentation:
            return len(self.presentation.slides)
        return 0

    def get_shape_count(self) -> int:
        """
        获取形状总数

        Returns:
            int: 形状总数
        """
        if not self.presentation:
            return 0

        count = 0
        for slide in self.presentation.slides:
            count += len(slide.shapes)
        return count
