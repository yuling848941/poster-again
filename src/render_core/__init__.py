# -*- coding: utf-8 -*-
"""
render_core：自包含的 PPTX→图片渲染核心（跨平台，安卓/嵌入/GUI 共用）
================================================================
包含：pillow_image_exporter（渲染引擎）+ render_precheck（白名单预检）
+ format_inheritance（OOXML 格式继承解析）+ shape_geometry（预设形状几何）
+ font_registry（字体解析与排版校准常数，捆绑字体在 resources/fonts/）

对外入口：
    render(pptx, output_dir, ...) -> RenderResult        # 一句话渲染（推荐）
    Personalizer / read_rows / match_fields               # 个性化批量管线
    PillowImageExporter / RenderPrecheck                  # 细粒度控制

依赖闭包：{标准库, python-pptx(→lxml), Pillow, numpy(仅文字渐变)}，
无 GUI / 无 pandas / 无配置文件 —— 安卓 Chaquopy 嵌入零改动即可使用。
字体策略默认 auto（Windows 系统优先 / 其他平台捆绑优先），
可用 font_registry.set_policy() 或环境变量 POSTERAGAIN_FONT_POLICY 覆盖。

嵌入到其他工程（如安卓）只需拷贝：
    src/render_core/  +  resources/fonts/
"""
import os
import tempfile
from typing import List, NamedTuple, Union

try:
    from src.render_core.pillow_image_exporter import PillowImageExporter
    from src.render_core.render_precheck import PrecheckIssue, RenderPrecheck
    from src.render_core import font_registry
    from src.render_core.personalize import Personalizer, match_fields, row_base_name
    from src.render_core.data_rows import read_rows
except ImportError:
    from render_core.pillow_image_exporter import PillowImageExporter
    from render_core.render_precheck import PrecheckIssue, RenderPrecheck
    from render_core import font_registry
    from render_core.personalize import Personalizer, match_fields, row_base_name
    from render_core.data_rows import read_rows

__all__ = ['render', 'RenderResult', 'PillowImageExporter', 'RenderPrecheck',
           'PrecheckIssue', 'font_registry', 'Personalizer', 'read_rows',
           'match_fields', 'row_base_name']


class RenderResult(NamedTuple):
    images: List[str]            # 生成的图片路径（按幻灯片顺序）
    issues: List[PrecheckIssue]  # 渲染预检问题（空列表 = 模板完全支持）


def render(pptx: Union[str, bytes], output_dir: str,
           image_format: str = 'PNG', quality: float = 1.0,
           precheck: bool = True) -> RenderResult:
    """
    一句话渲染入口（headless）。

    Args:
        pptx: .pptx 文件路径，或文件内容 bytes（安卓场景：assets 读出后直接传）
        output_dir: 输出目录（不存在则创建）
        image_format: 'PNG' | 'JPG'
        quality: 缩放倍数（1.0 = 96dpi 基准，2.0/3.0 增强）
        precheck: 是否先跑白名单预检（问题码见 render_precheck 模块文档）

    Returns:
        RenderResult(images, issues)；issues 非空时图片仍会生成（不支持元素跳过并告警）
    """
    tmp_path = None
    try:
        if isinstance(pptx, (bytes, bytearray)):
            fd, tmp_path = tempfile.mkstemp(suffix='.pptx')
            with os.fdopen(fd, 'wb') as f:
                f.write(bytes(pptx))
            pptx_path = tmp_path
        else:
            pptx_path = pptx
        os.makedirs(output_dir, exist_ok=True)
        exporter = PillowImageExporter()
        issues = list(exporter.precheck(pptx_path)) if precheck else []
        images = exporter.export_to_images(pptx_path, output_dir,
                                           image_format, quality)
        return RenderResult(images=images, issues=issues)
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
