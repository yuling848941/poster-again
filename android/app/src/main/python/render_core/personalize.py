# -*- coding: utf-8 -*-
"""
模板个性化引擎（数据行 -> 个性化海报，一条管线）
================================================================
占位符约定（与 Windows 版 pptx_processor 语义一致，见 AGENTS.md）：
1. Alt Text（替代文本 cNvPr@descr）以 ph_ 开头 -> 整个文本框内容替换
2. 形状名称以 ph_ 开头 -> 整个文本框内容替换
3. 文本内 {{字段名}} 标记 -> 行内替换（支持 {{字段名}} / {{ph_字段名}}）

批量性能（移动端关键）：模板 Presentation 只加载/解析一次；逐行
"替换 -> 渲染 -> 恢复"——幻灯片 XML 缓存还原，避免每行重新解压整个 pptx
（50MB 模板每行可省 ~1.5s 的 zip 解析）。
"""
import io
import logging
import os
import re
import tempfile
from typing import Dict, List, Optional, Tuple

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .pillow_image_exporter import PillowImageExporter
from .format_inheritance import load_theme

logger = logging.getLogger(__name__)

PLACEHOLDER_PREFIX = 'ph_'
TEXT_MARKER_RE = re.compile(r'\{\{\s*([^{}]+?)\s*\}\}')


def _iter_shapes(shapes):
    """递归遍历形状（含组合形状内部，与渲染器 z 序遍历一致）"""
    for sh in shapes:
        yield sh
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                yield from _iter_shapes(sh.shapes)
            except Exception:
                pass


def _alt_text(shape) -> str:
    """替代文本：python-pptx 属性优先，WPS 写法的 cNvPr@descr 兜底"""
    t = getattr(shape, 'alt_text', '') or ''
    if t:
        return t.strip()
    for el in shape._element.iter():
        if el.tag.endswith('}cNvPr'):
            d = el.get('descr') or ''
            if d:
                return d.strip()
            break
    return ''


def _placeholder_key(shape) -> Optional[str]:
    alt = _alt_text(shape)
    if alt.startswith(PLACEHOLDER_PREFIX):
        return alt
    name = getattr(shape, 'name', '') or ''
    if name.startswith(PLACEHOLDER_PREFIX):
        return name
    return None


def _set_shape_text(shape, new_text: str):
    """整框替换：保留首段首 run 格式写入新文本，清空其余（与 pptx_processor 一致）"""
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = new_text
        for run in p0.runs[1:]:
            run.text = ''
    else:
        p0.add_run().text = new_text
    for p in tf.paragraphs[1:]:
        for run in p.runs:
            run.text = ''


def _replace_markers(shape, row: Dict[str, object]) -> bool:
    """行内替换 {{字段名}} / {{ph_字段名}}（与 pptx_processor 一致）"""
    replaced = False
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            text = run.text
            for key, value in row.items():
                bare = key[len(PLACEHOLDER_PREFIX):] if key.startswith(PLACEHOLDER_PREFIX) else key
                for marker in (f'{{{{{bare}}}}}', f'{{{{{key}}}}}'):
                    if marker in text:
                        text = text.replace(marker, str(value))
                        replaced = True
            if text != run.text:
                run.text = text
    return replaced


def match_fields(fields: List[str], headers: List[str]) -> Dict[str, str]:
    """
    模板占位符 <-> 数据表头 匹配：精确 -> 去 ph_ 前缀 -> 忽略大小写/空白
    -> 唯一子串兜底（占位符名是列名的子串且唯一命中时才采用，如
    ph_姓名 -> "营销员姓名"；多候选不猜，留给手动，与 GUI 自动匹配行为一致）。
    返回 {占位符: 表头}；未匹配的占位符不在结果中（渲染时保留原占位文本）。
    """
    exact = set(headers)
    lower = {}
    for h in headers:
        lower.setdefault(str(h).strip().lower(), h)
    result = {}
    for f in fields:
        bare = f[len(PLACEHOLDER_PREFIX):] if f.startswith(PLACEHOLDER_PREFIX) else f
        if f in exact:
            result[f] = f
        elif bare in exact:
            result[f] = bare
        else:
            hit = lower.get(f.strip().lower()) or lower.get(bare.strip().lower())
            if hit is not None:
                result[f] = hit
                continue
            # 唯一子串：bare 清理后为某列清理后表头的子串且全表唯一
            clean_f = re.sub(r'[\s\W_]+', '', bare).lower()
            if clean_f:
                hits = [h for h in headers
                        if clean_f in re.sub(r'[\s\W_]+', '', str(h)).lower()]
                if len(hits) == 1:
                    result[f] = hits[0]
    return result


def sanitize_filename(name: str) -> str:
    """与 Windows 版 ppt_generator._sanitize_filename 同规则：非法字符 -> '_'，去首尾空格点"""
    for ch in '<>:"/\\|?*':
        name = name.replace(ch, '_')
    name = name.strip('. ')
    return name or 'output'


def row_base_name(headers: List[str], row: Dict[str, str], index: int) -> str:
    """
    输出文件基础名（与 Windows 版 ppt_generator 批量导出同规则）：
    f"{index+1}_{姓名}"——姓名列 = 精确 '姓名' 列，否则第一个含 '姓名' 二字的列；
    找不到/值为空 -> 'output'。序号从 1 起，天然避免同名行冲突。
    """
    name_value = ''
    if '姓名' in headers:
        name_value = row.get('姓名', '') or ''
    else:
        for h in headers:
            if '姓名' in str(h):
                name_value = row.get(h, '') or ''
                break
    name_value = str(name_value).strip()
    if not name_value or name_value.lower() == 'nan':
        name_value = 'output'
    return sanitize_filename(f"{index + 1}_{name_value}")


def snapshot_slides(prs) -> List[bytes]:
    """缓存各幻灯片元素的原始 XML，供 restore_slides 恢复（模板只加载一次的关键）。"""
    return [etree.tostring(s._element) for s in prs.slides]


def restore_slides(prs, slide_xml: List[bytes]) -> None:
    """把幻灯片恢复到 snapshot_slides 时的原状（配合就地替换复用同一 Presentation）。

    只就地替换 spTree 的子元素——python-pptx 的 Slide.shapes 是惰性缓存，
    持有 spTree 元素对象引用；整树替换会使缓存指向游离的旧树，后续替换改到
    孤儿节点上（静默失效）。占位符替换只动 spTree 内的文本，母版/版式/主题不受
    影响，因此恢复 spTree 即等价于重新加载模板，但无需重新解压解析整个包。
    """
    P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    for slide, xml in zip(prs.slides, slide_xml):
        fresh = etree.fromstring(xml)
        fresh_sp = fresh.find(f'{{{P}}}cSld/{{{P}}}spTree')
        spTree = slide.shapes._spTree   # python-pptx 缓存持有的同一元素
        for child in list(spTree):
            spTree.remove(child)
        for child in fresh_sp:
            spTree.append(child)


class Personalizer:
    """模板个性化会话：一次加载，逐行 替换->渲染->恢复"""

    def __init__(self, pptx):
        """
        Args:
            pptx: .pptx 文件路径或字节（安卓 assets/文件选择器场景直接传字节）
        """
        if isinstance(pptx, (bytes, bytearray)):
            self._template_bytes = bytes(pptx)
            fd, tmp = tempfile.mkstemp(suffix='.pptx')
            try:
                with os.fdopen(fd, 'wb') as f:
                    f.write(self._template_bytes)
                self._theme = load_theme(tmp)
            finally:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        else:
            with open(pptx, 'rb') as f:
                self._template_bytes = f.read()
            self._theme = load_theme(pptx)
        from . import font_registry
        font_registry.load_embedded_fonts(self._template_bytes)
        self._prs = Presentation(io.BytesIO(self._template_bytes))
        # 幻灯片原始 XML 缓存（恢复用；主题/图片等部件不受文本替换影响）
        self._slide_xml = snapshot_slides(self._prs)
        self._exporter = PillowImageExporter()

    # ---------------- 查询 ----------------

    def fields(self) -> List[str]:
        """模板中的占位符列表（含 ph_ 前缀，{{x}} 归一为 ph_x）"""
        found, seen = [], set()

        def add(key):
            if key and key not in seen:
                seen.add(key)
                found.append(key)

        for slide in self._prs.slides:
            for sh in _iter_shapes(slide.shapes):
                key = _placeholder_key(sh)
                if key:
                    add(key)
                    continue
                if sh.has_text_frame:
                    for m in TEXT_MARKER_RE.findall(sh.text_frame.text):
                        add(PLACEHOLDER_PREFIX + m)
        return sorted(found)

    # ---------------- 替换与渲染 ----------------

    def apply(self, row: Dict[str, object]) -> int:
        """把一行数据写入当前 Presentation（就地修改），返回替换计数"""
        replaced = 0
        for slide in self._prs.slides:
            for sh in _iter_shapes(slide.shapes):
                key = _placeholder_key(sh)
                if key and key in row:
                    if sh.has_text_frame:
                        _set_shape_text(sh, str(row[key]))
                        replaced += 1
                elif sh.has_text_frame:
                    if _replace_markers(sh, row):
                        replaced += 1
        return replaced

    def restore(self):
        """恢复模板原状（配合 apply 复用同一 Presentation）。"""
        restore_slides(self._prs, self._slide_xml)

    def render_row(self, row: Dict[str, object], output_dir: str,
                   image_format: str = 'PNG', quality: float = 1.0) -> List[str]:
        """一行数据 -> 个性化图片（替换 -> 渲染 -> 自动恢复）"""
        self.apply(row)
        try:
            return self._exporter.export_presentation(
                self._prs, output_dir, image_format, quality, theme=self._theme)
        finally:
            self.restore()
