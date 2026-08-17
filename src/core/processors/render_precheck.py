# -*- coding: utf-8 -*-
"""
纯 Python 渲染预检（白名单制）

在渲染前扫描 PPTX，识别 PillowImageExporter 不支持/可能渲染错误的元素，
输出结构化问题清单。问题码（code）稳定，供后续后端链（回退 COM）程序化决策。

⚠️ 白名单与渲染器能力同源：修改 pillow_image_exporter.py 的支持范围时，
必须同步更新本模块的检查项。

问题码一览：
  ROTATION          形状旋转/翻转
  GROUP             组合形状
  TABLE             表格
  CHART             图表
  SMARTART          SmartArt
  OLE               OLE/嵌入对象
  CONNECTOR         连接线
  PIC_EFFECTS       图片效果(alphaModFix/duotone等)
  SHAPE_GEOM        custGeom 或白名单外 prst（文字仍会渲染，形状外观不会）
  SHAPE_FILL        形状路径渐变/图案填充
  SHAPE_EFFECTS     outerShdw 以外的形状效果
  SHAPE_LINE        形状非纯色/虚线描边
  VERT_TEXT         竖排文字
  NO_WRAP           禁止自动换行
  WORDART           艺术字变形
  BULLETS           项目符号/编号
  LNSPC_PTS         固定磅值行距(spcPts)
  ALIGN_JUST        两端对齐
  INHERIT_UNRESOLVED 继承链存在无法解析的构造
  GRAD_PATH         路径渐变（仅支持线性 lin）
  TEXT_EFFECTS      文字 glow/倒影/3D（outerShdw/描边已支持，不算）
  FONT_MISSING      字体未安装
  BG_REF            bgRef 主题背景样式无法解析
  BG_PATTERN        背景图案填充

已支持（对应问题码已退役）：格式继承（INHERIT_RPR/COLOR/SIZE）、
normAutofit 缩字（AUTOFIT_SCALE）、文字纯色描边（TEXT_STROKE）、
图片裁剪 srcRect（PIC_CROP）、bgPr/bgRef 背景、图片文字填充 blipFill、
18 种预设形状几何 + 纯色/线性渐变/图片形状填充 + outerShdw 形状阴影 + 纯色实线描边。
"""

import logging
import re
from dataclasses import dataclass
from typing import List, Optional, Set

from pptx import Presentation
from pptx.oxml.ns import qn

try:
    from src.core.processors.format_inheritance import (
        FormatResolver, load_theme, resolve_bg_ref,
    )
    from src.core.processors.shape_geometry import SUPPORTED_PRST
except ImportError:
    from core.processors.format_inheritance import (
        FormatResolver, load_theme, resolve_bg_ref,
    )
    from core.processors.shape_geometry import SUPPORTED_PRST

logger = logging.getLogger(__name__)


@dataclass
class PrecheckIssue:
    """单个预检问题"""
    slide: int      # 页码（从 1 开始）
    shape: str      # 形状名
    code: str       # 稳定问题码（见模块文档）
    message: str    # 人类可读说明


class RenderPrecheck:
    """PPTX 渲染预检器（白名单制）"""

    def __init__(self, installed_fonts: Optional[Set[str]] = None):
        """
        Args:
            installed_fonts: 系统已安装字体名集合（小写）。None 表示跳过字体检查。
        """
        self.installed_fonts = installed_fonts

    # ---------------- 入口 ----------------

    def check_pptx(self, pptx_path: str) -> List[PrecheckIssue]:
        """扫描整个 PPTX，返回问题清单（空列表 = 完全可渲染）"""
        prs = Presentation(pptx_path)
        theme = load_theme(pptx_path)
        self.theme = theme
        issues = []
        for idx, slide in enumerate(prs.slides, 1):
            resolver = FormatResolver(slide, theme.colors, theme.fonts)
            issues.extend(self._check_slide(slide, idx, resolver))
        return issues

    # ---------------- 幻灯片级 ----------------

    def _check_slide(self, slide, idx: int, resolver) -> List[PrecheckIssue]:
        issues = []
        # 背景继承链：slide -> layout -> master，任一层有显式 bgPr 即可
        containers = (slide, slide.slide_layout, slide.slide_layout.slide_master)
        bgs = [c._element.find(qn('p:cSld') + '/' + qn('p:bg'))
               for c in containers]
        bg = next((b for b in bgs if b is not None), None)
        if bg is not None:
            bgRef = bg.find(qn('p:bgRef'))
            if bgRef is not None:
                if resolve_bg_ref(bgRef, getattr(self, 'theme', None)) is None:
                    issues.append(PrecheckIssue(
                        idx, '(幻灯片背景)', 'BG_REF',
                        'bgRef 引用的主题背景样式无法解析，将沿继承链回退'))
            else:
                bgPr = bg.find(qn('p:bgPr'))
                if bgPr is not None and bgPr.find(qn('a:pattFill')) is not None:
                    issues.append(PrecheckIssue(
                        idx, '(幻灯片背景)', 'BG_PATTERN', '背景图案填充不支持'))
        for shape in slide.shapes:
            issues.extend(self._check_shape(shape, idx, resolver))
        if resolver.unresolved:
            issues.append(PrecheckIssue(
                idx, '(格式继承)', 'INHERIT_UNRESOLVED',
                '格式继承链存在无法解析的构造，对应文字可能退化为默认样式'))
        return issues

    # ---------------- 形状级 ----------------

    def _check_shape(self, shape, idx: int, resolver) -> List[PrecheckIssue]:
        el = shape._element
        name = getattr(shape, 'name', '') or '(未命名)'
        tag = el.tag.split('}')[-1]
        issues = []

        # 旋转/翻转
        xfrm = el.find(qn('p:spPr') + '/' + qn('a:xfrm')) if tag != 'graphicFrame' else None
        if xfrm is not None and (int(xfrm.get('rot', '0')) or xfrm.get('flipH') or xfrm.get('flipV')):
            issues.append(PrecheckIssue(
                idx, name, 'ROTATION', f'形状旋转/翻转（rot={xfrm.get("rot")}, '
                f'flipH={xfrm.get("flipH")}, flipV={xfrm.get("flipV")}）不会渲染'))

        # 组合/连接线/嵌入框架
        if tag == 'grpSp':
            issues.append(PrecheckIssue(idx, name, 'GROUP', '组合形状整体不会渲染'))
            return issues
        if tag == 'cxnSp':
            issues.append(PrecheckIssue(idx, name, 'CONNECTOR', '连接线不会渲染'))
            return issues
        if tag == 'graphicFrame':
            issues.extend(self._check_graphic_frame(el, idx, name))
            return issues

        if tag == 'pic':
            issues.extend(self._check_picture(el, idx, name))
            return issues

        if tag == 'sp':
            issues.extend(self._check_sp(el, shape, idx, name, resolver))
        else:
            issues.append(PrecheckIssue(idx, name, 'SHAPE_GEOM', f'未知形状类型 {tag}'))
        return issues

    def _check_graphic_frame(self, el, idx, name):
        issues = []
        if el.find(qn('a:graphic') + '/' + qn('a:graphicData') + '/' + qn('a:tbl')) is not None:
            issues.append(PrecheckIssue(idx, name, 'TABLE', '表格不会渲染'))
        gd = el.find(qn('a:graphic') + '/' + qn('a:graphicData'))
        uri = gd.get('uri', '') if gd is not None else ''
        if 'chart' in uri:
            issues.append(PrecheckIssue(idx, name, 'CHART', '图表不会渲染'))
        elif 'diagram' in uri:
            issues.append(PrecheckIssue(idx, name, 'SMARTART', 'SmartArt 不会渲染'))
        elif 'ole' in uri.lower():
            issues.append(PrecheckIssue(idx, name, 'OLE', 'OLE 嵌入对象不会渲染'))
        return issues

    def _check_picture(self, el, idx, name):
        issues = []
        blipFill = el.find(qn('p:blipFill'))
        if blipFill is not None:
            blip = blipFill.find(qn('a:blip'))
            if blip is not None and len(blip) > 0:
                issues.append(PrecheckIssue(idx, name, 'PIC_EFFECTS',
                                            '图片效果(alphaModFix/duotone等)不会应用'))
        return issues

    def _check_sp(self, el, shape, idx, name, resolver):
        issues = []
        spPr = el.find(qn('p:spPr'))
        has_text = shape.has_text_frame

        # 形状几何/填充/效果/描边 —— 白名单与 shape_geometry.SUPPORTED_PRST 同源
        if spPr is not None:
            geom = spPr.find(qn('a:prstGeom'))
            if spPr.find(qn('a:custGeom')) is not None:
                issues.append(PrecheckIssue(
                    idx, name, 'SHAPE_GEOM',
                    '自定义几何(custGeom)不会渲染' + ('，仅渲染其中文字' if has_text else '')))
            elif geom is not None and geom.get('prst') not in SUPPORTED_PRST:
                issues.append(PrecheckIssue(
                    idx, name, 'SHAPE_GEOM',
                    f'形状几何({geom.get("prst")})不在支持列表'
                    + ('，仅渲染其中文字' if has_text else '')))
            # 填充：纯色/线性渐变/图片已支持；路径渐变与图案不支持
            grad = spPr.find(qn('a:gradFill'))
            if grad is not None and grad.find(qn('a:lin')) is None:
                issues.append(PrecheckIssue(
                    idx, name, 'SHAPE_FILL', '形状路径渐变填充不支持'))
            if spPr.find(qn('a:pattFill')) is not None:
                issues.append(PrecheckIssue(
                    idx, name, 'SHAPE_FILL', '形状图案填充不支持'))
            # 效果：outerShdw 已支持，其余不支持
            eff = spPr.find(qn('a:effectLst'))
            if eff is not None and len(eff) > 0 and \
                    any(et.tag != qn('a:outerShdw') for et in eff):
                issues.append(PrecheckIssue(
                    idx, name, 'SHAPE_EFFECTS',
                    f'形状效果({[et.tag.split("}")[-1] for et in eff if et.tag != qn("a:outerShdw")]})不支持'))
            # 描边：纯色实线已支持
            ln = spPr.find(qn('a:ln'))
            if ln is not None:
                dash_el = ln.find(qn('a:prstDash'))
                dashed = dash_el is not None and dash_el.get('val') != 'solid'
                no_solid = ln.find(qn('a:solidFill')) is None and not dashed
                if dashed:
                    issues.append(PrecheckIssue(
                        idx, name, 'SHAPE_LINE', '形状虚线描边按实线渲染'))
                elif no_solid and len(ln) > 0:
                    issues.append(PrecheckIssue(
                        idx, name, 'SHAPE_LINE', '形状非纯色描边不支持'))

        if has_text:
            issues.extend(self._check_text_frame(shape, idx, name, resolver))
        return issues

    # ---------------- 文本框级 ----------------

    def _check_text_frame(self, shape, idx, name, resolver):
        issues = []
        tf = shape.text_frame
        bodyPr = tf._txBody.find(qn('a:bodyPr'))

        # 换行走继承解析后的有效值（anchor ctr/b 已支持）
        bp = resolver.body_props(shape)
        if bp['wrap'] == 'none':
            issues.append(PrecheckIssue(idx, name, 'NO_WRAP', '禁止自动换行(wrap=none)不支持'))

        if bodyPr is not None:
            if bodyPr.get('vert') not in (None, 'horz'):
                issues.append(PrecheckIssue(idx, name, 'VERT_TEXT',
                                            f'竖排文字 vert={bodyPr.get("vert")} 不支持'))
            if bodyPr.find(qn('a:prstTxWarp')) is not None:
                issues.append(PrecheckIssue(idx, name, 'WORDART', '艺术字变形不支持'))

        for para in tf.paragraphs:
            pPr = para._p.find(qn('a:pPr'))
            if pPr is not None:
                for bu in ('a:buChar', 'a:buAutoNum', 'a:buImg'):
                    if pPr.find(qn(bu)) is not None:
                        issues.append(PrecheckIssue(idx, name, 'BULLETS', '项目符号/编号不会渲染'))
                        break
                if pPr.find(qn('a:lnSpc') + '/' + qn('a:spcPts')) is not None:
                    issues.append(PrecheckIssue(idx, name, 'LNSPC_PTS',
                                                '固定磅值行距(spcPts)按单倍行距渲染'))
                if pPr.get('algn') in ('just', 'justLow'):
                    issues.append(PrecheckIssue(idx, name, 'ALIGN_JUST', '两端对齐按左对齐渲染'))
            for run in para.runs:
                if run.text:
                    issues.extend(self._check_run(run, para, shape, idx, name, resolver))
        return issues

    def _check_run(self, run, para, shape, idx, name, resolver):
        issues = []
        rPr = run._r.find(qn('a:rPr'))
        if rPr is not None:
            issues.extend(self._check_run_explicit(rPr, idx, name))

        # 字体检查走继承解析后的有效字体（+mn/+mj 主题引用已展开）
        props = resolver.run_props(shape, para, run)
        if props['typeface'] and self.installed_fonts is not None:
            tfn = props['typeface'].lower()
            if not (tfn in self.installed_fonts or
                    any(tfn in f for f in self.installed_fonts)):
                issues.append(PrecheckIssue(
                    idx, name, 'FONT_MISSING',
                    f'字体 {props["typeface"]!r} 未安装，将回退微软雅黑'
                    f'（与 PowerPoint 替换结果可能不同）'))
        return issues

    def _check_run_explicit(self, rPr, idx, name):
        issues = []
        grad = rPr.find(qn('a:gradFill'))
        if grad is not None and grad.find(qn('a:lin')) is None:
            issues.append(PrecheckIssue(idx, name, 'GRAD_PATH', '路径渐变仅支持线性(lin)'))
        for eff in ('a:glow', 'a:reflection', 'a:scene3d', 'a:relOff', 'a:softEdge'):
            if rPr.find(qn('a:effectLst') + '/' + qn(eff)) is not None:
                issues.append(PrecheckIssue(idx, name, 'TEXT_EFFECTS', f'文字效果({eff[2:]})不支持'))
                break
        return issues
