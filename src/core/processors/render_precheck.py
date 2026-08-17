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
  PIC_CROP          图片裁剪(srcRect)
  PIC_EFFECTS       图片效果(alphaModFix/duotone等)
  SHAPE_GEOM        非矩形形状几何（文字仍会渲染，形状外观不会）
  SHAPE_FILL        形状底色填充（不会渲染）
  SHAPE_EFFECTS     形状效果（阴影等）
  SHAPE_LINE        形状描边
  ANCHOR            垂直对齐非顶端
  AUTOFIT_SCALE     normAutofit 自动缩字
  VERT_TEXT         竖排文字
  NO_WRAP           禁止自动换行
  WORDART           艺术字变形
  BULLETS           项目符号/编号
  LNSPC_PTS         固定磅值行距(spcPts)
  ALIGN_JUST        两端对齐
  INHERIT_RPR       run 无任何格式属性（格式全继承，渲染将退化为默认样式）
  INHERIT_COLOR     run 未显式指定颜色/渐变（颜色继承）
  INHERIT_SIZE      run 未显式指定字号（字号继承）
  GRAD_PATH         路径渐变（仅支持线性 lin）
  TEXT_STROKE       文字描边
  TEXT_EFFECTS      文字 glow/倒影/3D（outerShdw 已支持，不算）
  FONT_MISSING      字体未安装
  BG_REF            背景引用母版/主题(bgRef)
  BG_PATTERN        背景图案填充
"""

import logging
import re
import zipfile
from dataclasses import dataclass
from typing import List, Optional, Set

from pptx import Presentation
from pptx.oxml.ns import qn

logger = logging.getLogger(__name__)


@dataclass
class PrecheckIssue:
    """单个预检问题"""
    slide: int      # 页码（从 1 开始）
    shape: str      # 形状名
    code: str       # 稳定问题码（见模块文档）
    message: str    # 人类可读说明


def _load_theme_fonts(pptx_path: str) -> dict:
    """解析 theme 的 fontScheme，解析 +mn-lt/+mn-ea/+mj-lt/+mj-ea 引用"""
    fonts = {'+mn-lt': '', '+mn-ea': '', '+mj-lt': '', '+mj-ea': ''}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = [n for n in z.namelist() if re.match(r'ppt/theme/theme\d+\.xml$', n)]
            if not names:
                return fonts
            xml = z.read(sorted(names)[0]).decode('utf-8', errors='ignore')
        m = re.search(r'<a:minorFont>(.*?)</a:minorFont>', xml, re.S)
        if m:
            mm = re.search(r'<a:latin typeface="([^"]*)"', m.group(1))
            fonts['+mn-lt'] = mm.group(1) if mm else ''
            mm = re.search(r'<a:ea typeface="([^"]*)"', m.group(1))
            fonts['+mn-ea'] = mm.group(1) if mm else ''
        m = re.search(r'<a:majorFont>(.*?)</a:majorFont>', xml, re.S)
        if m:
            mm = re.search(r'<a:latin typeface="([^"]*)"', m.group(1))
            fonts['+mj-lt'] = mm.group(1) if mm else ''
            mm = re.search(r'<a:ea typeface="([^"]*)"', m.group(1))
            fonts['+mj-ea'] = mm.group(1) if mm else ''
    except Exception as e:
        logger.debug(f"解析主题字体失败: {e}")
    return fonts


class RenderPrecheck:
    """PPTX 渲染预检器（白名单制）"""

    def __init__(self, installed_fonts: Optional[Set[str]] = None):
        """
        Args:
            installed_fonts: 系统已安装字体名集合（小写）。None 表示跳过字体检查。
        """
        self.installed_fonts = installed_fonts
        self.theme_fonts = {}

    # ---------------- 入口 ----------------

    def check_pptx(self, pptx_path: str) -> List[PrecheckIssue]:
        """扫描整个 PPTX，返回问题清单（空列表 = 完全可渲染）"""
        prs = Presentation(pptx_path)
        self.theme_fonts = _load_theme_fonts(pptx_path)
        issues = []
        for idx, slide in enumerate(prs.slides, 1):
            issues.extend(self._check_slide(slide, idx))
        return issues

    # ---------------- 幻灯片级 ----------------

    def _check_slide(self, slide, idx: int) -> List[PrecheckIssue]:
        issues = []
        bg = slide._element.find(qn('p:cSld') + '/' + qn('p:bg'))
        if bg is not None:
            if bg.find(qn('p:bgPr')) is None:
                issues.append(PrecheckIssue(
                    idx, '(幻灯片背景)', 'BG_REF',
                    '背景通过 bgRef 引用母版/主题，仅支持显式 bgPr 背景填充'))
            else:
                bgPr = bg.find(qn('p:bgPr'))
                if bgPr.find(qn('a:pattFill')) is not None:
                    issues.append(PrecheckIssue(
                        idx, '(幻灯片背景)', 'BG_PATTERN', '背景图案填充不支持'))
        for shape in slide.shapes:
            issues.extend(self._check_shape(shape, idx))
        return issues

    # ---------------- 形状级 ----------------

    def _check_shape(self, shape, idx: int) -> List[PrecheckIssue]:
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
            issues.extend(self._check_sp(el, shape, idx, name))
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
            if blipFill.find(qn('a:srcRect')) is not None:
                issues.append(PrecheckIssue(idx, name, 'PIC_CROP', '图片裁剪(srcRect)不会应用'))
            blip = blipFill.find(qn('a:blip'))
            if blip is not None and len(blip) > 0:
                issues.append(PrecheckIssue(idx, name, 'PIC_EFFECTS',
                                            '图片效果(alphaModFix/duotone等)不会应用'))
        return issues

    def _check_sp(self, el, shape, idx, name):
        issues = []
        spPr = el.find(qn('p:spPr'))
        has_text = shape.has_text_frame

        # 形状几何/填充/效果/描边（有文字时文字仍会渲染，但形状外观不会）
        if spPr is not None:
            geom = spPr.find(qn('a:prstGeom'))
            if (geom is not None and geom.get('prst') != 'rect') or \
                    spPr.find(qn('a:custGeom')) is not None:
                prst = geom.get('prst') if geom is not None else 'custGeom'
                issues.append(PrecheckIssue(
                    idx, name, 'SHAPE_GEOM',
                    f'非矩形形状几何({prst})不会渲染' + ('，仅渲染其中文字' if has_text else '')))
            for fill_tag, code, desc in (
                    ('a:solidFill', 'SHAPE_FILL', '纯色'),
                    ('a:gradFill', 'SHAPE_FILL', '渐变'),
                    ('a:blipFill', 'SHAPE_FILL', '图片'),
                    ('a:pattFill', 'SHAPE_FILL', '图案')):
                if spPr.find(qn(fill_tag)) is not None:
                    issues.append(PrecheckIssue(
                        idx, name, code, f'形状底色({desc}填充)不会渲染'))
                    break
            if spPr.find(qn('a:effectLst')) is not None and \
                    len(spPr.find(qn('a:effectLst'))) > 0:
                issues.append(PrecheckIssue(idx, name, 'SHAPE_EFFECTS', '形状效果(阴影等)不会渲染'))
            if spPr.find(qn('a:ln')) is not None:
                issues.append(PrecheckIssue(idx, name, 'SHAPE_LINE', '形状描边不会渲染'))

        if has_text:
            issues.extend(self._check_text_frame(shape, idx, name))
        return issues

    # ---------------- 文本框级 ----------------

    def _check_text_frame(self, shape, idx, name):
        issues = []
        tf = shape.text_frame
        bodyPr = tf._txBody.find(qn('a:bodyPr'))

        if bodyPr is not None:
            anchor = bodyPr.get('anchor')
            if anchor not in (None, 't'):
                issues.append(PrecheckIssue(
                    idx, name, 'ANCHOR', f'垂直对齐 anchor={anchor}（当前仅支持顶端 t）'))
            autofit = bodyPr.find(qn('a:normAutofit'))
            if autofit is not None and (autofit.get('fontScale') or autofit.get('lnSpcReduction')):
                issues.append(PrecheckIssue(
                    idx, name, 'AUTOFIT_SCALE',
                    f'自动缩字 fontScale={autofit.get("fontScale")}，文字将按原始字号渲染'))
            if bodyPr.get('vert') not in (None, 'horz'):
                issues.append(PrecheckIssue(idx, name, 'VERT_TEXT',
                                            f'竖排文字 vert={bodyPr.get("vert")} 不支持'))
            if bodyPr.get('wrap') == 'none':
                issues.append(PrecheckIssue(idx, name, 'NO_WRAP', '禁止自动换行(wrap=none)不支持'))
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
                    issues.extend(self._check_run(run, idx, name))

        if not any(True for para in tf.paragraphs for run in para.runs if run.text):
            pass  # 空文本框无需检查
        return issues

    def _check_run(self, run, idx, name):
        issues = []
        rPr = run._r.find(qn('a:rPr'))
        if rPr is None:
            issues.append(PrecheckIssue(
                idx, name, 'INHERIT_RPR',
                '文字未书写任何格式属性，格式完全依赖继承，将按默认样式(黑 18pt)渲染'))
            return issues

        if rPr.find(qn('a:solidFill')) is None and rPr.find(qn('a:gradFill')) is None:
            issues.append(PrecheckIssue(
                idx, name, 'INHERIT_COLOR', '文字未显式指定颜色/渐变，颜色依赖继承（将渲染为黑色）'))
        if rPr.get('sz') is None:
            issues.append(PrecheckIssue(
                idx, name, 'INHERIT_SIZE', '文字未显式指定字号，字号依赖继承（将渲染为 18pt）'))

        grad = rPr.find(qn('a:gradFill'))
        if grad is not None and grad.find(qn('a:lin')) is None:
            issues.append(PrecheckIssue(idx, name, 'GRAD_PATH', '路径渐变仅支持线性(lin)'))
        if rPr.find(qn('a:ln')) is not None:
            issues.append(PrecheckIssue(idx, name, 'TEXT_STROKE', '文字描边不会渲染'))
        for eff in ('a:glow', 'a:reflection', 'a:scene3d', 'a:relOff', 'a:softEdge'):
            if rPr.find(qn('a:effectLst') + '/' + qn(eff)) is not None:
                issues.append(PrecheckIssue(idx, name, 'TEXT_EFFECTS', f'文字效果({eff[2:]})不支持'))
                break

        # 字体检查（latin/ea；sym 跳过）
        for tag in ('a:ea', 'a:latin'):
            tf_el = rPr.find(qn(tag))
            if tf_el is None:
                continue
            typeface = tf_el.get('typeface') or ''
            if not typeface or typeface.startswith('+'):
                typeface = self.theme_fonts.get(typeface, '')
            if typeface and self.installed_fonts is not None:
                tfn = typeface.lower()
                if not (tfn in self.installed_fonts or
                        any(tfn in f for f in self.installed_fonts)):
                    issues.append(PrecheckIssue(
                        idx, name, 'FONT_MISSING',
                        f'字体 {typeface!r} 未安装，将回退微软雅黑（与 PowerPoint 替换结果可能不同）'))
        return issues
