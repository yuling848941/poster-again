# -*- coding: utf-8 -*-
"""
PPTX 格式继承解析器

按 OOXML 继承链解析 run / 段落 / bodyPr 的有效格式：
  run:  显式 rPr → 段落 defRPr → 形状 lstStyle(lvlN) → 版式占位符 lstStyle(lvlN)
        → 母版占位符 lstStyle(lvlN) → 母版 txStyles(bodyStyle/titleStyle/otherStyle)
        → 渲染默认（黑 18pt）
  段落: pPr → 各级 lstStyle 的 lvlNpPr（lnSpc/algn 等）
  bodyPr: 形状 bodyPr → 版式占位符 bodyPr → 母版占位符 bodyPr → 默认值

支持解析的属性：sz/b/i/spc、纯色/线性渐变/图片(blipFill)文字填充、
latin/ea 字体（含 +mn/+mj 主题引用）、lnSpc(spcPct)、algn、
anchor、内边距、wrap。
解析失败（未知构造）时置 unresolved=True，由预检据此告警。
"""

import logging
import re
import zipfile
from typing import Optional

from pptx.oxml.ns import qn

logger = logging.getLogger(__name__)

EMU_PER_PT = 12700

# 默认 clrMap（bg1->lt1 等）；母版可覆盖，此处按 OOXML 默认映射
_DEFAULT_CLR_MAP = {
    'bg1': 'lt1', 'tx1': 'dk1', 'bg2': 'lt2', 'tx2': 'dk2',
}


def load_theme(pptx_path: str) -> 'ThemeData':
    """从 pptx zip 解析 theme 的配色 + 字体方案 + 背景填充样式表"""
    theme = ThemeData()
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = [n for n in z.namelist() if re.match(r'ppt/theme/theme\d+\.xml$', n)]
            if not names:
                return theme
            xml = z.read(sorted(names)[0]).decode('utf-8', errors='ignore')
        for name in ('dk1', 'lt1', 'dk2', 'lt2',
                     'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'):
            m = re.search(rf'<a:{name}>(.*?)</a:{name}>', xml, re.S)
            if not m:
                continue
            inner = m.group(1)
            m2 = re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', inner)
            if m2:
                theme.colors[name] = '#' + m2.group(1).upper()
                continue
            m2 = re.search(r'sysClr val="\w+" lastClr="([0-9A-Fa-f]{6})"', inner)
            if m2:
                theme.colors[name] = '#' + m2.group(1).upper()
        for scheme, prefix in (('minorFont', '+mn'), ('majorFont', '+mj')):
            m = re.search(rf'<a:{scheme}>(.*?)</a:{scheme}>', xml, re.S)
            if m:
                mm = re.search(r'<a:latin typeface="([^"]*)"', m.group(1))
                theme.fonts[prefix + '-lt'] = mm.group(1) if mm else ''
                mm = re.search(r'<a:ea typeface="([^"]*)"', m.group(1))
                theme.fonts[prefix + '-ea'] = mm.group(1) if mm else ''
        # 背景填充样式表（bgRef idx 1001+ 引用）
        m = re.search(r'<a:bgFillStyleLst>(.*?)</a:bgFillStyleLst>', xml, re.S)
        if m:
            try:
                from lxml import etree
                root = etree.fromstring(
                    f'<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                    f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                    f'{m.group(1)}</root>')
                theme.bg_fills = list(root)
            except Exception as e:
                logger.debug(f"解析背景样式表失败: {e}")
    except Exception as e:
        logger.debug(f"解析主题失败: {e}")
    return theme


class ThemeData:
    """主题数据：配色 / 字体方案 / 背景填充样式表"""

    def __init__(self):
        self.colors = {'dk1': '#000000', 'lt1': '#FFFFFF',
                       'dk2': '#000000', 'lt2': '#FFFFFF'}
        self.fonts = {'+mn-lt': '', '+mn-ea': '', '+mj-lt': '', '+mj-ea': ''}
        self.bg_fills = []  # lxml 元素列表（bgRef idx-1001 索引）


def resolve_bg_ref(bgRef, theme: 'ThemeData'):
    """解析 p:bgRef（引用主题 bgFillStyleLst[idx-1001]）
    返回 ('solid', rgb, alpha) | ('grad', stops, ang) | None（无法解析）"""
    try:
        from lxml import etree
        idx = int(bgRef.get('idx', '0'))
        if not theme.bg_fills or not (1001 <= idx <= 1000 + len(theme.bg_fills)):
            return None
        el = theme.bg_fills[idx - 1001]
        from lxml import etree
        tag = etree.QName(el).localname
        # phClr：bgRef 自带颜色子元素（通常 schemeClr bg1 -> lt1）
        ph_rgb = (255, 255, 255)
        if len(bgRef):
            ph_rgb, _ = parse_color_el(list(bgRef)[0], theme.colors)

        def color_of(clr_el):
            from lxml import etree as _e
            ctag = _e.QName(clr_el).localname
            if ctag == 'schemeClr' and clr_el.get('val') in ('phClr', 'bg1', 'lt1'):
                alpha = 1.0
                a = clr_el.find(qn('a:alpha'))
                if a is not None:
                    alpha = int(a.get('val')) / 100000.0
                return ph_rgb, alpha
            return parse_color_el(clr_el, theme.colors)

        if tag == 'solidFill':
            rgb, a = color_of(list(el)[0])
            return ('solid', rgb, a)
        if tag == 'gradFill':
            stops = []
            for gs in el.findall(qn('a:gsLst') + '/' + qn('a:gs')):
                pos = int(gs.get('pos')) / 100000.0
                rgb, a = color_of(list(gs)[0])
                stops.append((pos, rgb, a))
            lin = el.find(qn('a:lin'))
            ang = int(lin.get('ang')) / 60000.0 if lin is not None else 0.0
            return ('grad', stops, ang)
        return None  # blipFill/pattFill 背景样式不支持
    except Exception as e:
        logger.debug(f"bgRef 解析失败: {e}")
        return None


def hex2rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def parse_color_el(clr_el, theme_colors) -> tuple:
    """解析颜色元素 -> (rgb_tuple, alpha 0~1)"""
    from lxml import etree
    tag = etree.QName(clr_el).localname
    if tag == 'srgbClr':
        rgb = hex2rgb(clr_el.get('val'))
    elif tag == 'schemeClr':
        name = _DEFAULT_CLR_MAP.get(clr_el.get('val'), clr_el.get('val'))
        rgb = hex2rgb(theme_colors.get(name, '#000000'))
    elif tag == 'sysClr':
        rgb = hex2rgb(clr_el.get('lastClr') or '000000')
    else:
        rgb = (0, 0, 0)
    alpha = 1.0
    a = clr_el.find(qn('a:alpha'))
    if a is not None:
        alpha = int(a.get('val')) / 100000.0
    return rgb, alpha


def parse_fill(rPr, theme_colors):
    """解析 rPr/defRPr 的填充子元素
    返回 None | ('solid', rgb, alpha) | ('grad', stops, ang_deg)
         | ('blip', rId, srcRect or None)"""
    if rPr is None:
        return None
    grad = rPr.find(qn('a:gradFill'))
    if grad is not None:
        stops = []
        for gs in grad.findall(qn('a:gsLst') + '/' + qn('a:gs')):
            pos = int(gs.get('pos')) / 100000.0
            rgb, a = parse_color_el(list(gs)[0], theme_colors)
            stops.append((pos, rgb, a))
        lin = grad.find(qn('a:lin'))
        ang = int(lin.get('ang')) / 60000.0 if lin is not None else 0.0
        return ('grad', stops, ang)
    blip_fill = rPr.find(qn('a:blipFill'))
    if blip_fill is not None:
        blip = blip_fill.find(qn('a:blip'))
        if blip is not None and blip.get(qn('r:embed')):
            src = blip_fill.find(qn('a:srcRect'))
            return ('blip', blip.get(qn('r:embed')), src)
    solid = rPr.find(qn('a:solidFill'))
    if solid is not None:
        rgb, a = parse_color_el(list(solid)[0], theme_colors)
        return ('solid', rgb, a)
    return None


def _ph_info(shape_el):
    """返回形状的 (type, idx, ph元素)；非占位符返回 None"""
    ph = shape_el.find(qn('p:nvSpPr') + '/' + qn('p:nvPr') + '/' + qn('p:ph'))
    if ph is None:
        ph = shape_el.find(qn('p:nvPicPr') + '/' + qn('p:nvPr') + '/' + qn('p:ph'))
        if ph is None:
            return None
    return (ph.get('type', 'obj'), int(ph.get('idx', '0')), ph)


def _lvl(para) -> int:
    pPr = para._p.find(qn('a:pPr'))
    if pPr is not None and pPr.get('lvl'):
        return int(pPr.get('lvl'))
    return 0


class FormatResolver:
    """格式继承解析器（容器可以是 Slide / SlideLayout / SlideMaster）"""

    def __init__(self, container, theme_colors, theme_fonts):
        tag = container._element.tag.split('}')[-1]
        if tag == 'sld':
            self.slide = container
            self.layout = container.slide_layout
            self.master = self.layout.slide_master
        elif tag == 'sldLayout':
            self.slide = None
            self.layout = container
            self.master = container.slide_master
        else:  # sldMaster
            self.slide = None
            self.layout = None
            self.master = container
        self.container_part = container.part
        self.theme_colors = theme_colors
        self.theme_fonts = theme_fonts
        self._layout_ph = None   # {idx: sp element}
        self._master_ph = None
        self._tx_styles = None   # {'body': lvl1pPr, 'title': ..., 'other': ...}
        self.unresolved = False  # 解析链上出现无法处理的构造（预检据此告警）

    # ---------------- 占位符映射 ----------------

    def _ph_map(self, shapes) -> dict:
        result = {}
        for sp in shapes:
            info = _ph_info(sp._element)
            if info:
                result[info[1]] = sp._element
        return result

    def _layout_ph_map(self):
        if self._layout_ph is None:
            self._layout_ph = (self._ph_map(self.layout.shapes)
                               if self.layout is not None else {})
        return self._layout_ph

    def _master_ph_map(self):
        if self._master_ph is None:
            self._master_ph = self._ph_map(self.master.shapes)
        return self._master_ph

    def _get_tx_styles(self):
        if self._tx_styles is None:
            master_el = self.master._element
            tx = master_el.find(qn('p:txStyles'))
            self._tx_styles = {
                'body': tx.find(qn('p:bodyStyle')) if tx is not None else None,
                'title': tx.find(qn('p:titleStyle')) if tx is not None else None,
                'other': tx.find(qn('p:otherStyle')) if tx is not None else None,
            }
        return self._tx_styles

    # ---------------- 继承链构造 ----------------

    def _chain_for(self, shape, lvl, para=None):
        """返回 [(元素, owner_part), ...] 按 OOXML 优先级顺序。
        元素是 rPr/defRPr 提供者所在层级（段落 defRPr / lstStyle lvlNpPr）。"""
        slide_part = self.container_part
        layout_part = self.layout.part if self.layout is not None else self.master.part
        master = self.master

        chain = []

        # 形状自身 lstStyle
        txBody = shape._element.find(qn('p:txBody'))
        if txBody is not None:
            lst = txBody.find(qn('a:lstStyle'))
            if lst is not None:
                lv = lst.find(qn(f'a:lvl{lvl + 1}pPr'))
                if lv is not None:
                    chain.append((lv, slide_part))

        info = _ph_info(shape._element)
        if info:
            ph_type, idx, _ = info
            # 版式占位符
            lv_sp = self._layout_ph_map().get(idx)
            if lv_sp is not None:
                txBody = lv_sp.find(qn('p:txBody'))
                if txBody is not None:
                    lst = txBody.find(qn('a:lstStyle'))
                    if lst is not None:
                        lv = lst.find(qn(f'a:lvl{lvl + 1}pPr'))
                        if lv is not None:
                            chain.append((lv, layout_part))
                    bodyPr = txBody.find(qn('a:bodyPr'))
                    if bodyPr is not None:
                        chain.append((('bodyPr', bodyPr), layout_part))
            # 母版占位符
            ma_sp = self._master_ph_map().get(idx)
            if ma_sp is not None:
                txBody = ma_sp.find(qn('p:txBody'))
                if txBody is not None:
                    lst = txBody.find(qn('a:lstStyle'))
                    if lst is not None:
                        lv = lst.find(qn(f'a:lvl{lvl + 1}pPr'))
                        if lv is not None:
                            chain.append((lv, master.part))
                    bodyPr = txBody.find(qn('a:bodyPr'))
                    if bodyPr is not None:
                        chain.append((('bodyPr', bodyPr), master.part))
            # 母版 txStyles
            key = 'title' if ph_type in ('title', 'ctrTitle') else (
                'other' if ph_type == '' else 'body')
            style = self._get_tx_styles()[key]
            if style is not None:
                lv = style.find(qn(f'a:lvl{lvl + 1}pPr'))
                if lv is not None:
                    chain.append((lv, master.part))
        else:
            style = self._get_tx_styles()['other']
            if style is not None:
                lv = style.find(qn(f'a:lvl{lvl + 1}pPr'))
                if lv is not None:
                    chain.append((lv, master.part))
        return chain

    # ---------------- 有效属性 ----------------

    def run_props(self, shape, para, run) -> dict:
        """解析 run 的有效格式。返回
        {sz_pt, bold, italic, spc_pt, fill, typeface}"""
        lvl = _lvl(para)
        result = {'sz_pt': 18.0, 'bold': False, 'italic': False,
                  'spc_pt': 0.0, 'fill': None, 'typeface': None}
        found = {'sz': False, 'bold': False, 'italic': False,
                 'spc': False, 'fill': False, 'typeface': False}

        candidates = []  # (rPr-like element, owner part)
        rPr = run._r.find(qn('a:rPr'))
        if rPr is not None:
            candidates.append((rPr, self.container_part))
        pPr = para._p.find(qn('a:pPr'))
        if pPr is not None and pPr.find(qn('a:defRPr')) is not None:
            candidates.append((pPr.find(qn('a:defRPr')), self.container_part))
        for el, owner in self._chain_for(shape, lvl):
            if isinstance(el, tuple):
                continue
            defRPr = el.find(qn('a:defRPr'))
            if defRPr is not None:
                candidates.append((defRPr, owner))

        for el, owner in candidates:
            try:
                if not found['sz'] and el.get('sz') is not None:
                    result['sz_pt'] = int(el.get('sz')) / 100.0
                    found['sz'] = True
                if not found['bold'] and el.get('b') is not None:
                    result['bold'] = el.get('b') == '1'
                    found['bold'] = True
                if not found['italic'] and el.get('i') is not None:
                    result['italic'] = el.get('i') == '1'
                    found['italic'] = True
                if not found['spc'] and el.get('spc') is not None:
                    result['spc_pt'] = int(el.get('spc')) / 100.0
                    found['spc'] = True
                if not found['fill']:
                    fill = parse_fill(el, self.theme_colors)
                    if fill is not None:
                        result['fill'] = self._resolve_fill(fill, owner)
                        found['fill'] = True
                if not found['typeface']:
                    for tag in ('a:ea', 'a:latin'):
                        tf_el = el.find(qn(tag))
                        if tf_el is not None and tf_el.get('typeface'):
                            tf = tf_el.get('typeface')
                            result['typeface'] = self.theme_fonts.get(tf, tf) or None
                            found['typeface'] = True
                            break
            except Exception as e:
                logger.debug(f"继承链属性解析失败: {e}")
                self.unresolved = True
        return result

    def _resolve_fill(self, fill, owner_part):
        """blip 填充在此处解析为图片字节（rId 相对于 owner part）"""
        if fill and fill[0] == 'blip':
            try:
                part = owner_part.related_part(fill[1])
                src = fill[2]
                rect = None
                if src is not None:
                    rect = {k: int(src.get(k, '0')) / 100000.0
                            for k in ('l', 't', 'r', 'b')}
                return ('blip', part.blob, rect)
            except Exception as e:
                logger.warning(f"文字图片填充解析失败（按黑色渲染）: {e}")
                self.unresolved = True
                return None
        return fill

    def para_props(self, shape, para) -> dict:
        """解析段落有效属性 {lnSpc_pct, algn}"""
        lvl = _lvl(para)
        result = {'lnSpc_pct': 1.0, 'algn': 'l'}
        found = {'lnSpc': False, 'algn': False}

        pPr = para._p.find(qn('a:pPr'))
        if pPr is not None:
            if pPr.get('algn'):
                result['algn'] = pPr.get('algn')
                found['algn'] = True
            sp = pPr.find(qn('a:lnSpc') + '/' + qn('a:spcPct'))
            if sp is not None:
                result['lnSpc_pct'] = int(sp.get('val')) / 100000.0
                found['lnSpc'] = True

        if not found['lnSpc'] or not found['algn']:
            for el, _ in self._chain_for(shape, lvl):
                if isinstance(el, tuple):
                    continue
                if not found['algn'] and el.get('algn'):
                    result['algn'] = el.get('algn')
                    found['algn'] = True
                if not found['lnSpc']:
                    sp = el.find(qn('a:lnSpc') + '/' + qn('a:spcPct'))
                    if sp is not None:
                        result['lnSpc_pct'] = int(sp.get('val')) / 100000.0
                        found['lnSpc'] = True
        return result

    def body_props(self, shape) -> dict:
        """解析 bodyPr 有效属性
        {anchor, l/r/t/bIns(EMU), wrap, font_scale, lnspc_reduction}"""
        result = {'anchor': 't', 'lIns': 91440, 'rIns': 91440,
                  'tIns': 45720, 'bIns': 45720, 'wrap': 'square',
                  'font_scale': 1.0, 'lnspc_reduction': 0.0}
        found = {'anchor': False, 'wrap': False,
                 'lIns': False, 'rIns': False, 'tIns': False, 'bIns': False}

        def apply(bodyPr):
            if bodyPr is None:
                return
            if not found['anchor'] and bodyPr.get('anchor') is not None:
                result['anchor'] = bodyPr.get('anchor')
                found['anchor'] = True
            if not found['wrap'] and bodyPr.get('wrap') is not None:
                result['wrap'] = bodyPr.get('wrap')
                found['wrap'] = True
            for k in ('lIns', 'rIns', 'tIns', 'bIns'):
                if not found[k] and bodyPr.get(k) is not None:
                    result[k] = int(bodyPr.get(k))
                    found[k] = True
            # normAutofit 缩放只看形状自身（PowerPoint 把计算值存在形状上）
            autofit = bodyPr.find(qn('a:normAutofit'))
            if autofit is not None:
                if autofit.get('fontScale'):
                    result['font_scale'] = int(autofit.get('fontScale')) / 100000.0
                if autofit.get('lnSpcReduction'):
                    result['lnspc_reduction'] = \
                        int(autofit.get('lnSpcReduction')) / 100000.0

        txBody = shape._element.find(qn('p:txBody'))
        if txBody is not None:
            apply(txBody.find(qn('a:bodyPr')))
        info = _ph_info(shape._element)
        if info:
            for el, _ in self._chain_for(shape, 0):
                if isinstance(el, tuple) and el[0] == 'bodyPr':
                    apply(el[1])
        return result
