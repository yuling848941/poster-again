# -*- coding: utf-8 -*-
"""
OOXML 预设形状几何（prstGeom）-> 可绘制路径

返回统一结构供渲染器绘制：
  ('polygon',   [(x, y), ...])     多边形顶点（像素坐标）
  ('roundrect', (x0, y0, x1, y1), radius_px)
  ('ellipse',   (x0, y0, x1, y1))
  None                              不支持的 prst

支持的 prst（SUPPORTED_PRST，预检与渲染器共用此白名单）：
rect / roundRect / ellipse / triangle / rtTriangle / diamond /
parallelogram / trapezoid / pentagon / hexagon / star5 /
rightArrow / leftArrow / upArrow / downArrow / chevron / line

avLst 调整值按 OOXML 默认值实现（gd fmla "val N"），
箭头类为视觉近似（注释标注），满足海报场景。
"""

import math

SUPPORTED_PRST = {
    'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond',
    'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'star5',
    'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'chevron', 'line',
}


def _parse_adj(avLst):
    """avLst -> {name: value(0~100000)}，无则空 dict（各形状用默认值）"""
    adj = {}
    if avLst is not None:
        for gd in avLst:
            fmla = gd.get('fmla', '')
            if fmla.startswith('val '):
                try:
                    adj[gd.get('name', '')] = int(fmla.split()[1])
                except (ValueError, IndexError):
                    pass
    return adj


def _regular_polygon(cx, cy, rx, ry, n, rotate_deg=-90.0):
    pts = []
    for i in range(n):
        t = math.radians(rotate_deg + i * 360.0 / n)
        pts.append((cx + rx * math.cos(t), cy + ry * math.sin(t)))
    return pts


def _star(cx, cy, rx, ry, points, inner_ratio):
    pts = []
    for i in range(points * 2):
        t = math.radians(-90.0 + i * 180.0 / points)
        r = 1.0 if i % 2 == 0 else inner_ratio
        pts.append((cx + rx * r * math.cos(t), cy + ry * r * math.sin(t)))
    return pts


def preset_geometry(prst, x, y, w, h, avLst=None):
    """计算预设形状路径。x/y/w/h 为像素坐标，返回见模块文档"""
    if prst is None:
        prst = 'rect'
    adj = _parse_adj(avLst)
    x0, y0, x1, y1 = x, y, x + w, y + h
    cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
    ss = min(w, h)  # OOXML ss：短边

    def A(name, default):
        return adj.get(name, default) / 100000.0

    if prst == 'rect':
        return ('polygon', [(x0, y0), (x1, y0), (x1, y1), (x0, y1)])

    if prst == 'roundRect':
        radius = ss * A('adj', 16667)
        return ('roundrect', (x0, y0, x1, y1), radius)

    if prst == 'ellipse':
        return ('ellipse', (x0, y0, x1, y1))

    if prst == 'triangle':
        return ('polygon', [(cx, y0), (x1, y1), (x0, y1)])

    if prst == 'rtTriangle':
        return ('polygon', [(x0, y0), (x1, y1), (x0, y1)])

    if prst == 'diamond':
        return ('polygon', [(cx, y0), (x1, cy), (cx, y1), (x0, cy)])

    if prst == 'parallelogram':
        dx = ss * A('adj', 25000)
        return ('polygon', [(x0 + dx, y0), (x1, y0), (x1 - dx, y1), (x0, y1)])

    if prst == 'trapezoid':
        dx = ss * A('adj', 25000)
        return ('polygon', [(x0 + dx, y0), (x1 - dx, y0), (x1, y1), (x0, y1)])

    if prst == 'pentagon':
        return ('polygon', _regular_polygon(cx, cy, w / 2.0, h / 2.0, 5))

    if prst == 'hexagon':
        # OOXML hexagon 横向：左右顶点外扩 adj*ss
        dx = ss * A('adj', 25000)
        return ('polygon', [
            (x0 + dx, y0), (x1 - dx, y0), (x1, cy),
            (x1 - dx, y1), (x0 + dx, y1), (x0, cy)])

    if prst == 'star5':
        ratio = 0.375 + 0.25 * A('adj', 37500)  # adj 37500 -> 0.469 经典五角星
        return ('polygon', _star(cx, cy, w / 2.0, h / 2.0, 5, max(0.1, ratio)))

    if prst in ('rightArrow', 'leftArrow', 'upArrow', 'downArrow'):
        # 视觉近似：轴 = adj1（轴干占比），头长 = adj2 * ss
        shaft = A('adj1', 50000)
        head_len = ss * A('adj2', 50000)
        if prst == 'rightArrow':
            sy0, sy1 = cy - h * shaft / 2, cy + h * shaft / 2
            return ('polygon', [
                (x0, sy0), (x1 - head_len, sy0), (x1 - head_len, y0),
                (x1, cy), (x1 - head_len, y1), (x1 - head_len, sy1), (x0, sy1)])
        if prst == 'leftArrow':
            sy0, sy1 = cy - h * shaft / 2, cy + h * shaft / 2
            return ('polygon', [
                (x1, sy0), (x0 + head_len, sy0), (x0 + head_len, y0),
                (x0, cy), (x0 + head_len, y1), (x0 + head_len, sy1), (x1, sy1)])
        if prst == 'upArrow':
            sx0, sx1 = cx - w * shaft / 2, cx + w * shaft / 2
            return ('polygon', [
                (sx0, y1), (sx0, y0 + head_len), (x0, y0 + head_len),
                (cx, y0), (x1, y0 + head_len), (sx1, y0 + head_len), (sx1, y1)])
        sx0, sx1 = cx - w * shaft / 2, cx + w * shaft / 2
        return ('polygon', [
            (sx0, y0), (sx0, y1 - head_len), (x0, y1 - head_len),
            (cx, y1), (x1, y1 - head_len), (sx1, y1 - head_len), (sx1, y0)])

    if prst == 'chevron':
        dx = ss * A('adj', 50000)
        return ('polygon', [
            (x0, y0), (x1 - dx, y0), (x1, cy),
            (x1 - dx, y1), (x0, y1), (x0 + dx, cy)])

    if prst == 'line':
        return ('polygon', [(x0, cy), (x1, cy), (x1, cy + 1), (x0, cy + 1)])

    return None  # 不支持


def draw_path(draw, path, fill=None, outline=None, width=1):
    """把 preset_geometry 的结果画到 ImageDraw 上"""
    kind = path[0]
    if kind == 'polygon':
        # 多边形需要整数坐标
        pts = [(round(px), round(py)) for px, py in path[1]]
        draw.polygon(pts, fill=fill, outline=outline, width=width)
    elif kind == 'roundrect':
        bbox = tuple(round(v) for v in path[1])
        radius = max(0, round(path[2]))
        draw.rounded_rectangle(bbox, radius=radius, fill=fill,
                               outline=outline, width=width)
    elif kind == 'ellipse':
        bbox = tuple(round(v) for v in path[1])
        draw.ellipse(bbox, fill=fill, outline=outline, width=width)
