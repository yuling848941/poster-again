# -*- coding: utf-8 -*-
"""兼容 shim：渲染引擎已迁至 src/render_core/，本模块转发导入（保持旧路径可用）"""
import sys as _sys

try:
    from src.render_core import pillow_image_exporter as _real
except ImportError:
    from render_core import pillow_image_exporter as _real

_sys.modules[__name__] = _real
