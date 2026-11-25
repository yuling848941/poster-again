#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
命令行入口（薄封装）：等价于 `python main.py <参数>`。

    python cli.py --template 模板.pptx --data 数据.xlsx --output 输出目录

参数与退出码见 `python cli.py --help`。GUI 仍用 `python main.py`（不带参数）。
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import main  # noqa: E402  （初始化路径与配置；CLI 模式不导入 GUI）

if __name__ == '__main__':
    main.main()
