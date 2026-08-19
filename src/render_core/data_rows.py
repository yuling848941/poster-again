# -*- coding: utf-8 -*-
"""
数据行读取（无 pandas，安卓/桌面共用）
================================================================
支持 csv（utf-8-sig/gbk/utf-8 自动探测）、json（对象数组或 {"rows": [...]}）、
xlsx（openpyxl，纯 Python 可入 Chaquopy）。
返回 (表头列表, 行字典列表)；单元格统一转字符串（None -> ''）。
"""
import io
import json
import csv
import os
from typing import List, Tuple, Dict


def read_rows(data_bytes: bytes, filename: str) -> Tuple[List[str], List[Dict[str, str]]]:
    """
    读取数据文件为 (headers, rows)。rows[i] 为 {表头: 值} 字典。

    扩展名优先；缺失/异常时按内容嗅探兜底（安卓 SAF 偶发返回截断显示名，
    如 "KA递交模" 丢了 ".xlsx"）。

    Raises:
        ValueError: 文件类型不支持 / 编码无法识别 / 结构不符合预期
    """
    ext = os.path.splitext(filename)[1].lower()
    if not data_bytes:
        raise ValueError('数据文件为空')
    if ext == '.csv':
        return _read_csv(data_bytes)
    if ext == '.json':
        return _read_json(data_bytes)
    if ext == '.xlsx':
        return _read_xlsx(data_bytes)
    if data_bytes[:4] == b'PK\x03\x04':
        if b'xl/' in data_bytes:
            return _read_xlsx(data_bytes)
        raise ValueError(f'不支持的数据文件类型: {filename or "(未知)"}（zip 但非 xlsx；支持 csv/json/xlsx）')
    if data_bytes[:2] == b'\xd0\xcf':
        raise ValueError('不支持 .xls 旧格式（请另存为 .xlsx）')
    if data_bytes.lstrip()[:1] in (b'{', b'['):
        return _read_json(data_bytes)
    return _read_csv(data_bytes)


def _read_csv(b: bytes):
    text = None
    for enc in ('utf-8-sig', 'gbk', 'utf-8'):
        try:
            text = b.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError('CSV 编码无法识别（已尝试 utf-8/gbk）')
    rows = [r for r in csv.reader(io.StringIO(text))
            if any(c.strip() for c in r)]
    if not rows:
        return [], []
    headers = [h.strip() for h in rows[0]]
    data = []
    for r in rows[1:]:
        data.append({headers[i]: (r[i].strip() if i < len(r) else '')
                     for i in range(len(headers))})
    return headers, data


def _read_json(b: bytes):
    try:
        obj = json.loads(b.decode('utf-8-sig'))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        raise ValueError(f'JSON 解析失败: {e}')
    if isinstance(obj, dict) and isinstance(obj.get('rows'), list):
        obj = obj['rows']
    if not isinstance(obj, list) or not obj or not all(isinstance(r, dict) for r in obj):
        raise ValueError('JSON 需为对象数组（或 {"rows": [对象, ...]}）')
    headers: List[str] = []
    for r in obj:
        for k in r:
            if k not in headers:
                headers.append(k)
    data = []
    for r in obj:
        data.append({h: ('' if r.get(h) is None else str(r.get(h, '')))
                     for h in headers})
    return headers, data


def _read_xlsx(b: bytes):
    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise ValueError('xlsx 支持需要 openpyxl（pip install openpyxl）') from e
    wb = load_workbook(io.BytesIO(b), read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        grid = [list(r) for r in ws.iter_rows(values_only=True)
                if any(v is not None and str(v).strip() for v in r)]
    finally:
        wb.close()
    if not grid:
        return [], []
    headers = [str(h).strip() if h is not None else '' for h in grid[0]]
    data = []
    for r in grid[1:]:
        data.append({headers[i]: ('' if (i >= len(r) or r[i] is None) else str(r[i]).strip())
                     for i in range(len(headers))})
    return headers, data
