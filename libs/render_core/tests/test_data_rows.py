# -*- coding: utf-8 -*-
"""
数据行读取（data_rows）单元测试：csv 三编码 / json 两种结构 / xlsx / 错误处理。

运行：.venv/Scripts/python.exe libs/render_core/tests/test_data_rows.py
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
# 本库（render_core）与它依赖的 fontfx 都在仓库里；已 pip 安装时这几行是冗余的
for _p in (os.path.join(_HERE, '..', 'src'),
           os.path.join(_HERE, '..', '..', 'fontfx', 'src')):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from render_core.data_rows import read_rows

PASS = 0
FAIL = 0


def check(desc, cond):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  ✅ {desc}")
    else:
        FAIL += 1
        print(f"  ❌ {desc}")


def main():
    print('(1) CSV')
    h, rows = read_rows('姓名,城市,金额\n李逵,深圳,8888\n王小明,东莞,66\n'.encode('utf-8-sig'), 'a.csv')
    check('utf-8-sig 表头', h == ['姓名', '城市', '金额'])
    check('两行数据', len(rows) == 2 and rows[0]['姓名'] == '李逵' and rows[1]['金额'] == '66')
    h2, rows2 = read_rows('姓名,城市\n李逵,深圳\n'.encode('gbk'), 'b.csv')
    check('gbk 自动探测', rows2[0]['城市'] == '深圳')
    h3, rows3 = read_rows('姓名\n\n李逵\n'.encode('utf-8'), 'c.csv')
    check('空行跳过', len(rows3) == 1)

    print('(2) JSON')
    h, rows = read_rows('[{"姓名":"李逵","金额":"88"},{"姓名":"王小明"}]'.encode(), 'd.json')
    check('对象数组 + 缺列补空', rows[1]['金额'] == '' and len(rows) == 2)
    h2, rows2 = read_rows('{"rows":[{"x":"1"}]}'.encode(), 'e.json')
    check('rows 包装结构', rows2[0]['x'] == '1')

    print('(3) XLSX')
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(['姓名', '城市'])
    ws.append(['李逵', '深圳'])
    ws.append([None, None])
    ws.append(['王小明', ''])
    import io
    buf = io.BytesIO()
    wb.save(buf)
    h, rows = read_rows(buf.getvalue(), 'f.xlsx')
    check('xlsx 表头', h == ['姓名', '城市'])
    check('xlsx 行（空行跳过/空单元格转空串）',
          len(rows) == 2 and rows[0]['城市'] == '深圳' and rows[1]['城市'] == '')

    print('(4) 错误处理')
    for desc, blob, name in [
        ('空文件', b'', 'a.csv'),
        ('zip 但非 xlsx', b'PK\x03\x04' + b'junk' * 10, 'a.zip'),
        ('坏 JSON', b'{oops', 'a.json'),
        ('未知编码 CSV', b'\xff\xfe\x00', 'a.csv'),
    ]:
        try:
            read_rows(blob, name)
            check(f'{desc} 应抛 ValueError', False)
        except ValueError:
            check(f'{desc} 抛 ValueError', True)

    print('(5) 内容嗅探兜底（安卓 SAF 截断名丢扩展名）')
    from openpyxl import Workbook
    import io as _io
    buf = _io.BytesIO()
    Workbook().save(buf)
    h, rows = read_rows(buf.getvalue(), 'KA递交模')
    check('截断名 xlsx 嗅探', isinstance(rows, list))
    h, rows = read_rows(b'a,b\n1,2\n', 'data')
    check('无扩展名文本按 CSV', h == ['a', 'b'] and rows[0]['a'] == '1')
    h, rows = read_rows(b'[{"x": 1}]', 'data')
    check('无扩展名 JSON 嗅探', h == ['x'] and rows[0]['x'] == '1')

    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
