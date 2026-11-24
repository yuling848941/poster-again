# -*- coding: utf-8 -*-
"""安卓桥接层：Kotlin(Chaquopy) <-> render_core 的薄封装。
会话式批量管线：模板 Presentation 只解析一次，逐行"替换->渲染->恢复"。
字段配置：自动匹配 + 手动绑定（数据列/固定值/前后缀，语义对齐 PC 端
matching_rules + text_addition_rules）。"""
import os

# 渲染引擎会把「缩放后的图片位图」缓存起来供逐行复用（海报背景图每行都一样）。
# 默认上限 256MB 是按桌面内存给的，手机上收到 64MB。
os.environ.setdefault('POSTERAGAIN_PIC_CACHE_MB', '64')

_sessions = {}


def set_env(key, value):
    """供 Kotlin 注入环境变量（如 POSTERAGAIN_FONT_POLICY，font_registry 惰性读取）"""
    import os
    os.environ[str(key)] = str(value)


def render_pptx_to_png(pptx_bytes, out_dir):
    """单模板快渲染（无数据替换）；预检问题 print 到 logcat"""
    from render_core import render, font_registry
    font_registry.set_policy('bundled_first')   # 安卓：始终用捆绑思源字体
    result = render(bytes(pptx_bytes), out_dir, 'PNG', 1.0)
    for issue in result.issues:
        print(f"[预检] {issue.code}: {issue.message}")
    return [str(p) for p in result.images]


def start_session(pptx_bytes, data_bytes=None, filename=''):
    """
    打开批量会话：解析模板占位符 + 数据行 + 字段匹配 + 模板渲染预检。
    Returns: JSON 字符串 {session, fields, headers, row_count, matched[],
    unmatched[], issues[{slide, shape, code, message}]}
    issues 为模板级渲染预检问题（近似渲染/字体回退提示），空列表 = 完全支持。
    """
    import json
    import os
    import tempfile
    from render_core import Personalizer, match_fields, RenderPrecheck, font_registry
    from render_core.data_rows import read_rows
    raw = bytes(pptx_bytes)
    pz = Personalizer(raw)
    fields = pz.fields()
    headers, rows = [], []
    if data_bytes is not None:
        headers, rows = read_rows(bytes(data_bytes), str(filename))
    matches = match_fields(fields, headers)

    # 模板级预检（一次，毫秒级 XML 扫描；字体集=捆绑思源，命中 FONT_MISSING 即安卓端会回退）
    issues = []
    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix='.pptx')
        with os.fdopen(fd, 'wb') as f:
            f.write(raw)
        installed = font_registry.available_typefaces()
        for it in RenderPrecheck(installed_fonts=installed).check_pptx(tmp_path):
            issues.append({'slide': it.slide, 'shape': it.shape,
                           'code': it.code, 'message': it.message})
    except Exception as e:  # 预检失败不阻塞会话
        print(f"[预检] 跳过: {e}")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    sid = 's%d' % (len(_sessions) + 1)
    _sessions[sid] = {'pz': pz, 'headers': headers, 'rows': rows,
                      'matches': matches, 'manual': {}}
    print(f"[批量] 模板占位符 {len(fields)} 个：{fields}")
    print(f"[批量] 数据 {len(rows)} 行 × {len(headers)} 列；匹配 {len(matches)} 个字段；预检 {len(issues)} 项")
    return json.dumps({
        'session': sid,
        'fields': fields,
        'headers': headers,
        'row_count': len(rows),
        'matched': [f'{f} → {h}' for f, h in matches.items()],
        'unmatched': [f for f in fields if f not in matches],
        'issues': issues,
    }, ensure_ascii=False)


def apply_field_configs(sid, configs_json):
    """
    设置手动字段配置（覆盖自动匹配）。
    configs_json: {字段: {"source": "column"|"fixed", "value": 列名或固定值,
                          "prefix": str, "suffix": str}}
    source=auto 的条目表示恢复自动匹配（移除手动配置）。
    Returns: JSON {effective: ["字段 → 展示串", ...]}
    """
    import json
    s = _sessions[sid]
    manual = {}
    for field, c in json.loads(str(configs_json)).items():
        if c.get('source') == 'auto':
            continue
        manual[field] = {'source': c.get('source', 'fixed'),
                         'value': str(c.get('value', '')),
                         'prefix': str(c.get('prefix', '')),
                         'suffix': str(c.get('suffix', ''))}
    s['manual'] = manual
    eff = []
    for f in s['pz'].fields():
        m = manual.get(f)
        if m:
            v = m['value'] if m['source'] == 'fixed' else f"[列]{m['value']}"
            extra = []
            if m['prefix']:
                extra.append(f"前缀'{m['prefix']}'")
            if m['suffix']:
                extra.append(f"后缀'{m['suffix']}'")
            eff.append(f"{f} → {v}{'（' + '，'.join(extra) + '）' if extra else ''}")
        elif f in s['matches']:
            eff.append(f"{f} → {s['matches'][f]}")
        else:
            eff.append(f"{f} →（未匹配）")
    print(f"[批量] 应用字段配置：手动 {len(manual)} 个")
    return json.dumps({'effective': eff}, ensure_ascii=False)


def _row_data(s, row):
    """按 自动匹配+手动配置 组装单行替换数据（含前后缀拼接）"""
    data = {}
    for f in s['pz'].fields():
        m = s.get('manual', {}).get(f)
        if m:
            if m['source'] == 'column':
                v = str(row.get(m['value'], '') or '')
            else:
                v = m['value']
            data[f] = m['prefix'] + v + m['suffix']
        elif f in s['matches'] and s['matches'][f] in row:
            data[f] = str(row[s['matches'][f]] or '')
    return data


def render_index(sid, index, out_dir):
    """渲染数据第 index 行 -> PNG 路径列表。
    产物按 PC 版规则重命名：单页 {序号_姓名}.png，多页 {序号_姓名}_i.png"""
    import os
    from render_core import row_base_name
    s = _sessions[sid]
    row = s['rows'][index]
    data = _row_data(s, row)
    images = [str(p) for p in s['pz'].render_row(data, str(out_dir), 'PNG', 1.0)]
    base = row_base_name(s['headers'], row, index)
    renamed = []
    for i, old in enumerate(images, 1):
        ext = os.path.splitext(old)[1] or '.png'
        new = os.path.join(os.path.dirname(old),
                           f"{base}{ext}" if len(images) == 1 else f"{base}_{i}{ext}")
        if new != old:
            if os.path.exists(new):
                os.remove(new)
            os.rename(old, new)
        renamed.append(new)
    return renamed


def render_preview(sid, out_dir):
    """无数据渲染一张模板预览"""
    images = _sessions[sid]['pz'].render_row({}, str(out_dir), 'PNG', 1.0)
    return [str(p) for p in images]


def end_session(sid):
    _sessions.pop(str(sid), None)
