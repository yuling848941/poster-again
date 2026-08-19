# -*- coding: utf-8 -*-
"""
安卓桥接层字段配置单测（PC 上直接 import poster_bridge 验证逻辑）
覆盖：手动绑定数据列 / 固定值 / 前后缀 / 恢复自动 / 渲染生效

运行：.venv/Scripts/python.exe tests/unit/test_poster_bridge.py
"""
import io
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
sys.path.insert(0, os.path.abspath(os.path.join(
    os.path.dirname(__file__), '..', '..', 'android', 'app', 'src', 'main', 'python')))

import logging
logging.basicConfig(level=logging.ERROR)

from pptx import Presentation
from pptx.util import Inches
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
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


def build_env():
    """模板（3 占位符）+ CSV 数据（含两个'险种'列制造多候选不猜）"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    from pptx.oxml.ns import qn
    for i, (name, ph, text) in enumerate([
            ('姓名', 'ph_姓名', '张三'), ('标保', 'ph_标保', '999'),
            ('险种', 'ph_险种', '万能险'), ('趸期', 'ph_个险趸期判定', '趸')]):
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3 + i * 1.2), Inches(8), Inches(1))
        run = tb.text_frame.paragraphs[0].add_run()
        run.text = text
        sp = tb._element
        sp.find(qn('p:nvSpPr')).find(qn('p:cNvPr')).set('name', ph)
        sp.find(qn('p:nvSpPr')).find(qn('p:cNvPr')).set('descr', ph)
    tpl = os.path.join(tempfile.mkdtemp(prefix='bridge_'), 't.pptx')
    prs.save(tpl)
    csv = '姓名,险种编号,险种名称,标保\n李逵,Z01,一生中意,20000\n'.encode('utf-8-sig')
    return tpl, csv


def main():
    import poster_bridge as pb
    tpl, csv = build_env()

    print("(1) start_session：headers 与匹配")
    r = json.loads(pb.start_session(open(tpl, 'rb').read(), csv, 'd.csv'))
    check('返回 headers', r['headers'] == ['姓名', '险种编号', '险种名称', '标保'])
    check('姓名/标保自动匹配', any('ph_姓名 → 姓名' in m for m in r['matched'])
          and any('ph_标保 → 标保' in m for m in r['matched']))
    check('ph_险种 多候选不猜（未匹配）', 'ph_险种' in r['unmatched'])
    check('ph_个险趸期判定 无列未匹配', 'ph_个险趸期判定' in r['unmatched'])

    print("(2) apply_field_configs：绑列 + 固定值 + 前后缀")
    cfgs = {
        'ph_险种': {'source': 'column', 'value': '险种名称', 'prefix': '', 'suffix': '（福运版）'},
        'ph_个险趸期判定': {'source': 'fixed', 'value': '个险', 'prefix': '', 'suffix': ''},
        'ph_标保': {'source': 'column', 'value': '标保', 'prefix': '', 'suffix': '元'},
        'ph_姓名': {'source': 'auto', 'value': '', 'prefix': '', 'suffix': ''},
    }
    eff = json.loads(pb.apply_field_configs(r['session'], json.dumps(cfgs, ensure_ascii=False)))
    eff_map = {e.split(' → ')[0]: e.split(' → ', 1)[1] for e in eff['effective']}
    check('险种绑列+后缀', '险种名称' in eff_map['ph_险种'] and '福运版' in eff_map['ph_险种'])
    check('趸期固定值', '个险' in eff_map['ph_个险趸期判定'])
    check('姓名恢复自动', eff_map['ph_姓名'] == '姓名')

    print("(3) render_index 生效（渲染产物个性化值正确）")
    s = pb._sessions[r['session']]
    data = pb._row_data(s, s['rows'][0])
    check('列绑定取值', data['ph_险种'] == '一生中意（福运版）')
    check('固定值取值', data['ph_个险趸期判定'] == '个险')
    check('后缀拼接', data['ph_标保'] == '20000元')
    check('自动匹配不受影响', data['ph_姓名'] == '李逵')
    imgs = pb.render_index(r['session'], 0, tempfile.mkdtemp(prefix='bridgeout_'))
    check('渲染出图并按规则命名', os.path.basename(imgs[0]) == '1_李逵.png')
    # 值上墙验证：pptx 会话的替换后文本
    txt = ''.join(pb._sessions[r['session']]['pz']._prs.slides[0]._element.itertext())
    check('渲染后已恢复（apply/restore 循环安全）', '张三' in txt)

    print("(4) 恢复自动")
    pb.apply_field_configs(r['session'], json.dumps(
        {'ph_险种': {'source': 'auto', 'value': '', 'prefix': '', 'suffix': ''}}))
    s2 = pb._sessions[r['session']]
    data2 = pb._row_data(s2, s2['rows'][0])
    check('险种恢复自动（未匹配→不替换）', 'ph_险种' not in data2)

    pb.end_session(r['session'])
    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
