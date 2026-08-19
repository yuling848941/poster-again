#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量容错测试：单行失败跳过 + 失败行记录 + row_indices 单独重试 + 模板预检上报

覆盖 batch_generate 的三种失败路径：
  1. PPT 模式下单行 replace_placeholders 异常（stage=ppt）
  2. 图片模式下导出+降级双双失败（stage=fallback）
  3. 模板预检问题进入 BatchResult.precheck_issues（FONT_MISSING 探针）
"""

import sys
import os
import tempfile
import shutil

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

import pandas as pd
from pptx import Presentation
from pptx.util import Inches

from src.ppt_generator import PPTGenerator, BatchResult, FailedRow

PASS = 0
FAIL = 0


def check(desc, cond):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  [PASS] {desc}")
    else:
        FAIL += 1
        print(f"  [FAIL] {desc}")


def build_template(path, font_name=None):
    """构造带 ph_姓名 / ph_金额 占位符的最小模板"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式

    box1 = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    box1.name = "ph_姓名"
    box1.text_frame.text = "姓名占位"
    box1.text_frame.word_wrap = True  # 消除 python-pptx 默认 wrap=none 触发的 NO_WRAP

    box2 = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    box2.name = "ph_金额"
    box2.text_frame.text = "金额占位"
    box2.text_frame.word_wrap = True

    if font_name:
        for box in (box1, box2):
            for para in box.text_frame.paragraphs:
                for run in para.runs:
                    run.font.name = font_name

    prs.save(path)


def build_data(path):
    """4 行数据，其中第 3 行（index=2）为毒行"""
    df = pd.DataFrame({
        "姓名": ["张三", "李四", "坏数据", "王五"],
        "金额": ["1000", "2000", "3000", "4000"],
    })
    df.to_excel(path, index=False)


def make_generator(template_path, data_path):
    gen = PPTGenerator()
    assert gen.load_template(template_path), "模板加载失败"
    assert gen.load_data(data_path), "数据加载失败"
    rules = gen.auto_match_placeholders()
    assert rules, "自动匹配失败"
    return gen


def main():
    tmpdir = tempfile.mkdtemp(prefix="batch_fault_tol_")
    try:
        template = os.path.join(tmpdir, "tpl.pptx")
        data = os.path.join(tmpdir, "data.xlsx")
        build_template(template)
        build_data(data)

        print("(1) PPT 模式：毒行失败被记录，其余行成功")
        out1 = os.path.join(tmpdir, "out_ppt")
        gen = make_generator(template, data)
        orig_replace = gen.template_processor.replace_placeholders

        def poison_replace(data_map):
            if data_map.get("ph_姓名") == "坏数据":
                raise ValueError("模拟毒行数据异常")
            return orig_replace(data_map)

        gen.template_processor.replace_placeholders = poison_replace
        messages = []
        result = gen.batch_generate(out1, progress_callback=lambda c, t, m: messages.append(m),
                                    generate_images=False)

        check("返回 BatchResult", isinstance(result, BatchResult))
        check("total = 4", result.total == 4)
        check("成功 3 行", result.success_count == 3)
        check("失败 1 行", len(result.failed_rows) == 1)
        if result.failed_rows:
            f = result.failed_rows[0]
            check("失败行是 index=2（Excel 第 4 行）", f.index == 2 and f.row_number == 4)
            check("失败行姓名为坏数据", f.name == "坏数据")
            check("失败阶段为 ppt", f.stage == "ppt")
            check("错误信息包含模拟异常", "模拟毒行" in f.error)
        check("int() 兼容旧用法", int(result) == 3)
        check("成功行文件存在", os.path.exists(os.path.join(out1, "1_张三.pptx"))
              and os.path.exists(os.path.join(out1, "4_王五.pptx")))
        check("毒行无残留文件", not os.path.exists(os.path.join(out1, "3_坏数据.pptx")))
        check("进度回调收到失败通知", any("失败" in m for m in messages))

        print("(2) row_indices 重试毒行（去掉毒桩后单独重跑）")
        gen.template_processor.replace_placeholders = orig_replace
        failed_idx = [f.index for f in result.failed_rows]
        result2 = gen.batch_generate(os.path.join(tmpdir, "out_ppt"),
                                     generate_images=False, row_indices=failed_idx)

        check("重试 total = 1", result2.total == 1)
        check("重试成功 1 行", result2.success_count == 1 and not result2.failed_rows)
        check("重试产物存在", os.path.exists(os.path.join(outdir(tmpdir, "out_ppt"), "3_坏数据.pptx")))

        print("(3) 图片模式：导出与降级 PPT 双失败（stage=fallback）")
        out3 = os.path.join(tmpdir, "out_img")
        gen3 = make_generator(template, data)
        orig_save = gen3.template_processor.save_presentation

        def poison_save(path):
            # 毒行的临时 PPT 保存与降级 PPT 保存都失败
            # （注意不能用 RuntimeError——那是"Office 缺失"的整批中止信号）
            if getattr(poison_save, "last_replace", None) == "坏数据":
                raise ValueError("模拟保存失败")
            return orig_save(path)

        orig_replace3 = gen3.template_processor.replace_placeholders

        def tracking_replace(data_map):
            poison_save.last_replace = data_map.get("ph_姓名")
            return orig_replace3(data_map)

        gen3.template_processor.replace_placeholders = tracking_replace
        gen3.template_processor.save_presentation = poison_save
        result3 = gen3.batch_generate(out3, generate_images=True)

        check("图片模式成功 3 行", result3.success_count == 3)
        check("图片模式失败 1 行", len(result3.failed_rows) == 1)
        if result3.failed_rows:
            check("失败阶段为 fallback", result3.failed_rows[0].stage == "fallback")

        print("(4) 模板预检：缺失字体进入 precheck_issues")
        tpl_missing_font = os.path.join(tmpdir, "tpl_font.pptx")
        build_template(tpl_missing_font, font_name="NoSuchFont-XYZ-123")
        out4 = os.path.join(tmpdir, "out_precheck")
        gen4 = make_generator(tpl_missing_font, data)
        result4 = gen4.batch_generate(out4, generate_images=False)

        check("预检问题非空", len(result4.precheck_issues) > 0)
        codes = {issue.code for issue in result4.precheck_issues}
        check("包含 FONT_MISSING", "FONT_MISSING" in codes)
        check("预检不影响生成（4 行成功）", result4.success_count == 4)
        check("健康模板预检为空", not make_generator(template, data)
              .batch_generate(os.path.join(tmpdir, "out_clean"), generate_images=False).precheck_issues)

        print("(5) row_indices 指向不存在的行为空任务")
        result5 = gen4.batch_generate(out4, generate_images=False, row_indices=[999])
        check("不存在行 total=0", result5.total == 0 and result5.success_count == 0)

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


def outdir(tmpdir, name):
    return os.path.join(tmpdir, name)


if __name__ == "__main__":
    sys.exit(main())
