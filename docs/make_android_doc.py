# -*- coding: utf-8 -*-
"""生成安卓端使用指南 docx（带模拟器实拍截图）"""
import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

SHOTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'shots')
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   'docs', '安卓端使用指南.docx')

ACCENT = RGBColor(0xC0, 0x00, 0x00)


def add_shot(doc, name, width=6.5, caption=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(os.path.join(SHOTS, name + '.png'), width=Cm(width))
    if caption:
        c = doc.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.add_run(caption)
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)


def h(doc, text, level=1):
    r = doc.add_heading(text, level=level)
    for run in r.runs:
        run.font.color.rgb = ACCENT if level == 1 else RGBColor(0, 0, 0)


def main():
    doc = Document()

    # 封面标题
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run('PosterAgain 安卓端使用指南')
    r.font.size = Pt(26)
    r.bold = True
    s = doc.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = s.add_run('批量生成个性化保险贺报 · 现场出图 · 一键转发')
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    add_shot(doc, '01_home', 5.0)

    # 一、简介
    h(doc, '一、简介')
    doc.add_paragraph(
        'PosterAgain 安卓版把 PPT 模板和 Excel 数据装进手机：选择模板与数据后，'
        '自动完成字段匹配并批量生成每个人的专属海报图片，可直接保存到相册或'
        '分享给客户。全部渲染在手机本地完成，不需要联网，也不需要安装 Office。')
    doc.add_paragraph('适合场景：营销员早会/现场接到承保或递交数据后，'
                      '即时生成贺报图片转发客户或朋友圈。')
    for line in (
        '支持的数据格式：Excel（.xlsx）、CSV、JSON',
        '支持的模板：.pptx（占位符写法见第六节）',
        '输出：PNG/JPG 图片，按「序号_姓名」命名（与电脑端一致）'):
        doc.add_paragraph(line, style='List Bullet')

    # 二、快速上手
    h(doc, '二、快速上手（三步出图）')
    h(doc, '第 1 步：选择模板和数据', 2)
    doc.add_paragraph(
        '打开 App 进入「任务」首页，依次点击两张卡片：先选 PPT 模板文件，'
        '再选数据文件（可从微信/QQ 接收后保存的文件中选择，也可通过系统'
        '文件选择器浏览）。')
    add_shot(doc, '01_home', 6.0, '图 1  首页：选择模板与数据')
    doc.add_paragraph('选好模板后，卡片会显示占位符数量：')
    add_shot(doc, '02_home_picked', 6.0, '图 2  模板已选中（显示 4 个占位符）')
    doc.add_paragraph('数据选好后，「开始生成」按钮会显示将生成的张数：')
    add_shot(doc, '03_home_ready', 6.0, '图 3  就绪状态（10 行数据 → 10 张海报）')
    doc.add_paragraph(
        '提示：「最近使用」区域会记录你用过的模板，点击可直接复用，'
        '不必每次找文件。')

    h(doc, '第 2 步：确认字段绑定', 2)
    doc.add_paragraph(
        '点击「开始生成」进入字段绑定页。系统会自动匹配模板占位符与数据列：'
        '匹配上的显示对应列名；未匹配的显示「未绑定」，这些字段将保留模板'
        '原样输出。')
    add_shot(doc, '04_after_start', 6.0, '图 4  字段绑定页（1 个字段未绑定）')
    doc.add_paragraph('点击任意字段可手动设置，三种方式：')
    for line in (
        '数据列：从数据的全部列中选择一列（适合列名不完全一致的情况）',
        '固定值：输入一个固定文字，应用到所有行（如「个险」「趸期」等模板固定文案）',
        '前缀/后缀：给该字段的值拼接文字（如姓名前加「客户：」、金额后加「元」）'):
        doc.add_paragraph(line, style='List Bullet')
    add_shot(doc, '05_binding_dialog', 6.0, '图 5  字段设置对话框')
    doc.add_paragraph(
        '配置会自动记忆：下次使用同一模板时自动套用，界面会提示'
        '「已套用上次保存的字段配置」。模板内容改动后旧配置自动失效。')

    h(doc, '第 3 步：批量渲染与转发', 2)
    doc.add_paragraph(
        '点击「开始渲染」，逐行生成全部海报（每张约 0.3~3.5 秒，视模板大小）。'
        '完成后展示全部结果的画廊：')
    add_shot(doc, '06_render_done', 6.0, '图 6  渲染完成（10 张海报画廊）')
    doc.add_paragraph('点击任意一张进入大图预览：')
    add_shot(doc, '07_preview', 6.0, '图 7  预览页（左右滑动切换）')
    for line in (
        '保存到相册：单张保存（相册 Pictures/PosterAgain 目录）',
        '分享：调起微信/QQ 等系统分享',
        '设为壁纸 / 删除',
        '渲染页底部「全部保存到相册」「分享」批量操作'):
        doc.add_paragraph(line, style='List Bullet')

    # 三、模板库
    h(doc, '三、模板库')
    doc.add_paragraph(
        '底部导航「模板」页集中管理用过的模板：显示使用次数与最近使用时间，'
        '点击复用，长按或按钮删除。')
    add_shot(doc, '08_templates', 6.0, '图 8  模板库')

    # 四、设置
    h(doc, '四、设置')
    add_shot(doc, '09_settings', 6.0, '图 9  设置页')
    for line in (
        '字体方案：内置思源黑体/宋体（开源授权），无需额外安装字体',
        '图片格式：PNG（无损）/ JPG（体积小）',
        '导出尺寸：原始 / 增强档位',
        '保存位置：默认相册 Pictures/PosterAgain'):
        doc.add_paragraph(line, style='List Bullet')

    # 五、常见问题
    h(doc, '五、常见问题')
    faqs = (
        ('生成的图片和电脑上 PowerPoint 打开的效果有细微差别？',
         '渲染引擎已对文字发光、渐变、阴影等效果按 PowerPoint 逐项校准，'
         '正常差异仅为抗锯齿级别。若模板使用了倒影、3D、艺术字变形等'
         '少数不支持的效果，会自动按近似方式渲染并在日志中提示。'),
        ('模板用了特殊字体，手机上会变吗？',
         '手机端内置思源黑体与思源宋体，常用中文字体直接可渲。若模板使用'
         '系统未安装的商业字体（如字魂系列），会回退到内置字体——建议模板'
         '制作时使用思源系列或在电脑端嵌入字体后使用。'),
        ('数据里的金额带逗号（如 88,888）有问题吗？',
         '没有问题。请确保 CSV 文件中该字段加了引号（Excel 另存的 CSV 会'
         '自动处理），xlsx 则无此顾虑。'),
        ('某一行数据有错误怎么办？',
         '渲染按行进行，单行数据异常不影响其他行；该行产物文件名会以'
         '「序号_output」标识，便于核对。'),
        ('支持哪些占位符写法？',
         '见下节。推荐使用「替代文本」方式，对模板画面零影响。'),
    )
    for q, a in faqs:
        p = doc.add_paragraph()
        r = p.add_run('Q：' + q)
        r.bold = True
        doc.add_paragraph('A：' + a)

    # 六、模板制作规范
    h(doc, '六、模板制作规范（给设计师）')
    doc.add_paragraph('占位符三种写法（优先级从高到低）：')
    for line in (
        '替代文本（推荐）：右键形状 → 编辑替代文本 → 填 ph_字段名（如 ph_姓名）',
        '形状名称：选择窗格中把形状命名为 ph_字段名',
        '文本标记：正文中直接写 {{字段名}}（行内替换，保留段落格式）'):
        doc.add_paragraph(line, style='List Bullet')
    doc.add_paragraph(
        '字段名与数据列名一致时自动匹配；含「姓名」的列会用作输出文件命名。'
        '文字效果方面支持发光、阴影、渐变填充、描边等常用效果；'
        '建议背景大图使用 JPEG 以减少内存占用。')

    doc.save(OUT)
    print('saved:', OUT, os.path.getsize(OUT) // 1024, 'KB')


if __name__ == '__main__':
    sys.exit(main())
