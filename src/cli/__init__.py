# -*- coding: utf-8 -*-
"""
PosterAgain 命令行入口（无人值守批量生成）

复用 GUI 同一条管线（PPTGenerator：加载模板 → 加载数据 → 匹配 → 批量生成），
因此容错（单行失败跳过并汇报）、模板预检、恢复式复用、内置/Office 引擎选择、
承保趸期数据预处理等行为与界面一致。不启动 GUI，可在脚本/CI/打包后的 exe 里调用：

    python main.py --template 模板.pptx --data 数据.xlsx --output 输出目录
    python cli.py  --template 模板.pptx --data 数据.csv  --output out --images
    PosterAgain.exe --template 模板.pptx --data 数据.xlsx --output out --preview

匹配规则来源（后者覆盖前者）：
  1. 自动匹配：占位符去掉 ph_ 前缀后与数据列名对应
  2. --preset 预设文件：界面「占位符配置」导出的 .pptcfg（matching_rules 列表，
     视为完整规则集，替换自动结果）；或界面「简化配置」导出的 gui_settings
     （text_additions / dundian_mappings / chengbao_mappings，叠加在自动结果上）
  3. --map 占位符=数据列、--add-prefix / --add-suffix 占位符=文字（可重复，叠加）

退出码：0 全部成功；1 有失败行（或预览时没有任何匹配）；2 致命错误（文件缺失、
无匹配规则、Office 引擎不可用等）。失败明细打印到 stdout 并写入日志。
没匹配到数据列的占位符会保留模板原文，生成前会明确警告。

说明：旧文档曾规划 --threshold（模糊匹配阈值），但匹配器是精确匹配、并无阈值
概念，故不提供该参数；--preview 只展示匹配结果不生成文件。
"""

import argparse
import json
import logging
import os
import sys

logger = logging.getLogger(__name__)

# main.py 据此判断是否走 CLI（出现任一 CLI 参数即进入命令行模式，否则启动 GUI）
CLI_FLAGS = (
    '--template', '-t', '--data', '-d', '--output', '-o', '--config',
    '--sheet', '--images', '--image-format', '--image-quality', '--engine',
    '--prefix', '--suffix', '--rows', '--preview', '--quiet', '-q',
    '--preset', '--map', '--add-prefix', '--add-suffix',
    '--help', '-h', '--version',
)


def is_cli_mode(argv) -> bool:
    """argv（不含程序名）里出现任一 CLI 参数 → 命令行模式"""
    return any(a in CLI_FLAGS for a in argv)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='PosterAgain',
        description='PPT 批量生成工具（命令行）：按数据把模板批量生成为个性化 PPT 或图片。',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            '示例:\n'
            '  生成 PPT:   PosterAgain --template 贺报.pptx --data 名单.xlsx --output out\n'
            '  生成图片:   PosterAgain -t 贺报.pptx -d 名单.csv -o out --images --image-quality 2\n'
            '  只看匹配:   PosterAgain -t 贺报.pptx -d 名单.xlsx --preview\n'
            '  指定引擎:   PosterAgain -t 贺报.pptx -d 名单.xlsx -o out --images --engine python\n'
            '  重跑失败行: PosterAgain -t 贺报.pptx -d 名单.xlsx -o out --rows 2,5\n'
            '  手工匹配:   PosterAgain -t 贺报.pptx -d 名单.xlsx -o out'
            ' --map ph_姓名=营销员姓名 --add-suffix ph_支公司=中支\n'
            '  用预设:     PosterAgain -t 贺报.pptx -d 名单.xlsx -o out --preset 承保.pptcfg\n'
            '\n'
            '退出码: 0 全部成功 / 1 有失败行 / 2 致命错误。'
        ),
    )
    p.add_argument('--template', '-t', required=True, metavar='FILE',
                   help='PPT 模板文件路径（.pptx）')
    p.add_argument('--data', '-d', required=True, metavar='FILE',
                   help='数据文件路径（.xlsx/.xls/.csv/.json）')
    p.add_argument('--output', '-o', metavar='DIR',
                   help='输出目录（生成文件时必填；--preview 时可省略）')
    p.add_argument('--config', metavar='FILE',
                   help='配置文件路径（.yaml/.json，可选；默认用运行时配置）')
    p.add_argument('--sheet', metavar='NAME',
                   help='Excel 工作表名（可选，默认第一个）')
    p.add_argument('--images', action='store_true',
                   help='输出图片而非 PPT（默认内置渲染引擎，无需 Office）')
    p.add_argument('--image-format', choices=['PNG', 'JPG', 'JPEG'], default='PNG',
                   help='图片格式（默认 PNG，仅 --images 时有效）')
    p.add_argument('--image-quality', type=float, default=1.0, metavar='N',
                   help='图片缩放倍数（1.0=原始, 2.0/3.0=增强，默认 1.0）')
    p.add_argument('--engine', choices=['python', 'office', 'auto'],
                   help='图片导出引擎：python 内置 / office 高保真(需Office) / auto；'
                        '默认取配置（缺省 python）')
    p.add_argument('--prefix', default='', metavar='STR',
                   help='输出文件名前缀（可选）')
    p.add_argument('--suffix', default='', metavar='STR',
                   help='输出文件名后缀（可选，加在扩展名之前）')
    p.add_argument('--rows', metavar='LIST',
                   help='只生成指定行（0 起行标签，逗号分隔，如 2,5；用于重跑失败行）')
    p.add_argument('--preset', metavar='FILE',
                   help='匹配预设文件（界面保存的 .pptcfg 或 .json）：含匹配规则与文字前后缀。'
                        'matching_rules 为列表时替换自动匹配结果，否则叠加')
    p.add_argument('--map', action='append', metavar='PH=COLUMN',
                   help='手工指定一条匹配规则（可重复），如 --map ph_姓名=营销员姓名')
    p.add_argument('--add-prefix', action='append', metavar='PH=TEXT',
                   help='给某占位符的替换文字加前缀（可重复），如 --add-prefix ph_标保=共计')
    p.add_argument('--add-suffix', action='append', metavar='PH=TEXT',
                   help='给某占位符的替换文字加后缀（可重复），如 --add-suffix ph_支公司=中支')
    p.add_argument('--preview', action='store_true',
                   help='只展示占位符与数据列的匹配结果，不生成文件')
    p.add_argument('--quiet', '-q', action='store_true',
                   help='不打印逐行进度（仍打印最终摘要）')
    p.add_argument('--version', action='version', version='PosterAgain 1.1.0')
    return p


def _attach_console_if_frozen() -> None:
    """打包成无控制台 exe（console=False）后从命令行调用时，附着到父控制台，
    使 stdout/stderr 可见。双击运行（无父控制台）时静默跳过，输出仍进日志文件。"""
    if not getattr(sys, 'frozen', False) or sys.platform != 'win32':
        return
    try:
        import ctypes
        if ctypes.windll.kernel32.AttachConsole(-1):  # ATTACH_PARENT_PROCESS
            sys.stdout = open('CONOUT$', 'w', encoding='utf-8', buffering=1)
            sys.stderr = open('CONOUT$', 'w', encoding='utf-8', buffering=1)
    except Exception as e:  # 附着失败不影响生成，日志文件仍有记录
        logger.debug(f"附着父控制台失败（输出仅进日志）: {e}")


def _load_config(path):
    """加载配置：--config 指定文件优先（缺失则报错），否则用运行时默认配置"""
    try:
        from src.config import ConfigManager
    except ImportError:
        from config_manager import ConfigManager
    if path:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"配置文件不存在: {path}")
        cm = ConfigManager()
        cm.config_file = os.path.abspath(path)
        cm.load_config()
        return cm
    return ConfigManager()


def _load_preset(path: str) -> dict:
    """读取匹配预设（界面保存的 .pptcfg / .json），返回
    {replace: 是否替换自动匹配, rules: {占位符: 数据列}, adds: {占位符: {prefix, suffix}}}。
    兼容三种写法：
    1. 界面「占位符配置」导出：顶层 matching_rules 列表（完整规则集 → 替换自动结果）
    2. 界面「简化配置」导出：gui_settings{text_additions, dundian_mappings, chengbao_mappings}
    3. 手写简写：matching_rules 字典 + text_additions 字典"""
    with open(path, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError('预设文件顶层应为 JSON 对象')

    def as_dict(v):
        return v if isinstance(v, dict) else {}

    rules, adds = {}, {}
    matching = data.get('matching_rules')
    replace = isinstance(matching, list)
    items = (matching if isinstance(matching, list)
             else [{'placeholder': k, 'column': v} for k, v in as_dict(matching).items()])
    for item in items:
        if not isinstance(item, dict):
            continue
        ph, col = item.get('placeholder'), item.get('column')
        if ph and col:
            rules[ph] = col
            ta = item.get('text_addition')
            if isinstance(ta, dict) and (ta.get('prefix') or ta.get('suffix')):
                adds[ph] = ta
    gui = as_dict(data.get('gui_settings'))
    for key in ('dundian_mappings', 'chengbao_mappings'):   # 趸期映射本质也是匹配规则
        for ph, col in as_dict(gui.get(key)).items():
            if ph and col:
                rules[ph] = col
    for ta_map in (gui.get('text_additions'), data.get('text_additions')):
        for ph, ta in as_dict(ta_map).items():
            if isinstance(ta, dict) and (ta.get('prefix') or ta.get('suffix')):
                adds[ph] = ta
    return {'replace': replace, 'rules': rules, 'adds': adds}


def _split_rule(text: str, option: str):
    """'ph_x=值' -> ('ph_x', '值')；格式不对抛 ValueError"""
    key, sep, value = text.partition('=')
    key, value = key.strip(), value.strip()
    if not sep or not key or not value:
        raise ValueError(f'{option} 应为 占位符=值 的形式: {text!r}')
    return key, value


def _parse_rows(text):
    """'2,5' -> [2, 5]；非法输入抛 ValueError"""
    if not text:
        return None
    rows = []
    for part in text.split(','):
        part = part.strip()
        if part:
            rows.append(int(part))
    return rows or None


def _preprocess_like_gui(gen) -> None:
    """与界面生成前的步骤保持一致：数据里没有「承保趸期数据」列时先试算
    （不传父窗口，不会弹对话框；需要人工录入的行留空，与界面无人值守时相同）"""
    try:
        data = gen.data_reader.data
        if data is not None and '承保趸期数据' not in data.columns:
            gen.data_reader._process_chengbao_term_data(parent_widget=None)
    except Exception as e:
        logger.warning(f"承保趸期数据预处理失败，将跳过此步骤: {e}")


def _apply_matching_setup(gen, args) -> list:
    """在自动匹配的基础上应用 --preset / --map / --add-*，返回需要告知用户的提示"""
    notes = []
    _preprocess_like_gui(gen)
    matcher = gen.exact_matcher
    data = gen.data_reader.data
    columns = set(data.columns) if data is not None else set()

    def add_rule(ph, col):
        if col not in columns:
            notes.append(f'匹配规则 {ph} -> {col}：数据里没有这一列，已忽略')
        else:
            gen.add_matching_rule(ph, col)

    if args.preset:
        preset = _load_preset(args.preset)
        if preset['replace']:
            # 界面「占位符配置」导出 = 完整规则集：整体替换（数据列缺失的规则会被跳过）
            res = matcher.import_matching_config(
                [{'placeholder': ph, 'column': col}
                 for ph, col in preset['rules'].items()])
            notes.extend(res.get('errors') or [])
            gen.matching_completed = bool(matcher.get_matching_rules())
        else:
            for ph, col in preset['rules'].items():
                add_rule(ph, col)
        for ph, ta in preset['adds'].items():
            matcher.set_text_addition_rule(ph, ta.get('prefix', ''), ta.get('suffix', ''))
    for item in (getattr(args, 'map', None) or []):
        add_rule(*_split_rule(item, '--map'))
    for item in (args.add_prefix or []):
        ph, text = _split_rule(item, '--add-prefix')
        matcher.set_text_addition_rule(ph, text, '')
    for item in (args.add_suffix or []):
        ph, text = _split_rule(item, '--add-suffix')
        matcher.set_text_addition_rule(ph, '', text)
    return notes


def _print_matching(gen, rules) -> None:
    """打印匹配结果（--preview 与生成前都展示），含每条的匹配规则与文字前后缀"""
    adds = gen.exact_matcher.get_all_text_addition_rules()
    print('\n占位符匹配结果:')
    if rules:
        for ph, col in rules.items():
            ta = adds.get(ph) or {}
            tag = ''
            if ta.get('prefix') or ta.get('suffix'):
                tag = (f"   [文字增加: 前缀 {ta.get('prefix', '')!r}"
                       f" 后缀 {ta.get('suffix', '')!r}]")
            print(f'  {ph}  <-  {col}{tag}')
    else:
        print('  （无匹配）')
    unmatched_ph = gen.exact_matcher.get_unmatched_placeholders()
    unmatched_col = gen.exact_matcher.get_unmatched_columns()
    if unmatched_ph:
        print(f'  未匹配占位符: {", ".join(unmatched_ph)}')
    if unmatched_col:
        print(f'  未使用数据列: {", ".join(unmatched_col)}')


def _print_summary(result, args) -> None:
    kind = '图片' if args.images else 'PPT'
    print('\n' + '=' * 50)
    print(f'批量生成{kind}完成: 成功 {result.success_count}/{result.total} 行')
    if result.degraded_rows:
        print(f'  降级输出 PPT（图片导出失败）{len(result.degraded_rows)} 行:')
        for fr in result.degraded_rows:
            print(f'    第 {fr.row_number} 行（{fr.name}）: {fr.error}')
    if result.failed_rows:
        print(f'  失败 {len(result.failed_rows)} 行:')
        for fr in result.failed_rows:
            print(f'    第 {fr.row_number} 行（{fr.name}）[{fr.stage}]: {fr.error}')
        print(f'  可用 --rows {",".join(str(fr.index) for fr in result.failed_rows)} 重跑失败行')
    if result.precheck_issues:
        print(f'  模板预检提示 {len(result.precheck_issues)} 项（近似渲染/缺失字体等，详见日志）')
    if result.cancelled:
        print('  （已中途取消，以上为已完成部分）')
    print(f'输出目录: {os.path.abspath(args.output)}')
    print('=' * 50)


def run(argv=None) -> int:
    """命令行主流程，返回退出码（0 成功 / 1 部分失败 / 2 致命错误）"""
    _attach_console_if_frozen()
    parser = build_parser()
    args = parser.parse_args(list(sys.argv[1:] if argv is None else argv))

    if not args.preview and not args.output:
        parser.error('生成文件需要 --output（或用 --preview 只看匹配结果）')
    try:
        rows = _parse_rows(args.rows)
    except ValueError:
        parser.error(f'--rows 格式错误（应为逗号分隔的整数）: {args.rows!r}')
    try:
        for item in (args.map or []):
            _split_rule(item, '--map')
        for item in (args.add_prefix or []):
            _split_rule(item, '--add-prefix')
        for item in (args.add_suffix or []):
            _split_rule(item, '--add-suffix')
    except ValueError as e:
        parser.error(str(e))
    if args.image_quality <= 0:
        parser.error('--image-quality 必须为正数')
    if args.preset and not os.path.isfile(args.preset):
        parser.error(f'预设文件不存在: {args.preset}')

    # 配置 / 引擎 / 字体策略
    try:
        cm = _load_config(args.config)
    except Exception as e:
        print(f'[错误] {e}', file=sys.stderr)
        return 2
    if args.engine:
        os.environ['POSTERAGAIN_IMAGE_ENGINE'] = args.engine
    try:
        from render_core import font_registry
        font_registry.set_policy(str(cm.get('image_generation.font_policy', 'auto') or 'auto'))
    except Exception as e:
        logger.debug(f"字体策略注入跳过: {e}")

    # 复用 GUI 同一条管线
    try:
        from src.ppt_generator import PPTGenerator
    except ImportError:
        from ppt_generator import PPTGenerator
    gen = PPTGenerator()

    if not os.path.isfile(args.template):
        print(f'[错误] 模板文件不存在: {args.template}', file=sys.stderr)
        return 2
    if not os.path.isfile(args.data):
        print(f'[错误] 数据文件不存在: {args.data}', file=sys.stderr)
        return 2
    if not gen.load_template(args.template):
        print(f'[错误] 无法加载模板: {args.template}', file=sys.stderr)
        return 2
    if not gen.load_data(args.data, args.sheet):
        print(f'[错误] 无法加载数据: {args.data}', file=sys.stderr)
        return 2

    gen.auto_match_placeholders()
    try:
        notes = _apply_matching_setup(gen, args)
    except (ValueError, OSError, json.JSONDecodeError) as e:
        print(f'[错误] 预设无效: {args.preset}: {e}', file=sys.stderr)
        return 2
    for line in notes:
        print(f'[提示] {line}')
    rules = gen.get_matching_rules()
    _print_matching(gen, rules)

    if args.preview:
        return 0 if rules else 1
    if not rules:
        print('[错误] 没有有效的匹配规则，无法生成（检查占位符与数据列名，'
              '或用 --map / --preset 指定）', file=sys.stderr)
        return 2
    unmatched = gen.exact_matcher.get_unmatched_placeholders()
    if unmatched:
        print(f'[警告] {len(unmatched)} 个占位符没有匹配到数据列，'
              f'输出中将保留模板原文: {", ".join(unmatched)}')

    def _progress(current, total, message):
        if not args.quiet:
            print(f'  [{current}/{total}] {message}', flush=True)

    try:
        result = gen.batch_generate(
            args.output,
            progress_callback=_progress,
            generate_images=args.images,
            image_format=args.image_format,
            image_quality=args.image_quality,
            row_indices=rows,
            filename_prefix=args.prefix,
            filename_suffix=args.suffix,
        )
    except RuntimeError as e:
        # 环境级致命错误（如 office 引擎但本机无 Office）
        print(f'[错误] {e}', file=sys.stderr)
        return 2
    except Exception as e:
        logger.exception('批量生成失败')
        print(f'[错误] 批量生成失败: {e}', file=sys.stderr)
        return 2

    _print_summary(result, args)
    if result.cancelled or result.failed_rows:
        return 1
    return 0


def main() -> None:
    sys.exit(run())


if __name__ == '__main__':
    main()
