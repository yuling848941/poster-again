"""
字体注册表：模板 typeface -> 字体文件路径 + 排版校准常数
================================================================
跨平台字体解析的单一事实来源（Windows / macOS / Android 共用）。
本模块自包含（仅标准库），安卓端 Chaquopy 嵌入引擎核心时连同
resources/fonts/ 一起带走即可，无需改动。

字体来源（按策略排序）：
- 捆绑字体 resources/fonts/：思源黑体 + 思源宋体（CN 子集，SIL OFL 1.1，
  可合法内嵌分发；微软雅黑为微软专有字体，禁止内嵌）
- 系统字体：Windows Fonts 路径表（自 pillow_image_exporter 迁移，保持原行为）

解析策略（set_policy() 或环境变量 POSTERAGAIN_FONT_POLICY）：
- auto（默认）：Windows 系统优先（现有渲染像素级不变），其他平台捆绑优先
- system_first / bundled_first：显式指定

排版校准常数与物理字体绑定（按文件路径查询 metrics_for）：
- 系统中文体字体：沿用 COM 对照微软雅黑实测值（原引擎常数等价改写，勿改）
- 思源黑体/宋体：experiments/calibrate_sourcehan.py COM 对照校准（2026-08-18，
  首行基线为仿射模型，黑体/宋体共享垂直度量，残差 ≤2px@q2）
"""

import io
import logging
import os
import sys
from typing import NamedTuple, Optional, Set

logger = logging.getLogger(__name__)


class FontMetrics(NamedTuple):
    """单物理字体的排版校准常数（COM 对照 PowerPoint 实测，experiments/calibrate_sourcehan.py）
    首行基线为仿射模型：baseline_em = (baseline_a + baseline_g × lnSpc)，乘字号px；
    纯比例模型是 a=0 的特例（msyh 历史 0.72×行高 ≡ a=0, g=0.72×1.25=0.9）"""
    line_height_ratio: float   # 单倍行距 / 字号（line_h = ratio×lnSpc×字号px）
    baseline_a: float          # 首行基线仿射截距（em）
    baseline_g: float          # 首行基线仿射 lnSpc 系数（em）
    strike_pos_em: float       # 删除线位于基线上方 em 数
    calibrated: bool           # 是否已 COM 对照校准


class FontSpec(NamedTuple):
    path: str          # 字体文件绝对路径
    metrics: FontMetrics
    source: str        # 'bundled' | 'system'
    family: str        # 'sans' | 'serif' | ''（默认链回退）


# 微软雅黑 COM 逐像素校准值（原 pillow_image_exporter 常数的等价改写，勿改）
# 历史公式 0.72 × line_h = 0.72×1.25×L×F ≡ (0 + 0.9×L)×F
METRICS_MS_YAHEI = FontMetrics(1.25, 0.0, 0.9, 0.245, calibrated=True)
# 思源黑体：experiments/calibrate_sourcehan.py COM 对照校准（2026-08-18，
# 仿射拟合残差 ≤2px@q2；删除线沿用雅黑 0.245em）
METRICS_SOURCE_HAN_SANS = FontMetrics(1.1942, 0.1852, 0.7803, 0.245, calibrated=True)
# 思源宋体：COM 对照实测与黑体逐项一致（Adobe 同族共享垂直度量设计，
# 独立校准轮验证残差同 ≤2px@q2；字形差异已单独核实）
METRICS_SOURCE_HAN_SERIF = FontMetrics(1.1942, 0.1852, 0.7803, 0.245, calibrated=True)
METRICS_SOURCE_HAN = METRICS_SOURCE_HAN_SANS  # 兼容别名

# ---------------- 系统字体表（原 _FONT_FILES，行为保持不变） ----------------
_WIN = os.environ.get('WINDIR', r'C:\Windows')
SYSTEM_FONT_FILES = {
    ('微软雅黑', False): os.path.join(_WIN, 'Fonts', 'msyh.ttc'),
    ('微软雅黑', True): os.path.join(_WIN, 'Fonts', 'msyhbd.ttc'),
    ('Microsoft YaHei', False): os.path.join(_WIN, 'Fonts', 'msyh.ttc'),
    ('Microsoft YaHei', True): os.path.join(_WIN, 'Fonts', 'msyhbd.ttc'),
    ('宋体', False): os.path.join(_WIN, 'Fonts', 'simsun.ttc'),
    ('宋体', True): os.path.join(_WIN, 'Fonts', 'simsun.ttc'),
    ('黑体', False): os.path.join(_WIN, 'Fonts', 'simhei.ttf'),
    ('黑体', True): os.path.join(_WIN, 'Fonts', 'simhei.ttf'),
    ('楷体', False): os.path.join(_WIN, 'Fonts', 'simkai.ttf'),
    ('楷体', True): os.path.join(_WIN, 'Fonts', 'simkai.ttf'),
}

# ---------------- 捆绑字体族（resources/fonts/） ----------------
_BUNDLED_FAMILIES = {
    'sans': {
        'names': ('思源黑体', '思源黑体 cn', 'source han sans cn', 'source han sans sc',
                  'source han sans', 'noto sans cjk sc',
                  '微软雅黑', '微软雅黑 light', 'microsoft yahei', 'microsoft yahei ui',
                  'microsoft yahei light', '黑体', 'simhei', '等线', 'dengxian'),
        'regular': 'SourceHanSansCN-Regular.otf',
        'bold': 'SourceHanSansCN-Bold.otf',
        'metrics': METRICS_SOURCE_HAN_SANS,
    },
    'serif': {
        'names': ('思源宋体', '思源宋体 cn', 'source han serif cn', 'source han serif sc',
                  'source han serif', 'noto serif cjk sc',
                  '宋体', '新宋体', 'simsun', 'nsimsun', '仿宋', 'fangsong'),
        'regular': 'SourceHanSerifCN-Regular.otf',
        'bold': 'SourceHanSerifCN-Bold.otf',
        'metrics': METRICS_SOURCE_HAN_SERIF,
    },
}

# 系统表键 -> 该文件存在时可提供的字体名（供预检/字体检查用）
_SYSTEM_NAME_ALIASES = {
    'msyh.ttc': ('微软雅黑', 'microsoft yahei', 'microsoft yahei ui',
                 '微软雅黑 light', 'microsoft yahei light'),
    'msyhbd.ttc': ('微软雅黑', 'microsoft yahei', 'microsoft yahei ui'),
    'simsun.ttc': ('宋体', 'simsun', '新宋体', 'nsimsun'),
    'simhei.ttf': ('黑体', 'simhei'),
    'simkai.ttf': ('楷体', 'kaiti', '楷体_gb2312', 'stkaiti'),
    'fangsong.ttf': ('仿宋', 'fangsong', '仿宋_gb2312'),
}

_policy: Optional[str] = None
_uncalibrated_warned = False

# ---------------- PPTX 内嵌字体（embeddedFontLst）----------------
# PowerPoint 语义：系统字体 -> 文档内嵌字体 -> 替换字体。
# 引擎同序：系统/捆绑都未命中时才用嵌入（Windows 已安装字体行为零回归；
# 安卓无系统字体，嵌入字体即"真字体"，与 PowerPoint 显示一致）
_embedded: dict = {}          # (typeface.strip().lower(), bold) -> FontSpec
_embedded_tmp: list = []      # 本会话解出的临时字体文件（切换文档时清理）
_embedded_src_key = None      # 当前注册来源标识（避免同文档重复解包）


def clear_embedded_fonts():
    """清空内嵌字体注册（删除临时文件）"""
    global _embedded_src_key
    _embedded.clear()
    for p in _embedded_tmp:
        try:
            os.remove(p)
        except OSError:
            pass
    _embedded_tmp.clear()
    _embedded_src_key = None


def _deobfuscate_odttf(data: bytes, part_name: str) -> bytes:
    """ODTTF 混淆字体：前 32 字节与部件名 GUID 字节（16B，重复两次）异或"""
    import re as _re
    m = _re.search(r'([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})',
                   part_name)
    if not m:
        return data
    guid = bytes.fromhex(m.group(1).replace('-', ''))
    head = bytearray(data[:32])
    for i in range(32):
        head[i] ^= guid[i % 16]
    return bytes(head) + data[32:]


_FONT_SIGS = (b'\x00\x01\x00\x00', b'OTTO', b'true', b'ttcf')


def _decode_fntdata(data: bytes, part_name: str) -> bytes:
    """PowerPoint 嵌入字体数据解码：
    1) 微软压缩格式（4B 解压大小 + 4B 压缩大小 + raw-deflate/zlib 流）
    2) ODTTF 混淆（前 32 字节 GUID 异或）
    返回字体原始字节；无法识别返回原数据（调用方校验签名）"""
    import struct
    import zlib
    if data[:4] in _FONT_SIGS:
        return data
    if len(data) > 8:
        try:
            raw = zlib.decompress(data[8:], -15)
            if raw[:4] in _FONT_SIGS:
                return raw
        except Exception:
            pass
        try:
            raw = zlib.decompress(data[8:])
            if raw[:4] in _FONT_SIGS:
                return raw
        except Exception:
            pass
    return _deobfuscate_odttf(data, part_name)


def load_embedded_fonts(pptx_source) -> int:
    """
    从 pptx（路径或字节）解析 embeddedFontLst，提取 fntdata 注册为嵌入字体。
    返回注册的 (typeface, bold) 条数；无嵌入字体返回 0。同文档重复调用为空操作。
    """
    global _embedded_src_key
    import zipfile
    if isinstance(pptx_source, (bytes, bytearray)):
        key = ('bytes', len(pptx_source))
    else:
        key = ('path', os.path.abspath(str(pptx_source)))
    if key == _embedded_src_key:
        return len(_embedded)
    clear_embedded_fonts()
    _embedded_src_key = key
    try:
        z = zipfile.ZipFile(io.BytesIO(bytes(pptx_source))
                            if isinstance(pptx_source, (bytes, bytearray))
                            else str(pptx_source))
    except Exception as e:
        logger.debug(f"打开 pptx 提取嵌入字体失败: {e}")
        return 0
    from lxml import etree
    P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    try:
        pres = etree.fromstring(z.read('ppt/presentation.xml'))
        rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        # .rels 的 Relationship 元素是 package/2006/relationships 命名空间，
        # 按 localname 取避免硬编码
        rid2target = {rel.get('Id'): rel.get('Target')
                      for rel in rels.iter()
                      if etree.QName(rel).localname == 'Relationship'}
        lst = pres.find(f'{{{P}}}embeddedFontLst')
        if lst is None:
            return 0
        import tempfile
        n = 0
        for ef in lst.iter(f'{{{P}}}embeddedFont'):
            font_el = ef.find(f'{{{P}}}font')
            typeface = font_el.get('typeface') if font_el is not None else None
            if not typeface:
                continue
            for style in ('regular', 'bold', 'italic', 'boldItalic'):
                el = ef.find(f'{{{P}}}{style}')
                if el is None:
                    continue
                target = rid2target.get(el.get(f'{{{R}}}id'))
                if not target:
                    continue
                part = 'ppt/' + target.lstrip('/')
                if part not in z.namelist():
                    continue
                data = _decode_fntdata(z.read(part), part)
                if data[:4] not in _FONT_SIGS:
                    logger.warning(
                        f"嵌入字体 {typeface!r}({style}) 为 PowerPoint 子集化压缩格式"
                        f"（微软私有，LibreOffice 亦不支持），已跳过——将按系统/捆绑字体回退；"
                        f"如需该字体保真渲染请用 COM 导出引擎")
                    continue
                fd, tmp = tempfile.mkstemp(suffix='.ttf')
                with os.fdopen(fd, 'wb') as f:
                    f.write(data)
                _embedded_tmp.append(tmp)
                bold = style in ('bold', 'boldItalic')
                _embedded[(typeface.strip().lower(), bold)] = \
                    FontSpec(tmp, METRICS_MS_YAHEI, 'embedded', '')
                n += 1
        if n:
            logger.info(f"已从文档提取 {len(_embedded)} 个嵌入字体条目（{n} 个字体文件）")
        return len(_embedded)
    except Exception as e:
        logger.warning(f"解析嵌入字体清单失败（忽略）: {e}")
        return 0


def _lookup_embedded(typeface: Optional[str], bold: bool) -> Optional[FontSpec]:
    for tf, b in _typeface_variants(typeface, bold):
        spec = _embedded.get((tf.strip().lower(), b))
        if spec is not None and os.path.exists(spec.path):
            return spec
    return None


# ---------------- 捆绑目录解析 ----------------

def bundled_font_dir() -> Optional[str]:
    """返回第一个存在的捆绑字体目录（env 覆盖 > frozen > 源码仓库）"""
    candidates = []
    env = os.environ.get('POSTERAGAIN_FONTS_DIR')
    if env:
        candidates.append(env)
    if getattr(sys, 'frozen', False):
        candidates.append(os.path.join(sys._MEIPASS, 'resources', 'fonts'))
        candidates.append(os.path.join(os.path.dirname(sys.executable),
                                       'resources', 'fonts'))
    here = os.path.dirname(os.path.abspath(__file__))
    # 逐级探测 resources/fonts：同级（安卓 Chaquopy/嵌入布局：包内自带）+
    # 向上 1~3 级（仓库布局 src/render_core -> 仓库根、隔离目录布局）
    candidates.append(os.path.join(here, 'resources', 'fonts'))
    for up in (1, 2, 3):
        candidates.append(os.path.normpath(os.path.join(
            here, *(['..'] * up), 'resources', 'fonts')))
    for d in candidates:
        if os.path.isdir(d):
            return d
    return None


def _bundled_file(filename: str) -> Optional[str]:
    d = bundled_font_dir()
    if d is None:
        return None
    p = os.path.join(d, filename)
    return p if os.path.exists(p) else None


# ---------------- 策略 ----------------

def set_policy(policy: str) -> None:
    """设置解析策略（应用层启动时调用；None/非法值回退 auto）"""
    global _policy
    if policy in ('auto', 'system_first', 'bundled_first'):
        _policy = policy
    else:
        logger.warning(f"非法字体策略 {policy!r}，回退 auto")
        _policy = 'auto'


def effective_policy() -> str:
    """实际生效策略：auto 在 Windows 上系统优先（渲染零回归），其他平台捆绑优先"""
    p = _policy
    if p is None:
        p = os.environ.get('POSTERAGAIN_FONT_POLICY', 'auto')
    if p == 'auto':
        return 'system_first' if sys.platform == 'win32' else 'bundled_first'
    return p


# ---------------- 解析 ----------------

def _typeface_variants(typeface: Optional[str], bold: bool):
    """与原 _font_path_for 一致的候选顺序：精确 -> 非粗体 -> 去 -Bold 后缀"""
    if not typeface:
        return []
    return [(typeface, bold), (typeface, False),
            (typeface.replace('-Bold', '').replace('-bold', ''), bold)]


# ---------------- 本机扩展字体目录（静默增强，隐藏能力）----------------
# 部分办公软件会在本机缓存其下载的字体（云字体为会员权益，授权仅限该软件内
# 使用）。引擎扫描这些目录纯属"有则引用、无则回退"的静默增强：不依赖、不
# 宣传、日志 debug 级。业务上要求字体一致时，正确做法是模板使用可分发字体
# 或购买商用授权。可用 POSTERAGAIN_WPS_FONTS_DIR 指定目录。
_extra_font_table: Optional[dict] = None   # (name.lower(), bold) -> path


def _wps_font_dirs():
    import glob as _glob
    env = os.environ.get('POSTERAGAIN_WPS_FONTS_DIR')
    roots = [env] if env else [
        os.path.join(os.environ.get('APPDATA', ''), 'kingsoft', 'office6', 'docerFonts'),
        os.path.join(os.environ.get('APPDATA', ''), 'kingsoft', 'hiddenfonts'),
    ]
    dirs = []
    for r in roots:
        if r and os.path.isdir(r):
            dirs.append(r)
            sub = os.path.join(r, 'fontsdk', 'cache', 'download')
            if os.path.isdir(sub):
                dirs.append(sub)
    return dirs


def _font_display_names(path: str):
    """返回 (家族名列表, 是否粗体)。粗体用 OS/2.fsSelection 位判定（可靠），
    名字猜测会误判（Adobe 系 Bold 的家族名常不含 Bold 字样）"""
    names = []
    bold = False
    try:
        from fontTools.ttLib import TTFont
        ft = TTFont(path, fontNumber=0, lazy=True)
        try:
            if 'OS/2' in ft:
                bold = bool(ft['OS/2'].fsSelection & 0x20)
            elif 'head' in ft:
                bold = bool(ft['head'].macStyle & 1)
            for nid in (1, 4):          # family / full name
                for rec in ft['name'].names:
                    if rec.nameID == nid:
                        s = rec.toUnicode()
                        if s and s not in names:
                            names.append(s)
        finally:
            ft.close()
    except Exception:
        pass
    if not names:
        try:
            from PIL import ImageFont
            names.append(ImageFont.truetype(path, 8).getname()[0])
        except Exception:
            return [], False
    return names, bold


def _scan_wps_fonts() -> dict:
    """扫描本机扩展字体目录，读字体内部名建表（惰性一次）"""
    global _extra_font_table
    if _extra_font_table is not None:
        return _extra_font_table
    table = {}
    import glob as _glob
    for d in _wps_font_dirs():
        for p in (_glob.glob(os.path.join(d, '**', '*.tt[fc]'), recursive=True)
                  + _glob.glob(os.path.join(d, '**', '*.otf'), recursive=True)):
            names, bold = _font_display_names(p)
            for name in names:
                table.setdefault((name.strip().lower(), bold), p)
    if table:
        logger.debug(f"扩展字体目录扫描到 {len(table)} 个条目")
    _extra_font_table = table
    return table


def _edit_distance_le1(a: str, b: str) -> bool:
    """中文名近名判定：编辑距离 <= 1（模板笔误容错，如 疯藏/馆藏）"""
    if a == b:
        return True
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    if la == lb:
        diff = sum(1 for x, y in zip(a, b) if x != y)
        return diff <= 1
    # 长度差 1：短的是否为长的子序列（单字符插入/删除）
    s, l = (a, b) if la < lb else (b, a)
    i = 0
    for ch in l:
        if i < len(s) and s[i] == ch:
            i += 1
    return i == len(s)


def _lookup_wps(typeface: Optional[str], bold: bool) -> Optional[FontSpec]:
    """本机扩展字体查找：精确命中 -> 近名容错（编辑距离<=1 且唯一候选才采用）"""
    table = _scan_wps_fonts()
    if not table:
        return None
    for tf, b in _typeface_variants(typeface, bold):
        hit = table.get((tf.strip().lower(), b))
        if hit is not None:
            return FontSpec(hit, METRICS_MS_YAHEI, 'local', '')
    # 近名容错（复刻 WPS 行为）：编辑距离<=1 且唯一候选才采用
    # （regular/bold 同名条目算同一候选，按请求的 bold 取变体）
    key = (typeface or '').strip().lower()
    if not key:
        return None
    near = {}
    for (name, b), p in table.items():
        if _edit_distance_le1(key, name):
            near.setdefault(name, {})[b] = p
    if len(near) == 1:
        name = next(iter(near))
        variants = near[name]
        path = variants.get(bold) or variants.get(False)
        if path:
            logger.debug(
                f"字体 {typeface!r} 未直接命中，采用本机近名字体 {name!r}"
                f"（编辑距离<=1）")
            return FontSpec(path, METRICS_MS_YAHEI, 'local', '')
    return None


def _lookup_system(typeface: Optional[str], bold: bool) -> Optional[FontSpec]:
    for tf, b in _typeface_variants(typeface, bold):
        path = SYSTEM_FONT_FILES.get((tf, b))
        if path and os.path.exists(path):
            return FontSpec(path, METRICS_MS_YAHEI, 'system', '')
    return None


def _lookup_bundled(typeface: Optional[str], bold: bool) -> Optional[FontSpec]:
    for tf, b in _typeface_variants(typeface, bold):
        name = tf.strip().lower()
        for fam, info in _BUNDLED_FAMILIES.items():
            if name in info['names']:
                filename = info['bold'] if (b and _bundled_file(info['bold'])) \
                    else info['regular']
                path = _bundled_file(filename)
                if path:
                    return FontSpec(path, info['metrics'], 'bundled', fam)
    return None


def resolve(typeface: Optional[str], bold: bool = False) -> FontSpec:
    """
    typeface -> FontSpec；永不抛异常，无任何命中时按默认链回退并告警。
    顺序：typeface 变体在【两个来源】都查完，才轮到微软雅黑默认链
    （否则 system_first 下系统雅黑会抢先命中，捆绑字体直接命中被短路）。
    系统与捆绑均未命中时尝试文档内嵌字体（PowerPoint 同语义：系统 -> 嵌入 -> 替换）。
    最终回退与原 _font_path_for 语义一致：微软雅黑(按策略查两来源) -> msyh 路径字面值。
    """
    global _uncalibrated_warned
    # 本机字体层 = 系统注册表 + 本机扩展字体目录（静默增强）
    def _lookup_local(tf, b):
        return _lookup_system(tf, b) or _lookup_wps(tf, b)

    first, second = ((_lookup_local, _lookup_bundled)
                     if effective_policy() == 'system_first'
                     else (_lookup_bundled, _lookup_local))

    spec = first(typeface, bold) or second(typeface, bold) \
        or _lookup_embedded(typeface, bold)
    if spec is None:
        spec = first('微软雅黑', bold) or second('微软雅黑', bold)
    if spec is not None:
        if not spec.metrics.calibrated and not _uncalibrated_warned:
            logger.warning(
                "该字体排版常数尚未 COM 校准，暂沿用校准字体值"
                "（多行文本行高可能有细微偏差）——用 experiments/calibrate_sourcehan.py 重校后消除")
            _uncalibrated_warned = True
        return spec

    if typeface:
        logger.warning(f"字体不可用（未注册或文件缺失），回退默认字体: {typeface!r}")
    # 最终回退：系统微软雅黑路径（与原 _DEFAULT_FONT 行为一致，可能不存在）
    msyh = SYSTEM_FONT_FILES.get(('微软雅黑', bold)) or SYSTEM_FONT_FILES[('微软雅黑', False)]
    return FontSpec(msyh, METRICS_MS_YAHEI, 'system', '')


def is_resolvable(typeface: str) -> bool:
    """该字体名是否有直接命中的字体文件（不含默认链回退）；供字体检查/预检用"""
    return any(l(typeface, b) is not None
               for l in (_lookup_system, _lookup_wps, _lookup_bundled,
                         _lookup_embedded)
               for b in (True, False))


def available_typefaces() -> Set[str]:
    """当前环境可用的字体名集合（小写）；系统表命中文件 + 捆绑族"""
    names = set()
    for (tf, _b), path in SYSTEM_FONT_FILES.items():
        if os.path.exists(path):
            names.add(tf.lower())
    for fname, aliases in _SYSTEM_NAME_ALIASES.items():
        for path in set(SYSTEM_FONT_FILES.values()):
            if os.path.basename(path).lower() == fname and os.path.exists(path):
                names.update(aliases)
    d = bundled_font_dir()
    if d:
        for info in _BUNDLED_FAMILIES.values():
            if os.path.exists(os.path.join(d, info['regular'])):
                names.update(info['names'])
    for (tf, _b) in _embedded:
        names.add(tf)
    return names


def metrics_for(path: str) -> FontMetrics:
    """物理字体文件 -> 校准常数；捆绑思源按文件名分黑体/宋体，其余沿用雅黑值"""
    global _uncalibrated_warned
    d = bundled_font_dir()
    if d and os.path.abspath(path).lower().startswith(os.path.abspath(d).lower()):
        metrics = (METRICS_SOURCE_HAN_SERIF
                   if os.path.basename(path).lower().startswith('sourcehanserif')
                   else METRICS_SOURCE_HAN_SANS)
        if not metrics.calibrated and not _uncalibrated_warned:
            logger.warning(
                "该字体排版常数尚未 COM 校准，暂沿用校准字体值"
                "（多行文本行高可能有细微偏差）——用 experiments/calibrate_sourcehan.py 重校后消除")
            _uncalibrated_warned = True
        return metrics
    return METRICS_MS_YAHEI
