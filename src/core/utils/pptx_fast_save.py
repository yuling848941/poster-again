# -*- coding: utf-8 -*-
"""PPTX 快速保存：已经压缩过的媒体不再重复 deflate。

python-pptx 写包时对所有部件一律 ZIP_DEFLATED。海报模板 99% 的体积是 PNG / JPEG，
它们本身已是压缩格式，再 deflate 一遍几乎不减体积，却占掉整次保存的大部分时间
（50MB 的模板实测约 2 秒）。PowerPoint 自己保存时对这类媒体用的就是「只存储不压缩」。

做法与 python-pptx 的 PackageWriter 完全一致——同样的部件、同样的顺序、同样的 XML，
唯一的区别是已压缩媒体改用 ZIP_STORED。

安全网：每个进程第一次走快速路径时，会把结果与标准保存逐成员比对（连
[Content_Types].xml 在内），一字不差才继续用；比对不过，或当前 python-pptx 版本的
内部接口对不上，就退回标准的 prs.save()，行为与原来完全相同。
"""

import logging
import os
import tempfile
import zipfile

logger = logging.getLogger(__name__)

# 本身已是压缩格式的部件扩展名（不带点，小写）：解压后体积不变，没必要再压缩
ALREADY_COMPRESSED = frozenset((
    'png', 'jpg', 'jpeg', 'gif', 'webp', 'heic', 'avif',
    'zip', 'xlsx', 'docx', 'pptx', 'vsdx',             # 嵌入的 Office 文档本身就是 zip
    'emz', 'wmz',                                       # 压缩过的图元文件
    'mp3', 'mp4', 'm4a', 'aac', 'ogg',
))

# 已压缩媒体合计不到这个体积时直接用标准保存（这点体积压缩很快，不值得多跑一次校验）
FAST_SAVE_MIN_MEDIA_BYTES = 4 * 1024 * 1024

# None = 本进程还没校验过；True = 快速路径可用；False = 不可用（退回标准保存）
_verified = None


def save_pptx(prs, path) -> str:
    """把 Presentation 写成 .pptx。返回 'fast'（走了快速路径）或 'normal'。"""
    global _verified
    if _verified is False or _compressed_bytes(prs) < FAST_SAVE_MIN_MEDIA_BYTES:
        prs.save(path)
        return 'normal'
    try:
        _write_fast(prs, path)
        if _verified is None:
            if not _matches_standard_save(prs, path):
                raise RuntimeError('快速保存的结果与标准保存不一致')
            _verified = True
            logger.info('PPTX 快速保存已启用（已压缩媒体不再重复压缩）')
    except Exception as e:
        _verified = False
        logger.warning(f'PPTX 快速保存不可用，本进程内改用标准保存: {e}')
        prs.save(path)
        return 'normal'
    return 'fast'


def _is_compressed(ext: str) -> bool:
    # python-pptx 各版本对 PackURI.ext 的返回不一致（'png' 与 '.png' 都出现过），两种都认
    return ext.lower().lstrip('.') in ALREADY_COMPRESSED


def _compressed_bytes(prs) -> int:
    total = 0
    for part in prs.part.package.iter_parts():
        if _is_compressed(part.partname.ext):
            total += len(part.blob)
    return total


class _PerMemberZipWriter:
    """python-pptx 的 _PhysPkgWriter 替身：接口相同，只是按成员选择压缩方式。"""

    def __init__(self, zf: zipfile.ZipFile):
        self._zf = zf

    def write(self, pack_uri, blob) -> None:
        info = zipfile.ZipInfo(pack_uri.membername)
        info.compress_type = (zipfile.ZIP_STORED if _is_compressed(pack_uri.ext)
                              else zipfile.ZIP_DEFLATED)
        self._zf.writestr(info, blob)


def _write_fast(prs, path) -> None:
    """用 python-pptx 自己的 PackageWriter 生成每个成员（内容与标准保存同源），
    只把「写进 zip」换成按成员选压缩方式的实现。"""
    try:
        from pptx.opc.serialized import PackageWriter
    except ImportError:                      # python-pptx 0.6.x
        from pptx.opc.package import PackageWriter

    package = prs.part.package
    parts = tuple(package.iter_parts())
    pkg_rels = getattr(package, "rels", None)
    if pkg_rels is None:                     # 部分版本只有私有属性
        pkg_rels = package._rels
    writer = PackageWriter(path, pkg_rels, parts)
    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED,
                         strict_timestamps=False) as zf:
        phys = _PerMemberZipWriter(zf)
        writer._write_content_types_stream(phys)
        writer._write_pkg_rels(phys)
        writer._write_parts(phys)


def _matches_standard_save(prs, path) -> bool:
    """快速保存的结果是否与 prs.save() 逐成员、逐字节相同（只允许压缩方式不同）。"""
    fd, std_path = tempfile.mkstemp(suffix='.pptx')
    os.close(fd)
    try:
        prs.save(std_path)
        with zipfile.ZipFile(path) as fast, zipfile.ZipFile(std_path) as std:
            if fast.namelist() != std.namelist():
                return False
            return all(fast.read(name) == std.read(name) for name in std.namelist())
    finally:
        try:
            os.remove(std_path)
        except OSError:
            pass
