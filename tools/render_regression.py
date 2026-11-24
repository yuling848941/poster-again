#!/usr/bin/env python3
"""渲染像素回归工具：采集 / 对比每一次渲染输出的像素指纹。

用途：改动渲染引擎（render_core / fontfx）或批量管线前后各采一次，逐像素确认输出没变。

三种用法：

    # 1) 跑任意命令（通常是全量测试），记录期间所有渲染结果的像素指纹
    python tools/render_regression.py run --out before.jsonl -- python tests/run_all.py

    # 2) 用真实模板走三条路径（库入口 render / 批量快速通道 Personalizer / 桌面慢路径）
    python tools/render_regression.py corpus --out before.jsonl "samples/templates/xxx.pptx" ...

    # 3) 对比两份指纹
    python tools/render_regression.py verify before.jsonl after.jsonl

`run` 的原理：往子进程注入一个 sitecustomize 钩子，包住
PillowImageExporter.export_presentation / _render_slide，对每张产出的图算
size + 像素 sha256。指纹按「包内容指纹 + 页 + 参数」归键，与输出文件名、临时目录无关。

退出码：verify 有差异时 1。
"""

import argparse
import hashlib
import io
import json
import os
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]

_HOOK_MODULE = "_render_regression_hook"
_HOOK_SOURCE = r'''
"""渲染回归钩子：由 tools/render_regression.py 通过 sitecustomize 注入子进程。"""
import hashlib
import io
import json
import os
import sys
import threading

_OUT = os.environ.get("PA_REGRESS_OUT")
_LOCK = threading.Lock()
_SRC = os.path.basename(sys.argv[0]) if sys.argv and sys.argv[0] else "?"


def _write(rec):
    line = json.dumps(rec, ensure_ascii=False)
    with _LOCK:
        with open(_OUT, "a", encoding="utf-8") as f:
            f.write(line + "\n")


def _deck_fp(prs):
    """模板内容指纹。注意不能用 prs.save() 的字节：zip 里带写入时间戳，跨运行不稳定。
    这里对所有非媒体部件（主题/母版/版式/幻灯片 XML 等）的内容做哈希。"""
    try:
        h = hashlib.sha256()
        for part in prs.part.package.iter_parts():
            name = str(part.partname)
            if name.startswith("/ppt/media/"):
                continue
            h.update(name.encode("utf-8"))
            h.update(bytes(1))
            h.update(part.blob)
            h.update(bytes(1))
        return h.hexdigest()[:20]
    except Exception as e:
        return "na:" + type(e).__name__


def _slide_no(prs, slide):
    try:
        return list(prs.slides).index(slide) + 1
    except Exception:
        return -1


def _slide_no_of_el(prs, el):
    """形状/容器元素 -> 所在页号；母版、版式上的元素不在任何一页里，返回 -1。"""
    node = el
    while node is not None and not node.tag.endswith("}cSld"):
        node = node.getparent()
    sld = node.getparent() if node is not None else None
    for i, s in enumerate(prs.slides, 1):
        if s._element is sld:
            return i
    return -1


def _img_fp(img):
    return hashlib.sha256(img.tobytes()).hexdigest()


def _patch(mod):
    Exp = getattr(mod, "PillowImageExporter", None)
    if Exp is None or getattr(Exp, "_pa_regress_patched", False):
        return
    Exp._pa_regress_patched = True

    def _emit(api, img, shape):
        rec = {"src": _SRC, "api": api, "size": list(img.size), "mode": img.mode,
               "sha": _img_fp(img)}
        if shape is not None:
            rec["shape"] = getattr(shape, "name", None)
            rec["sid"] = getattr(shape, "shape_id", None)
            try:
                prs = shape.part.package.presentation_part.presentation
                rec["deck"] = _deck_fp(prs)
                rec["slide"] = _slide_no_of_el(prs, shape._element)
            except Exception:
                pass
        _write(rec)

    def _wrap(api, orig):
        """所有绘制原语的第一个参数都是被就地修改的画布：调用后记录画布指纹。"""
        def wrapped(self, canvas, *args, **kwargs):
            r = orig(self, canvas, *args, **kwargs)
            if _OUT:
                try:
                    shape = next((a for a in args
                                  if hasattr(a, "_element") and hasattr(a, "part")), None)
                    _emit(api, canvas, shape)
                except Exception as e:
                    _write({"src": _SRC, "api": api, "err": repr(e)})
            return r
        return wrapped

    for _name in ("_draw_slide_background", "_draw_picture", "_draw_autoshape",
                  "_draw_text_frame", "_draw_group", "_draw_shape"):
        setattr(Exp, _name, _wrap(_name, getattr(Exp, _name)))

    orig_render = Exp._render_slide
    orig_export = Exp.export_presentation

    def _render_slide(self, prs, slide, theme, quality):
        img = orig_render(self, prs, slide, theme, quality)
        if _OUT:
            try:
                _write({"src": _SRC, "api": "slide", "deck": _deck_fp(prs),
                        "slide": _slide_no(prs, slide), "q": round(float(quality), 3),
                        "size": list(img.size), "mode": img.mode, "sha": _img_fp(img)})
            except Exception as e:
                _write({"src": _SRC, "api": "slide", "err": repr(e)})
        return img

    def export_presentation(self, prs, output_dir, image_format="PNG", quality=1.0, theme=None):
        files = orig_export(self, prs, output_dir, image_format, quality, theme)
        if _OUT:
            from PIL import Image
            try:
                fp = _deck_fp(prs)
                for i, p in enumerate(files, 1):
                    with Image.open(p) as im:
                        _write({"src": _SRC, "api": "export", "deck": fp, "slide": i,
                                "fmt": str(image_format).upper(),
                                "q": round(float(quality), 3),
                                "size": list(im.size), "mode": im.mode, "sha": _img_fp(im)})
            except Exception as e:
                _write({"src": _SRC, "api": "export", "err": repr(e)})
        return files

    Exp._render_slide = _render_slide
    Exp.export_presentation = export_presentation


_TARGETS = {"pillow_image_exporter"}

try:
    import importlib.abc
    import importlib.machinery

    class _Finder(importlib.abc.MetaPathFinder):
        """只在 pillow_image_exporter 被加载后打补丁；其余模块一律不插手。"""

        def find_spec(self, fullname, path=None, target=None):
            if fullname.rsplit(".", 1)[-1] not in _TARGETS:
                return None
            spec = importlib.machinery.PathFinder.find_spec(fullname, path)
            if spec is None or spec.loader is None:
                return None
            orig_exec = spec.loader.exec_module

            def exec_module(module, _orig=orig_exec):
                _orig(module)
                try:
                    _patch(module)
                except Exception as e:
                    if _OUT:
                        _write({"src": _SRC, "api": "hook", "err": repr(e)})

            spec.loader.exec_module = exec_module
            return spec

    sys.meta_path.insert(0, _Finder())
except Exception as _e:  # 钩子装不上也不能影响被测程序
    if _OUT:
        try:
            with open(_OUT, "a", encoding="utf-8") as _f:
                _f.write(json.dumps({"src": _SRC, "api": "hook-init", "err": repr(_e)}) + "\n")
        except Exception:
            pass
'''


def _read(path):
    recs = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                recs.append(json.loads(line))
    return recs


def _key(rec):
    """一条指纹的归键：哪次绘制（文件/用例、模板、页、形状、参数）。不含像素本身。"""
    return (rec.get("src"), rec.get("api"), rec.get("file"), rec.get("deck"),
            rec.get("slide"), rec.get("shape"), rec.get("sid"),
            rec.get("fmt"), rec.get("q"), rec.get("row"))


def cmd_run(args):
    out = Path(args.out).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    if args.fresh and out.exists():
        out.unlink()
    hook_dir = Path(tempfile.mkdtemp(prefix="pa_regress_hook_"))
    (hook_dir / (_HOOK_MODULE + ".py")).write_text(_HOOK_SOURCE, encoding="utf-8")
    (hook_dir / "sitecustomize.py").write_text(
        "import os, sys\n"
        "if os.environ.get('PA_REGRESS_HOOK_DIR'):\n"
        "    import %s  # noqa: F401\n" % _HOOK_MODULE, encoding="utf-8")
    env = dict(os.environ)
    sep = ";" if os.name == "nt" else ":"
    env["PYTHONPATH"] = str(hook_dir) + (sep + env["PYTHONPATH"] if env.get("PYTHONPATH") else "")
    env["PA_REGRESS_HOOK_DIR"] = str(hook_dir)
    env["PA_REGRESS_OUT"] = str(out)
    env["PYTHONIOENCODING"] = env.get("PYTHONIOENCODING", "utf-8")
    print(f"[regression] 记录到 {out}；执行: {' '.join(args.argv)}", flush=True)
    rc = subprocess.call(args.argv, cwd=str(REPO_ROOT), env=env)
    n = len(_read(out)) if out.exists() else 0
    print(f"[regression] 命令退出码 {rc}，共记录 {n} 条渲染指纹", flush=True)
    return rc


def _file_fp(path):
    from PIL import Image
    with Image.open(path) as im:
        return {"size": list(im.size), "mode": im.mode,
                "sha": hashlib.sha256(im.tobytes()).hexdigest()}


def _clean_dir(d):
    import shutil
    p = Path(d)
    if p.exists():
        shutil.rmtree(p)
    p.mkdir(parents=True)
    return p


def _record(out, src, path, files, **tags):
    with open(out, "a", encoding="utf-8") as f:
        for i, p in enumerate(files, 1):
            rec = {"src": src, "api": "file", "file": path, "slide": i}
            rec.update(tags)
            rec.update(_file_fp(p))
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")


def cmd_corpus(args):
    """真实模板走三条路径；同时直接报告「快速通道 vs 桌面慢路径」是否逐像素一致。"""
    os.environ["POSTERAGAIN_IMAGE_ENGINE"] = "python"     # 回归只测内置引擎，绝不驱动 Office
    sys.path.insert(0, str(REPO_ROOT))
    import logging
    logging.disable(logging.CRITICAL)
    from PIL import Image, ImageChops
    from render_core import render, Personalizer
    from src.core.processors.pptx_processor import PPTXProcessor

    out = Path(args.out).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    if args.fresh and out.exists():
        out.unlink()
    tmp = Path(tempfile.mkdtemp(prefix="pa_regress_corpus_"))
    try:
        for tpl in args.files:
            name = os.path.basename(tpl)
            if not os.path.exists(tpl):
                print(f"  [跳过] 不存在: {tpl}")
                continue
            print(f"  [渲染] {name}", flush=True)
            # 1) 库入口
            d = _clean_dir(tmp / "lib")
            res = render(tpl, str(d), "PNG", 1.0)
            _record(out, "corpus:render", name, res.images, fmt="PNG", q=1.0)
            # 2) 批量快速通道：两行不同数据，顺带验证 restore()
            pz = Personalizer(tpl)
            fields = pz.fields()
            rows = [{f: f"张三{i}" for i, f in enumerate(fields)},
                    {f: ("李四五六七八九十" if i % 2 else "12,345元") for i, f in enumerate(fields)}]
            fast = []
            for r, row in enumerate(rows):
                d = _clean_dir(tmp / f"fast{r}")
                files = pz.render_row(row, str(d), "PNG", 1.0)
                fast.append(files)
                _record(out, "corpus:fast", name, files, fmt="PNG", q=1.0, row=r)
            # 3) 桌面慢路径（与 batch_generate 的 Office 分支同一实现）
            for r, row in enumerate(rows):
                proc = PPTXProcessor()
                proc.load_template(tpl)
                proc.replace_placeholders(row)
                d = _clean_dir(tmp / f"slow{r}")
                files = proc.export_to_images(output_dir=str(d), format="PNG", quality=1.0,
                                              base_filename=f"{r + 1}_张三")
                _record(out, "corpus:slow", name, files, fmt="PNG", q=1.0, row=r)
                if len(files) == len(fast[r]) and files:
                    for a, b in zip(fast[r], files):
                        with Image.open(a) as ia, Image.open(b) as ib:
                            same = ia.size == ib.size and ImageChops.difference(
                                ia.convert("RGB"), ib.convert("RGB")).getbbox() is None
                        print(f"    快速通道 vs 慢路径 第{r + 1}行 {os.path.basename(a)}: "
                              f"{'逐像素一致' if same else '!! 有差异 !!'}")
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    n = len(_read(out))
    print(f"[regression] corpus 完成，当前指纹文件共 {n} 条", flush=True)
    return 0


def _fp_key(rec):
    """一条指纹的完整键：哪次绘制 + 画布尺寸/模式 + 像素哈希"""
    return (_key(rec), tuple(rec.get("size") or ()), rec.get("mode"), rec.get("sha"),
            rec.get("err"))


def cmd_verify(args):
    """逐次绘制对比：同一次绘制（同模板/页/形状/参数）的画布必须逐像素一致。"""
    def fp(rec):
        full = _fp_key(rec)
        if getattr(args, "ignore_deck", False):
            key = list(full[0])
            key[3] = None                      # 去掉 deck 字段
            full = (tuple(key),) + full[1:]
        return full

    before = Counter(fp(r) for r in _read(args.before))
    after = Counter(fp(r) for r in _read(args.after))

    def by_call(counter):
        out = {}
        for full, n in counter.items():
            out.setdefault(full[0], {})[full[1:]] = n
        return out

    calls_before, calls_after = by_call(before), by_call(after)
    changed = [k for k, v in calls_before.items() if k in calls_after and calls_after[k] != v]
    missing = [k for k in calls_before if k not in calls_after]
    added = [k for k in calls_after if k not in calls_before]
    err = [r for r in _read(args.after) if r.get("err")]

    total = sum(before.values())
    same = total - sum(sum(calls_before[k].values()) for k in changed + missing)
    print(f"基线 {sum(before.values())} 条 / 现状 {sum(after.values())} 条指纹")
    print(f"  逐像素一致:      {same}")
    print(f"  像素有变化:      {len(changed)} 次绘制")
    print(f"  基线有、现状没有: {len(missing)} 次绘制")
    print(f"  现状新增:        {len(added)} 次绘制")
    if err:
        print(f"  钩子记录的异常:  {len(err)} 条，例: {err[0]}")
    for title, lst in (("像素有变化", changed), ("基线有、现状没有", missing), ("现状新增", added)):
        for k in lst[:args.show]:
            print(f"    [{title}] {k}")
    ok = not (changed or missing or err)
    print("结论: " + ("与基线逐像素一致 ✅" if ok else "与基线有差异 ❌"))
    return 0 if ok else 1


def main():
    ap = argparse.ArgumentParser(description="渲染像素回归")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("run", help="在钩子下执行一条命令并记录所有渲染指纹")
    p.add_argument("--out", required=True)
    p.add_argument("--fresh", action="store_true", help="先清空已有指纹文件")
    p.add_argument("argv", nargs=argparse.REMAINDER)
    p.set_defaults(fn=cmd_run)
    p = sub.add_parser("corpus", help="真实模板走 render / 快速通道 / 慢路径")
    p.add_argument("--out", required=True)
    p.add_argument("--fresh", action="store_true")
    p.add_argument("files", nargs="*")
    p.set_defaults(fn=cmd_corpus)
    p = sub.add_parser("verify", help="对比两份指纹")
    p.add_argument("before")
    p.add_argument("after")
    p.add_argument("--show", type=int, default=20)
    p.add_argument("--ignore-deck", action="store_true",
                   help="比对时不看模板标识（旧基线是用带时间戳的 zip 字节算的标识，跨运行不稳定）")
    p.set_defaults(fn=cmd_verify)
    args = ap.parse_args()
    if args.fn is cmd_run:
        argv = [c for c in args.argv if c != "--"]
        if not argv:
            ap.error("run 需要一条命令，例如: run --out x.jsonl -- python tests/run_all.py")
        # 裸 `python` 一律解释为当前解释器（避免落到系统 Python / 应用商店别名）
        if os.path.basename(argv[0]).lower() in ("python", "python3", "py"):
            argv[0] = sys.executable
        args.argv = argv
    sys.exit(args.fn(args))


if __name__ == "__main__":
    main()
