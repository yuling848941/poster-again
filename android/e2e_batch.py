# -*- coding: utf-8 -*-
"""模拟器批量流程自动化驱动（多屏 UI 版）：
首页选模板/数据 -> 字段绑定 -> 开始渲染 -> 完成网格 -> 大图预览。
全部按可见文本定位（Compose Text 会暴露为 uiautomator TextView 节点）。"""
import os
import re
import subprocess
import sys
import time
from pathlib import Path

ADB = r"C:\Users\Administrator\AppData\Local\Android\Sdk\platform-tools\adb.exe"
PKG = "com.posteragain.app"
HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "e2e_shots"
_ENV = {**os.environ, "MSYS_NO_PATHCONV": "1"}


def sh(*args, binary=False):
    r = subprocess.run([ADB, *args], capture_output=True, env=_ENV, timeout=120)
    return r.stdout if binary else r.stdout.decode('utf-8', errors='ignore')


def dump():
    sh("shell", "uiautomator", "dump", "/sdcard/ui.xml")
    return sh("shell", "cat", "/sdcard/ui.xml")


def bounds_center(node_xml):
    b = re.search(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', node_xml)
    if not b:
        return None
    l, t, r, bt = map(int, b.groups())
    return (l + r) // 2, (t + bt) // 2


def bounds_of(xml, text):
    m = re.search(r'<node[^>]*text="%s"[^>]*>' % re.escape(text), xml)
    return bounds_center(m.group(0)) if m else None


def tap(xy):
    sh("shell", "input", "tap", str(xy[0]), str(xy[1]))


def wait_pred(pred, timeout=25, interval=1.5):
    t0 = time.time()
    while time.time() - t0 < timeout:
        xy = pred(dump())
        if xy:
            return xy
        time.sleep(interval)
    return None


def wait_text(text, timeout=25):
    return wait_pred(lambda xml: bounds_of(xml, text), timeout)


def find_contains(xml, substr):
    """子串匹配节点文本（如 "张海报已就绪" 匹配 "5 张海报已就绪"）"""
    for m in re.finditer(r'<node[^>]*text="([^"]+)"[^>]*>', xml):
        if substr in m.group(1):
            return bounds_center(m.group(0))
    return None


def wait_contains(substr, timeout=25):
    return wait_pred(lambda xml: find_contains(xml, substr), timeout)


def wait_prefix(prefix, timeout=10):
    """按前缀匹配节点文本（如 "开始生成（5 张）"）"""
    def find(xml):
        for m in re.finditer(r'<node[^>]*text="([^"]+)"[^>]*>', xml):
            t = m.group(1)
            if len(t) >= 4 and t.startswith(prefix):
                return bounds_center(m.group(0))
        return None
    return wait_pred(find, timeout)


def pick_file(filename):
    """在文件选择器当前列表中滚屏查找并点击目标文件行。
    点击用「行主体偏移」（文本左侧 +280px），避开 preview 图标；
    点击后通过 E2E logcat 确认回调真的触发，未触发则重试。"""
    def find_row():
        xml = dump()
        for m in re.finditer(r'<node[^>]*text="([^"]+)"[^>]*bounds="(\[[\d,\]\[]+\])"', xml):
            t = m.group(1)
            if len(t) >= 4 and (t == filename or filename.startswith(t)):
                b = re.search(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', m.group(0))
                l, tp, r, bt = map(int, b.groups())
                return (min(l + 280, r - 20), (tp + bt) // 2)
        return None

    # 先滚一屏（列表懒加载，目标常在首屏之外），再交替 查找/滚屏
    sh("shell", "input", "swipe", "540", "1800", "540", "600", "400")
    time.sleep(1.5)
    for _ in range(8):
        xy = find_row()
        if xy:
            sh("logcat", "-c")           # 只看点击之后的新日志
            tap(xy)
            time.sleep(3)
            if "pickUri entry" in sh("logcat", "-d", "-s", "E2E:V"):
                return
            print("    [重试] 点击未触发回调")
            # 选择器可能已关闭（取消了），重新打开由调用方负责；这里先等界面稳定
            time.sleep(2)
            if "Files in Downloads" not in dump() and "Recent" not in dump():
                return                   # 选择器已关，交由上层重开
        else:
            sh("shell", "input", "swipe", "540", "1800", "540", "600", "400")
            time.sleep(1.5)
    raise RuntimeError(f"选择器中找不到/选不中 {filename}")


def shot(name):
    sh("shell", "screencap", "-p", f"/sdcard/{name}.png")
    OUT_DIR.mkdir(exist_ok=True)
    sh("pull", f"/sdcard/{name}.png", str(OUT_DIR / f"{name}.png"))
    print(f"    [截图] {OUT_DIR / (name + '.png')}")


def logcat_grep(tag):
    out = sh("logcat", "-d")
    return [l for l in out.splitlines() if tag in l]


def tap_text(text, timeout=15):
    xy = wait_text(text, timeout)
    if not xy:
        raise RuntimeError(f"找不到文本: {text}")
    tap(xy)
    return xy


def main():
    data_file = sys.argv[1] if len(sys.argv) > 1 else "batch_data.csv"
    tpl_file = sys.argv[2] if len(sys.argv) > 2 else "batch_tpl.pptx"
    repo = HERE.parent
    # 安装最新构建的 debug APK（避免对旧包跑 E2E 得出错误结论）
    apk = HERE / "app" / "build" / "outputs" / "apk" / "debug" / "app-debug.apk"
    if apk.exists():
        print("[0] 安装最新 debug APK")
        sh("install", "-r", str(apk))
    else:
        print(f"[0] [警告] 未找到 {apk}，使用设备上已安装的版本")
    sh("push", str(repo / data_file), "/sdcard/Download/")
    sh("push", str(repo / tpl_file), "/sdcard/Download/")
    sh("shell", "pm", "clear", PKG)      # 清旧会话/产物/记录
    time.sleep(1)
    sh("shell", "am", "start", "-n", f"{PKG}/.MainActivity")
    time.sleep(8)
    sh("logcat", "-c")

    print(f"[1] 选择模板 {tpl_file}")
    for i in range(3):
        xy = wait_text("点击选择 PPT 模板文件（.pptx）", timeout=5)
        if xy:
            tap(xy)
            time.sleep(3)
        try:
            pick_file(tpl_file)
        except RuntimeError as e:
            print("   ", e)
        if wait_text("模板就绪，还需选择数据文件", timeout=35):
            break
        print(f"    [重开] 模板未就绪，重试第{i+1}轮")
    else:
        raise RuntimeError("模板会话未建立")
    for l in logcat_grep("[批量]")[:3]:
        print("   ", l.split("python.stdout:")[-1].strip())

    print(f"[2] 选择数据 {data_file}")
    for i in range(3):
        xy = wait_text("点击选择数据文件（Excel / CSV / JSON）", timeout=5)
        if xy:
            tap(xy)
            time.sleep(3)
        try:
            pick_file(data_file)
        except RuntimeError as e:
            print("   ", e)
        if wait_prefix("开始生成（", timeout=35):
            break
        print(f"    [重开] 数据未就绪，重试第{i+1}轮")
    else:
        raise RuntimeError("数据会话未建立（CTA 未显示张数）")
    n_target = 0
    for l in logcat_grep("[批量]"):
        m = re.search(r"数据 (\d+) 行", l)
        if m:
            n_target = int(m.group(1))
    print(f"    目标行数: {n_target}")
    shot("01_ready")

    print("[3] 进入字段绑定")
    tap_text_bounds = wait_prefix("开始生成（", timeout=5)
    tap(tap_text_bounds)
    if not wait_text("字段绑定", timeout=15):
        raise RuntimeError("未进入字段绑定页")
    if not wait_text("未绑定的字段将保留模板占位原样", timeout=5):
        raise RuntimeError("字段绑定页缺少底部提示")
    shot("02_binding")

    print("[4] 开始渲染")
    tap_text("开始渲染")
    # 小模板渲染极快，可能瞬间完成：『正在生成』与完成态都算进入
    t0 = time.time()
    rendering_seen = False
    while time.time() - t0 < 20:
        xml = dump()
        if bounds_of(xml, "正在生成"):
            rendering_seen = True
            break
        if find_contains(xml, "张海报已就绪"):
            break
        time.sleep(1)
    if rendering_seen:
        shot("03_rendering")
    elif not find_contains(dump(), "张海报已就绪"):
        raise RuntimeError("未进入渲染状态")
    ok = wait_contains("张海报已就绪", timeout=300)
    shot("04_done")
    if not ok:
        raise RuntimeError("渲染超时未完成")

    print("[5] 大图预览（第一张）")
    # 产物名取决于表头是否含姓名类列（英文表头回退为 output），点第一个网格单元即可
    xy = wait_pred(lambda xml: find_contains(xml, "output"), timeout=10)
    if not xy:
        raise RuntimeError("完成网格中未找到可点的海报单元")
    tap(xy)
    if not wait_text("保存到相册", timeout=10):
        raise RuntimeError("未进入大图预览")
    shot("05_preview")

    # 产物核对：应用私有目录 files/batch/row_*
    rows = sh("shell", "run-as", PKG, "ls", "files/batch").split()
    n_rows = len([d for d in rows if d.startswith("row_")])
    print(f"    row 目录数: {n_rows}")
    print("DONE" if n_rows >= n_target else "INCOMPLETE")


if __name__ == "__main__":
    main()
