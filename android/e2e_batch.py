# -*- coding: utf-8 -*-
"""模拟器批量流程自动化驱动：①选模板 ②选数据 ③批量渲染，全 uiautomator 定位。"""
import os
import re
import subprocess
import sys
import time

ADB = r"C:\Users\Administrator\AppData\Local\Android\Sdk\platform-tools\adb.exe"
PKG = "com.posteragain.app"
_ENV = {**os.environ, "MSYS_NO_PATHCONV": "1"}


def sh(*args, binary=False):
    r = subprocess.run([ADB, *args], capture_output=True, env=_ENV, timeout=60)
    return r.stdout if binary else r.stdout.decode('utf-8', errors='ignore')


def dump():
    sh("shell", "uiautomator", "dump", "/sdcard/ui.xml")
    return sh("shell", "cat", "/sdcard/ui.xml")


def bounds_of(xml, text):
    m = re.search(r'<node[^>]*text="%s"[^>]*>' % re.escape(text), xml)
    if not m:
        return None
    b = re.search(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', m.group(0))
    if not b:
        return None
    l, t, r, bt = map(int, b.groups())
    return (l + r) // 2, (t + bt) // 2


def tap(xy):
    sh("shell", "input", "tap", str(xy[0]), str(xy[1]))


def wait_text(text, timeout=25):
    t0 = time.time()
    while time.time() - t0 < timeout:
        xy = bounds_of(dump(), text)
        if xy:
            return xy
        time.sleep(1.5)
    return None


def wait_prefix(filename, timeout=10):
    """长文件名在列表中被截断显示（如 "10月递交贺"），按前缀匹配节点"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        for m in re.finditer(r'<node[^>]*text="([^"]+)"[^>]*>', dump()):
            t = m.group(1)
            if len(t) >= 4 and filename.startswith(t):
                b = re.search(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"',
                              m.group(0))
                if b:
                    l, tp, r, bt = map(int, b.groups())
                    return (l + r) // 2, (tp + bt) // 2
        time.sleep(1.5)
    return None


def pick_file(filename):
    """系统文件选择器：开抽屉 -> Downloads -> （必要时滚屏）点击目标文件"""
    entered = False
    for attempt in range(3):
        tap((72, 152))                   # hamburger
        time.sleep(2)
        xy = wait_text("Downloads", timeout=8)
        if not xy:
            continue
        tap(xy)
        time.sleep(2.5)
        if "Recent files" not in dump():  # 根页特征消失 = 已进入 Downloads 列表
            entered = True
            break
        print(f"    [重试] 第{attempt+1}次未进入 Downloads 列表")
    if not entered:
        raise RuntimeError("无法进入 Downloads 列表")
    for _ in range(5):                    # 列表可能超屏，滚屏查找
        xy = wait_text(filename, timeout=3) or wait_prefix(filename, timeout=2)
        if xy:
            tap(xy)
            return
        sh("shell", "input", "swipe", "360", "950", "360", "350", "400")
        time.sleep(1.5)
    raise RuntimeError(f"Downloads 中找不到 {filename}")


def clickable_list(xml):
    """按文档序返回本应用所有 clickable 节点中心坐标（Compose 文本不可见，按序定位）"""
    out = []
    for m in re.finditer(r'<node[^>]*clickable="true"[^>]*>', xml):
        if PKG not in m.group(0):
            continue
        b = re.search(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', m.group(0))
        if b:
            l, t, r, bt = map(int, b.groups())
            out.append(((l + r) // 2, (t + bt) // 2))
    return out


def tap_action(index, timeout=15):
    """点第 index 个可点击控件（0=①模板 1=②数据 2=③批量 3=预览）"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        nodes = clickable_list(dump())
        if len(nodes) > index:
            tap(nodes[index])
            return
        time.sleep(1.5)
    raise RuntimeError(f"找不到第 {index} 个可点击控件")


def logcat_grep(tag):
    out = sh("logcat", "-d")
    return [l for l in out.splitlines() if tag in l]


def main():
    data_file = sys.argv[1] if len(sys.argv) > 1 else "batch_data.csv"
    tpl_file = sys.argv[2] if len(sys.argv) > 2 else "batch_tpl.pptx"
    sh("shell", "pm", "clear", PKG)      # 清旧批量产物，避免干扰行数等待
    time.sleep(1)
    sh("shell", "am", "start", "-n", f"{PKG}/.MainActivity")
    time.sleep(7)
    sh("logcat", "-c")

    print(f"[1] 选择模板 {tpl_file}")
    tap_action(0)
    time.sleep(4)
    pick_file(tpl_file)
    time.sleep(6)
    for l in logcat_grep("[批量]"):
        print("   ", l.split("python.stdout:")[-1].strip())

    print(f"[2] 选择数据 {data_file}")
    tap_action(1)
    time.sleep(4)
    pick_file(data_file)
    time.sleep(6)
    for l in logcat_grep("[批量]")[-2:]:
        print("   ", l.split("python.stdout:")[-1].strip())

    # 行数从会话状态拿（"数据 N 行 × M 列"）——必须在 logcat -c 之前解析
    n_target = 5
    for l in logcat_grep("[批量]"):
        m = re.search(r"数据 (\d+) 行", l)
        if m:
            n_target = int(m.group(1))
    print(f"    目标行数: {n_target}")

    print("[3] 批量渲染")
    sh("logcat", "-c")
    t0 = time.time()
    tap_action(2)
    import os

    def row_png(i):
        # 产物已按 PC 规则重命名（如 1_陈春兰.png），列 row 目录找 .png
        out = sh("shell", "run-as", PKG, "ls", f"files/batch/row_{i}")
        files = [f for f in out.split() if f.endswith(".png")]
        return files[0] if files else ""

    for _ in range(60):
        time.sleep(2)
        out = sh("shell", "run-as", PKG, "ls", "files/batch")
        n = len([d for d in out.split() if d.startswith("row_")])
        if n >= n_target and all(row_png(i) for i in range(n_target)):
            break
    dt = time.time() - t0
    print(f"    耗时 {dt:.0f}s（含 UI 等待轮询）")
    ok = 0
    for i in range(n_target):
        name = row_png(i)
        if name:
            ok += 1
            print(f"    row_{i}: {name}")
    print("DONE" if ok >= n_target else "INCOMPLETE")


if __name__ == "__main__":
    main()
