#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI 冒烟测试（offscreen）：MainWindow 可构造、容错/MRU 新钩子在位
"""

import sys
import os
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

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


def main():
    # 隔离：避免读写仓库根的运行时快照
    tmpdir = tempfile.mkdtemp(prefix="gui_smoke_")
    original_cwd = os.getcwd()
    os.chdir(tmpdir)

    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication(sys.argv)

    try:
        from src.gui.main_window import MainWindow
        win = MainWindow()

        check("MainWindow 构造成功", win is not None)
        check("存在 _collect_matching_rules_from_table",
              callable(getattr(win, "_collect_matching_rules_from_table", None)))
        check("存在 retry_failed_rows", callable(getattr(win, "retry_failed_rows", None)))
        check("存在 show_recent_menu", callable(getattr(win, "show_recent_menu", None)))
        check("存在 _apply_recent_path", callable(getattr(win, "_apply_recent_path", None)))

        # 模拟 worker 发射新版 generation_completed(object) 摘要（纯数据路径，
        # 弹窗会阻塞——这里 monkeypatch QMessageBox 避免真弹）
        from PySide6.QtWidgets import QMessageBox
        shown = []
        QMessageBox.information = staticmethod(lambda *a, **k: shown.append(a) or QMessageBox.StandardButton.Ok)
        QMessageBox.warning = staticmethod(lambda *a, **k: shown.append(a) or QMessageBox.StandardButton.Ok)
        # 失败分支走 QMessageBox 实例 box.exec()，offscreen 下会永久阻塞
        QMessageBox.exec = lambda self, *a, **k: shown.append(("exec",)) or QMessageBox.StandardButton.Ok

        summary = {"success": 3, "failed": [], "total": 3, "is_retry": False, "precheck": []}
        win.on_generation_completed(summary)
        check("完成摘要弹窗触发", len(shown) == 1)

        win._last_failed_rows = [{"index": 1, "row_number": 3, "name": "李四",
                                  "stage": "ppt", "error": "x"}]
        win._last_batch_output_path = os.path.join(tmpdir, "out")
        summary2 = {"success": 3, "failed": win._last_failed_rows, "total": 4,
                    "is_retry": False, "precheck": []}
        win.on_generation_completed(summary2)
        check("部分失败弹窗触发", len(shown) == 2)

        # 验证 generation_completed 信号可携带 object 摘要（跨线程由 Qt 自动排队）
        received = []
        win.worker_thread.generation_completed.connect(lambda s: received.append(s))
        win.worker_thread.generation_completed.emit(summary)
        check("generation_completed(object) 信号 emit/receive 正常",
              received and isinstance(received[0], dict) and received[0]["success"] == 3)
        check("worker 存在 set_retry_rows",
              callable(getattr(win.worker_thread, "set_retry_rows", None)))

        win.close()
    except Exception as e:
        import traceback
        traceback.print_exc()
        check(f"MainWindow 冒烟（异常: {e}）", False)
    finally:
        os.chdir(original_cwd)
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)

    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
