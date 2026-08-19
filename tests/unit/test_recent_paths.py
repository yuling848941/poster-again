#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最近使用（MRU）路径列表测试：去重置顶、限长、失效清理、配置持久化往返
"""

import sys
import os
import tempfile
import shutil

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from src.config import ConfigManager
from src.gui.path_manager import PathManager

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
    tmpdir = tempfile.mkdtemp(prefix="mru_test_")
    # ConfigManager 开发模式优先使用 CWD 下的 PosterAgain_config.yaml，
    # 必须 chdir 到临时目录隔离，避免污染仓库根的运行时快照
    original_cwd = os.getcwd()
    os.chdir(tmpdir)
    try:
        # 准备三个真实存在的文件（validate_path 校验存在性）
        paths = [os.path.join(tmpdir, f"tpl{i}.pptx") for i in range(10)]
        for p in paths:
            open(p, 'w').close()

        print("(1) ConfigMixin：去重置顶 + 限长 8")
        cfg_file = os.path.join(tmpdir, "config_test.yaml")
        cm = ConfigManager(config_file=cfg_file)
        for p in paths:
            cm.update_recent_path('template', p)
        recents = cm.load_recent_paths('template')
        check("限长 8", len(recents) == 8)
        check("最新在前", recents[0] == paths[-1])
        check("去重（重复添加不增长）", (cm.update_recent_path('template', paths[-1]),
                                       len(cm.load_recent_paths('template')) == 8)[1])
        check("列表新→旧有序", recents == [paths[-1], paths[-2], paths[-3], paths[-4],
                                           paths[-5], paths[-6], paths[-7], paths[-8]])

        print("(2) 不存在的路径被忽略")
        before = cm.load_recent_paths('template')
        cm.update_recent_path('template', os.path.join(tmpdir, "ghost.pptx"))
        check("幽灵路径未入列", cm.load_recent_paths('template') == before)

        print("(3) 配置持久化往返（新实例读取同一文件）")
        cm2 = ConfigManager(config_file=cfg_file)
        check("重新加载后列表一致", cm2.load_recent_paths('template') == before)

        print("(4) 三类型互不干扰")
        cm.update_recent_path('data', paths[0])
        cm.update_recent_path('output', tmpdir)
        check("data 列表独立", cm.load_recent_paths('data') == [paths[0]])
        check("output 列表独立", cm.load_recent_paths('output') == [tmpdir])
        check("template 列表不受影响", cm.load_recent_paths('template') == before)

        print("(5) PathManager：add/get/prune + save_path_on_select 同步 MRU")
        pm = PathManager(cm)
        pm.add_recent('data', paths[1])
        check("add_recent 生效", pm.get_recents('data') == [paths[1], paths[0]])
        pm.save_path_on_select('data', paths[0])
        check("save_path_on_select 同步 MRU 置顶", pm.get_recents('data') == [paths[0], paths[1]])
        pm.prune_recents('data', [paths[1]])
        check("prune 覆写", pm.get_recents('data') == [paths[1]])
        check("未知类型返回空列表", pm.get_recents('unknown') == [])

        print("(6) 默认配置含三个 recent 键")
        cm3 = ConfigManager(config_file=os.path.join(tmpdir, "config_fresh.yaml"))
        for key in ('paths.recent_templates', 'paths.recent_data_files', 'paths.recent_output_dirs'):
            check(f"{key} 默认空列表", cm3.get(key, None) == [])

    finally:
        os.chdir(original_cwd)
        shutil.rmtree(tmpdir, ignore_errors=True)

    print(f"\n结果: {PASS} 通过, {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
