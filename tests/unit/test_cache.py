#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试配置管理器中的 Office 缓存机制（保存/加载/清除/状态）

注意：原测试针对已删除的 OfficeSuiteManager，现改为验证当前
ConfigManager 提供的 office_cache 相关方法。
"""

import os
import sys
import tempfile
import shutil

# 将项目根目录加入 sys.path，确保 src 包可导入
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from src.config import ConfigManager


def test_office_cache():
    """测试 Office 缓存机制：保存/加载/清除/状态"""
    print("=" * 60)
    print("测试 Office 缓存机制")
    print("=" * 60)

    # 使用独立临时目录，避免污染用户配置（ConfigManager 开发环境会探测
    # cwd 下的 PosterAgain_config.yaml，因此切换到临时目录）
    tmp_dir = tempfile.mkdtemp()
    old_cwd = os.getcwd()
    try:
        os.chdir(tmp_dir)
        config_file = os.path.join(tmp_dir, "test_config.yaml")
        config_manager = ConfigManager(config_file)

        # 初始状态：未缓存
        print("\n[1] 初始缓存状态...")
        status = config_manager.get_office_cache_status()
        print(f"    状态: {status}")
        assert config_manager.load_office_cache() is None, "初始状态不应有缓存"

        # 保存缓存
        print("\n[2] 保存 Office 缓存...")
        # path 必须指向真实存在的路径：load_office_cache 会校验它（CI 上没有
        # Office，写死的 C:\Program Files\... 会校验失败）。用临时目录自身。
        office_info = {
            "type": "microsoft",
            "name": "Microsoft Office PowerPoint",
            "path": tmp_dir,
            "version": "16.0",
        }
        ok = config_manager.save_office_cache(office_info)
        assert ok, "保存缓存失败"
        print("    OK 缓存已保存")

        # 加载缓存
        print("\n[3] 加载 Office 缓存...")
        loaded = config_manager.load_office_cache()
        assert loaded is not None, "加载缓存失败"
        assert loaded.get("type") == "microsoft", f"缓存类型不匹配: {loaded.get('type')}"
        assert loaded.get("name") == office_info["name"], "缓存名称不匹配"
        print(f"    OK 缓存加载成功: type={loaded.get('type')}, name={loaded.get('name')}")

        # 缓存状态
        print("\n[4] 缓存状态...")
        status = config_manager.get_office_cache_status()
        print(f"    状态: {status}")
        assert status.get("cached") is True, "缓存状态应为已缓存"

        # 清除缓存
        print("\n[5] 清除 Office 缓存...")
        ok = config_manager.clear_office_cache()
        assert ok, "清除缓存失败"
        assert config_manager.load_office_cache() is None, "清除后不应有缓存"
        print("    OK 缓存已清除")

        print("\n" + "=" * 60)
        print("测试通过！")
        print("=" * 60)
        return True

    finally:
        os.chdir(old_cwd)
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    success = test_office_cache()
    sys.exit(0 if success else 1)
