#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试缓存机制
"""

import sys
import os
import time

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.config import ConfigManager
from src.gui.office_suite_detector import OfficeSuiteManager

def test_cache():
    """测试缓存机制"""
    print("=" * 60)
    print("测试启动性能优化 - 缓存机制")
    print("=" * 60)

    # 创建配置管理器
    config_manager = ConfigManager()

    # 第一次初始化（应该会进行实际检测）
    print("\n[1] 第一次初始化办公套件管理器...")
    start = time.time()
    manager = OfficeSuiteManager(config_manager)
    manager.initialize()
    elapsed1 = time.time() - start
    print(f"    耗时: {elapsed1:.3f}秒")

    # 检查缓存
    print("\n[2] 检查缓存...")
    cache_data = config_manager.load_office_cache()
    if cache_data:
        print("    ✓ 缓存已保存")
        print(f"    缓存时间: {cache_data.get('timestamp', 'N/A')}")
        print(f"    有效期至: {cache_data.get('valid_until', 'N/A')}")
        print(f"    可用套件: {cache_data.get('available_suites', [])}")
        print(f"    当前选择: {cache_data.get('current_suite', 'N/A')}")
    else:
        print("    ✗ 未找到缓存")

    # 第二次初始化（应该使用缓存）
    print("\n[3] 第二次初始化办公套件管理器（应使用缓存）...")
    start = time.time()
    manager2 = OfficeSuiteManager(config_manager)
    manager2.initialize()
    elapsed2 = time.time() - start
    print(f"    耗时: {elapsed2:.3f}秒")

    # 性能对比
    print("\n[4] 性能对比:")
    if elapsed1 > 0 and elapsed2 > 0:
        improvement = (elapsed1 - elapsed2) / elapsed1 * 100
        print(f"    首次初始化: {elapsed1:.3f}秒")
        print(f"    使用缓存:   {elapsed2:.3f}秒")
        print(f"    性能提升:   {improvement:.1f}%")

    # 测试清除缓存
    print("\n[5] 测试清除缓存...")
    if config_manager.clear_office_cache():
        print("    ✓ 缓存已清除")
    else:
        print("    ✗ 清除缓存失败")

    # 验证缓存已清除
    print("\n[6] 验证缓存已清除...")
    cache_data = config_manager.load_office_cache()
    if cache_data:
        print("    ✗ 缓存仍然存在")
    else:
        print("    ✓ 缓存已成功清除")

    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

if __name__ == "__main__":
    test_cache()
