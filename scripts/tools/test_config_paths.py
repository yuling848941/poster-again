#!/usr/bin/env python
"""
测试配置路径记忆功能
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'src'))

from config_manager import ConfigManager

def test_config_path_memory():
    """测试配置路径记忆功能"""
    print("=== Test Config Path Memory Feature ===\n")

    # 创建测试配置管理器
    test_config_file = "config/test_paths_config.yaml"
    cm = ConfigManager(test_config_file)

    # 测试获取默认路径
    print("1. Get default config paths:")
    save_dir = cm.get("paths.last_config_save_dir", "")
    load_dir = cm.get("paths.last_config_load_dir", "")
    print(f"   - Save dir: {save_dir or '(not set)'}")
    print(f"   - Load dir: {load_dir or '(not set)'}")
    print()

    # 测试设置路径
    print("2. Set test paths:")
    cm.set("paths.last_config_save_dir", "C:/Users/Test/Documents/Configs")
    cm.set("paths.last_config_load_dir", "D:/Project/Configs")
    print(f"   - Save dir: C:/Users/Test/Documents/Configs")
    print(f"   - Load dir: D:/Project/Configs")
    print()

    # 测试保存配置
    print("3. Save config:")
    success = cm.save_config()
    if success:
        print(f"   [OK] Config saved to: {test_config_file}")
    else:
        print(f"   [FAIL] Config save failed")
    print()

    # 测试读取配置
    print("4. Reload config:")
    cm2 = ConfigManager(test_config_file)
    save_dir = cm2.get("paths.last_config_save_dir", "")
    load_dir = cm2.get("paths.last_config_load_dir", "")
    print(f"   - Save dir: {save_dir}")
    print(f"   - Load dir: {load_dir}")
    print()

    # 验证路径正确性
    print("5. Verification:")
    if save_dir == "C:/Users/Test/Documents/Configs" and load_dir == "D:/Project/Configs":
        print("   [OK] Config path memory feature works!")
    else:
        print("   [FAIL] Config path memory feature error")

    # 清理测试文件
    if os.path.exists(test_config_file):
        os.remove(test_config_file)
        print(f"\nCleaned up test file: {test_config_file}")

if __name__ == "__main__":
    test_config_path_memory()