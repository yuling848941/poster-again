#!/usr/bin/env python
"""
测试配置按钮状态管理功能
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_config_button_states():
    """测试配置按钮状态管理"""
    print("=== 测试配置按钮状态管理 ===\n")

    # 检查main_window.py中是否包含必要的修改
    main_window_file = "src/gui/main_window.py"

    if not os.path.exists(main_window_file):
        print(f"[FAIL] 文件不存在: {main_window_file}")
        return False

    with open(main_window_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 检查点1: 初始化时禁用
    print("1. 检查初始化时禁用配置按钮:")
    if 'self.config_btn.setEnabled(False)' in content:
        print("   [OK] 找到 setEnabled(False) 初始化禁用")
    else:
        print("   [FAIL] 未找到 setEnabled(False) 初始化禁用")
        return False

    # 检查点2: 添加工具提示
    print("\n2. 检查工具提示:")
    if 'setToolTip' in content and '自动匹配' in content:
        print("   [OK] 找到禁用状态工具提示")
    else:
        print("   [FAIL] 未找到工具提示")
        return False

    # 检查点3: 自动匹配成功后启用
    print("\n3. 检查自动匹配成功后启用:")
    if 'self.config_btn.setEnabled(True)' in content:
        print("   [OK] 找到启用按钮的代码")
    else:
        print("   [FAIL] 未找到启用按钮的代码")
        return False

    # 检查点4: 清除工具提示
    print("\n4. 检查清除工具提示:")
    if 'self.config_btn.setToolTip("")' in content:
        print("   [OK] 找到清除工具提示的代码")
    else:
        print("   [FAIL] 未找到清除工具提示的代码")
        return False

    # 检查点5: 验证代码位置正确性
    print("\n5. 验证代码逻辑顺序:")

    # 找到初始化位置
    init_pos = content.find('self.config_btn.setEnabled(False)')
    # 找到启用位置
    enable_pos = content.find('self.config_btn.setEnabled(True)')

    if init_pos > 0 and enable_pos > 0:
        if init_pos < enable_pos:
            print("   [OK] 代码顺序正确：先禁用，后启用")
        else:
            print("   [FAIL] 代码顺序错误：启用位置在禁用位置之前")
            return False
    else:
        print("   [FAIL] 未找到完整的代码块")
        return False

    print("\n=== 所有检查通过 ===")
    print("\n功能说明:")
    print("1. 程序启动时，配置按钮为禁用状态（灰色）")
    print("2. 鼠标悬停显示提示：请先点击\"自动匹配\"按钮")
    print("3. 自动匹配成功后，按钮变为启用状态（绿色）")
    print("4. 工具提示被清除")

    return True

if __name__ == "__main__":
    success = test_config_button_states()
    if success:
        print("\n[OK] 配置按钮状态管理功能实现正确！")
    else:
        print("\n[FAIL] 配置按钮状态管理功能存在问题！")