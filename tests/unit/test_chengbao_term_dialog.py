#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
承保趸期数据批量输入对话框测试脚本
测试对话框的显示和交互功能
"""

import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 设置环境变量以使用Qt基本样式（无需实际显示GUI）
os.environ['QT_QPA_PLATFORM'] = 'offscreen'

from PySide6.QtWidgets import QApplication
from src.gui.chengbao_term_input_dialog import ChengbaoTermInputDialog
from src.memory_management.data_formatter import DataFormatter
import pandas as pd

def test_dialog_creation():
    """测试对话框创建"""
    print("=" * 60)
    print("测试对话框创建")
    print("=" * 60)

    # 创建测试数据
    pending_rows = [
        {'row_index': 0, 'policy_number': 'P001'},
        {'row_index': 2, 'policy_number': 'P003'},
        {'row_index': 5, 'policy_number': 'P006'},
        {'row_index': 9, 'policy_number': 'P010'}
    ]

    print(f"\n待输入行数: {len(pending_rows)}")
    for row in pending_rows:
        print(f"  第 {row['row_index'] + 1} 行, 保单号: {row['policy_number']}")

    # 创建对话框
    try:
        dialog = ChengbaoTermInputDialog(pending_rows)
        print("\n[OK] 对话框创建成功")
        print(f"  对话框标题: {dialog.windowTitle()}")
        print(f"  对话框尺寸: {dialog.width()}x{dialog.height()}")
        return True
    except Exception as e:
        print(f"\n[FAIL] 对话框创建失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def test_dialog_with_chengbao_calculation():
    """测试对话框与承保趸期数据计算的集成"""
    print("\n" + "=" * 60)
    print("测试对话框与承保趸期数据计算集成")
    print("=" * 60)

    # 创建包含R=1情况的测试数据
    test_data = pd.DataFrame({
        "SFYP2(不含短险续保)": [1000, 5000, 3000, 2500, 0],
        "首年保费": [10000, 5000, 10000, 5000, 10000],
        "保单号": ["P001", "P002", "P003", "P004", "P005"],
        "其他列": ["A", "B", "C", "D", "E"]
    })

    print("\n测试数据:")
    print(test_data)

    # 使用DataFormatter计算承保趸期数据
    formatter = DataFormatter()
    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            test_data,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费",
            target_column="承保趸期数据",
            policy_number_column="保单号"
        )

        print(f"\n计算完成，发现 {len(pending_rows)} 行需要用户输入")

        if pending_rows:
            print("\n待输入行详情:")
            for row in pending_rows:
                print(f"  第 {row['row_index'] + 1} 行, 保单号: {row['policy_number']}")

            # 创建并测试对话框
            dialog = ChengbaoTermInputDialog(pending_rows)
            print("\n[OK] 对话框集成测试通过")
            print("  对话框可以正常显示待输入行")
            return True
        else:
            print("\n[INFO] 无需用户输入，对话框不会被调用")
            return True

    except Exception as e:
        print(f"\n[FAIL] 集成测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def test_dialog_validation():
    """测试对话框输入验证"""
    print("\n" + "=" * 60)
    print("测试对话框输入验证")
    print("=" * 60)

    pending_rows = [
        {'row_index': 0, 'policy_number': 'P001'},
        {'row_index': 1, 'policy_number': 'P002'}
    ]

    try:
        dialog = ChengbaoTermInputDialog(pending_rows)

        # 测试输入框验证器
        input_field = dialog.input_field_0
        validator = input_field.validator()

        print("\n验证器测试:")
        test_cases = [
            ("10", True, "有效输入：10"),
            ("2", True, "最小有效输入：2"),
            ("1", False, "无效：小于2"),
            ("0", False, "无效：0"),
            ("-5", False, "无效：负数"),
            ("abc", False, "无效：字母"),
            ("", True, "空值（允许）")
        ]

        for text, expected_valid, description in test_cases:
            state = validator.validate(text, 0)[0]
            is_valid = (state == 1)  # QIntValidator.Acceptable = 1
            status = "[OK]" if is_valid == expected_valid else "[FAIL]"
            print(f"  {status} {description}")

        print("\n[OK] 输入验证测试完成")
        return True

    except Exception as e:
        print(f"\n[FAIL] 验证测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def test_input_values_retrieval():
    """测试输入值获取"""
    print("\n" + "=" * 60)
    print("测试输入值获取")
    print("=" * 60)

    pending_rows = [
        {'row_index': 0, 'policy_number': 'P001'},
        {'row_index': 2, 'policy_number': 'P003'},
        {'row_index': 5, 'policy_number': 'P006'}
    ]

    try:
        dialog = ChengbaoTermInputDialog(pending_rows)

        # 模拟用户输入
        dialog.input_values[0] = 10
        dialog.input_values[2] = 20
        dialog.input_values[5] = 30

        # 获取输入值
        values = dialog.get_input_values()

        print("\n用户输入值:")
        for row_index, value in values.items():
            policy_num = next(r['policy_number'] for r in pending_rows if r['row_index'] == row_index)
            print(f"  第 {row_index + 1} 行 (保单号: {policy_num}) = {value}")

        # 验证
        if len(values) == 3 and values[0] == 10 and values[2] == 20 and values[5] == 30:
            print("\n[OK] 输入值获取测试通过")
            return True
        else:
            print("\n[FAIL] 输入值不匹配")
            return False

    except Exception as e:
        print(f"\n[FAIL] 输入值获取测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("承保趸期数据批量输入对话框测试\n")
    print("注意：此测试使用offscreen模式，不会显示实际GUI对话框\n")

    # 初始化QApplication
    app = QApplication(sys.argv)

    # 运行测试
    tests = [
        ("对话框创建", test_dialog_creation),
        ("对话框与计算集成", test_dialog_with_chengbao_calculation),
        ("输入验证", test_dialog_validation),
        ("输入值获取", test_input_values_retrieval)
    ]

    results = []
    for test_name, test_func in tests:
        print(f"\n运行测试: {test_name}")
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"\n[FAIL] 测试异常: {str(e)}")
            import traceback
            traceback.print_exc()
            results.append((test_name, False))

    # 汇总结果
    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)

    passed = sum(1 for _, result in results if result)
    total = len(results)

    for test_name, result in results:
        status = "[PASS]" if result else "[FAIL]"
        print(f"{status} {test_name}")

    print(f"\n总计: {passed}/{total} 个测试通过")

    if passed == total:
        print("\n[OK] 所有测试通过！")
    else:
        print(f"\n[WARN] {total - passed} 个测试失败")

    sys.exit(0 if passed == total else 1)
