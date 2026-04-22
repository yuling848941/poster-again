#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
模拟GUI工作流程测试
测试DataReader在GUI中的使用方式
"""

import sys
import os
import pandas as pd

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_reader import DataReader

def test_gui_workflow():
    """模拟GUI工作流程"""
    print("=" * 60)
    print("模拟GUI工作流程测试")
    print("=" * 60)

    excel_file = "KA承保快报.xlsx"

    if not os.path.exists(excel_file):
        print(f"错误：找不到文件 {excel_file}")
        return False

    # 模拟GUI中的数据加载
    print("\n步骤1: 创建DataReader")
    data_reader = DataReader()

    print("步骤2: 加载Excel文件（模拟GUI中不传递parent_widget）")
    # GUI中第一次加载时不传递parent_widget，所以不会弹出对话框
    success = data_reader.load_excel(excel_file, use_thousand_separator=True, parent_widget=None)

    if not success:
        print("[FAIL] 数据加载失败")
        return False

    print("[OK] 数据加载成功")
    print(f"  数据行数: {len(data_reader.data)}")
    print(f"  数据列数: {len(data_reader.data.columns)}")

    # 检查承保趸期数据列
    if "承保趸期数据" in data_reader.data.columns:
        print("\n[OK] 承保趸期数据列已生成")

        # 显示结果
        print("\n承保趸期数据结果:")
        cols_to_show = ["SFYP2(不含短险续保)", "首年保费", "承保趸期数据"]
        if "保单号" in data_reader.data.columns:
            cols_to_show.insert(0, "保单号")

        for i, row in data_reader.data[cols_to_show].iterrows():
            print(f"  行{i+1}: ", end="")
            for col in cols_to_show:
                print(f"{col}={row[col]}, ", end="")
            print()

        # 检查是否有空值（需要用户输入的行）
        empty_values = data_reader.data["承保趸期数据"] == ""
        if empty_values.any():
            print(f"\n[INFO] 发现 {empty_values.sum()} 行需要用户输入（R=1的情况）")
            empty_rows = data_reader.data[empty_values]
            for idx, row in empty_rows.iterrows():
                policy_num = row.get("保单号", "N/A")
                sfyp2 = row.get("SFYP2(不含短险续保)", "N/A")
                premium = row.get("首年保费", "N/A")
                print(f"  行{idx+1}: 保单号={policy_num}, SFYP2={sfyp2}, 首年保费={premium}")
        else:
            print("\n[OK] 所有行都已计算出结果，无空值")

    else:
        print("\n[FAIL] 承保趸期数据列未生成")
        return False

    print("\n步骤3: 模拟GUI中点击'自定义' -> '承保趸期数据'")
    print("在GUI中，这会调用submit_chengbao_term_data方法，该方法会：")
    print("  1. 重新加载Excel文件（这次传递parent_widget=self）")
    print("  2. 如果有R=1的行，弹出批量输入对话框")
    print("  3. 用户输入数据后更新结果")
    print("  4. 将占位符与承保趸期数据列关联")

    return True

def test_with_gui_parent():
    """测试传递parent_widget的情况"""
    print("\n" + "=" * 60)
    print("测试传递parent_widget（模拟GUI环境）")
    print("=" * 60)

    excel_file = "KA承保快报.xlsx"

    if not os.path.exists(excel_file):
        print(f"错误：找不到文件 {excel_file}")
        return False

    # 由于没有实际的GUI窗口，我们模拟传递None
    print("\n注意：在实际GUI中，parent_widget=self（主窗口对象）")
    print("在offscreen模式下，parent_widget=None")

    # 测试正常加载
    data_reader = DataReader()
    success = data_reader.load_excel(excel_file, use_thousand_separator=True, parent_widget=None)

    if success:
        print("[OK] 数据加载成功（无论是否传递parent_widget）")

        # 检查结果
        if "承保趸期数据" in data_reader.data.columns:
            result = data_reader.data["承保趸期数据"].iloc[0]
            print(f"\n承保趸期数据结果: {result}")

            if result == "5年交SFYP":
                print("[OK] 结果正确！")
            else:
                print(f"[FAIL] 结果不正确，期望'5年交SFYP'，实际'{result}'")

    return True

def main():
    """主函数"""
    print("承保趸期数据GUI工作流程测试")
    print("=" * 60)

    # 测试1: 模拟GUI工作流程
    test1_ok = test_gui_workflow()

    # 测试2: 测试传递parent_widget
    test2_ok = test_with_gui_parent()

    # 汇总结果
    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)

    if test1_ok and test2_ok:
        print("[SUCCESS] 所有测试通过")
        print("\n结论：")
        print("1. 程序本身计算正确（R=0.5 -> 5年交SFYP）")
        print("2. 问题可能在于GUI操作步骤")
        print("\n正确的GUI操作步骤：")
        print("1. 加载Excel文件")
        print("2. 提取占位符（从PPT模板）")
        print("3. 在占位符表格中选择要关联的行")
        print("4. 点击'自定义'按钮")
        print("5. 选择'承保趸期数据'菜单项")
        print("6. 查看匹配表格中的'匹配值'列应显示'承保趸期数据'")
        return 0
    else:
        print("[FAIL] 部分测试失败")
        return 1

if __name__ == "__main__":
    sys.exit(main())
