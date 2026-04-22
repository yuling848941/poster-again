#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证双弹框问题的根本原因
"""

import sys
import os
import pandas as pd

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.memory_management.data_formatter import DataFormatter

def verify_root_cause():
    """验证根本原因"""
    print("="*70)
    print("验证双弹框问题的根本原因")
    print("="*70)

    excel_file = "KA承保快报.xlsx"

    # 读取Excel并计算承保趸期数据
    df = pd.read_excel(excel_file)
    formatter = DataFormatter()

    # 计算承保趸期数据
    sfyp_col = "SFYP2(不含短险续保)"
    premium_col = "首年保费"

    processed_data, pending_rows = formatter.calculate_chengbao_term_data(
        df.copy(),
        sfyp2_column=sfyp_col,
        premium_column=premium_col,
        target_column="承保趸期数据"
    )

    # 模拟配置管理器的检测逻辑
    print("\n[Step 1] 模拟配置管理器的检测逻辑")
    chengbao_data_column = processed_data['承保趸期数据']

    user_input_rows_wrong = []  # 错误的逻辑：检测包含"年交SFYP"的行
    for i, value in enumerate(chengbao_data_column):
        if value and "年交SFYP" in value:
            user_input_rows_wrong.append(i)

    print(f"  错误的检测逻辑（包含'年交SFYP'）:")
    print(f"    检测到的行: {user_input_rows_wrong}")
    print(f"    行数: {len(user_input_rows_wrong)}")

    # 显示每行的值
    print(f"\n[Step 2] 各行的承保趸期数据值")
    for i, value in enumerate(chengbao_data_column):
        print(f"  行{i}: '{value}'")

        # 计算R值
        sfyp = df.iloc[i][sfyp_col]
        premium = df.iloc[i][premium_col]
        r = round(sfyp / premium, 3) if premium != 0 else 0

        if value and "年交SFYP" in value:
            print(f"      ↑ 被错误检测为需要用户输入（因为包含'年交SFYP'）")
        elif not value:
            print(f"      ↑ 正确：空值，需要用户输入")
        else:
            print(f"      ↑ 正确：已自动计算")

    # 正确的检测逻辑
    print(f"\n[Step 3] 正确的检测逻辑（空值行）")
    user_input_rows_correct = []
    for i, value in enumerate(chengbao_data_column):
        if not value:  # 空值才需要用户输入
            user_input_rows_correct.append(i)

    print(f"  正确的检测逻辑（空值）:")
    print(f"    检测到的行: {user_input_rows_correct}")
    print(f"    行数: {len(user_input_rows_correct)}")

    # 对比结果
    print(f"\n[Step 4] 对比结果")
    print(f"  实际R=1的行: [0]")
    print(f"  错误的检测: {user_input_rows_wrong} ({len(user_input_rows_wrong)}行)")
    print(f"  正确的检测: {user_input_rows_correct} ({len(user_input_rows_correct)}行)")

    if len(user_input_rows_wrong) > len(user_input_rows_correct):
        print(f"\n  [问题确认] 错误的检测逻辑多检测了 {len(user_input_rows_wrong) - len(user_input_rows_correct)} 行")
        print(f"  原因: 将自动计算的'5年交SFYP'误判为需要用户输入")

    print(f"\n{'='*70}")
    print("结论：")
    print("  双弹框问题的根本原因是检测逻辑错误")
    print("  错误的逻辑：检测包含'年交SFYP'的行")
    print("  正确的逻辑：检测空值行")
    print("{'='*70}")

if __name__ == "__main__":
    verify_root_cause()
