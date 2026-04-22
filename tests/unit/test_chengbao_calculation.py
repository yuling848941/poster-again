#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
深入测试承保趸期数据计算过程
分析为什么检测到2行而不是1行
"""

import sys
import os
import pandas as pd

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.data_reader import DataReader
from src.memory_management.data_formatter import DataFormatter

def test_chengbao_calculation_details():
    """详细测试承保趸期数据计算"""
    print("="*70)
    print("深入测试承保趸期数据计算过程")
    print("="*70)

    excel_file = "KA承保快报.xlsx"

    if not os.path.exists(excel_file):
        print(f"\nERROR: 文件不存在 - {excel_file}")
        return

    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)

        print(f"\n[Step 1] 原始Excel数据分析")
        print(f"  数据形状: {df.shape}")
        print(f"  列名: {list(df.columns)}")

        # 找到SFYP2和首年保费列
        sfyp_col = None
        premium_col = None

        for col in df.columns:
            if 'SFYP2' in col and '不含短险续保' in col:
                sfyp_col = col
            if '首年保费' in col:
                premium_col = col

        print(f"\n[Step 2] 找到相关列")
        print(f"  SFYP2列: {sfyp_col}")
        print(f"  首年保费列: {premium_col}")

        if not sfyp_col or not premium_col:
            print(f"\nERROR: 未找到必要的列")
            return

        # 手动计算R值
        print(f"\n[Step 3] 手动计算R值")
        r_values = []
        for i in range(len(df)):
            sfyp = df.iloc[i][sfyp_col]
            premium = df.iloc[i][premium_col]

            if pd.isna(sfyp) or pd.isna(premium) or premium == 0:
                r = 0
            else:
                r = round(sfyp / premium, 3)

            r_values.append(r)

            policy_num = df.iloc[i]['保单号'] if '保单号' in df.columns else f"Row{i}"
            print(f"  行{i}: 保单号={policy_num}, SFYP2={sfyp}, 首年保费={premium}, R={r}")

        # 统计R=1的行
        r1_rows = [i for i, r in enumerate(r_values) if r == 1]
        print(f"\n  R=1的行: {r1_rows}")

        # 使用DataFormatter计算
        print(f"\n[Step 4] 使用DataFormatter计算承保趸期数据")
        formatter = DataFormatter()

        processed_data, pending_rows = formatter.calculate_chengbao_term_data(
            df.copy(),
            sfyp2_column=sfyp_col,
            premium_column=premium_col,
            target_column="承保趸期数据",
            policy_number_column="保单号"
        )

        print(f"  处理后数据形状: {processed_data.shape}")
        print(f"  需要用户输入的行: {pending_rows}")

        # 检查承保趸期数据列
        if "承保趸期数据" in processed_data.columns:
            print(f"\n[Step 5] 承保趸期数据列内容")
            chengbao_col = processed_data['承保趸期数据']
            for i, val in enumerate(chengbao_col):
                print(f"  行{i}: {val}")

            # 检查哪些行包含"年交SFYP"
            pending_from_data = []
            for i, val in enumerate(chengbao_col):
                if pd.notna(val) and '年交SFYP' in str(val):
                    pending_from_data.append(i)

            print(f"\n  数据中包含'年交SFYP'的行: {pending_from_data}")

        print(f"\n[Step 6] 对比结果")
        print(f"  手动计算R=1的行: {r1_rows}")
        print(f"  pending_rows: {pending_rows}")
        print(f"  数据中'年交SFYP'的行: {pending_from_data}")

        if len(pending_from_data) != len(r1_rows):
            print(f"\n  [问题] 数量不匹配!")
            print(f"    预期: {len(r1_rows)} 行")
            print(f"    实际: {len(pending_from_data)} 行")
        else:
            print(f"\n  [正常] 数量匹配")

        print(f"\n{'='*70}")

    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_chengbao_calculation_details()
