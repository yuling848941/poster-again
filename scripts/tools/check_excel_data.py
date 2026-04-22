#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
检查KA承保快报.xlsx中的实际数据
"""

import sys
import os
import pandas as pd

# 添加src路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def check_excel_data():
    excel_file = "KA承保快报.xlsx"

    if not os.path.exists(excel_file):
        print(f"ERROR: 文件不存在 - {excel_file}")
        return

    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)

        print("=" * 70)
        print(f"Excel文件内容: {excel_file}")
        print("=" * 70)

        # 显示所有列名
        print(f"\n列数: {len(df.columns)}")
        print(f"行数: {len(df)}")

        print(f"\n列名:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1}. {col}")

        # 检查是否有SFYP2和首年保费列
        has_sfyp = any('SFYP2' in col for col in df.columns)
        has_premium = any('首年保费' in col for col in df.columns)

        print(f"\n数据检查:")
        print(f"  包含SFYP2列: {'是' if has_sfyp else '否'}")
        print(f"  包含首年保费列: {'是' if has_premium else '否'}")

        # 显示前5行数据
        print(f"\n前5行数据:")
        print(df.head().to_string())

        # 检查是否会有承保趸期数据列
        # 如果没有这个列，程序会计算生成
        has_chengbao = '承保趸期数据' in df.columns

        print(f"\n检查承保趸期数据:")
        print(f"  已有承保趸期数据列: {'是' if has_chengbao else '否'}")

        if has_chengbao:
            chengbao_data = df['承保趸期数据']
            print(f"\n承保趸期数据列内容:")
            for i, value in enumerate(chengbao_data):
                print(f"  行{i}: {value}")

            # 统计需要用户输入的行
            user_input_rows = []
            for i, value in enumerate(chengbao_data):
                if pd.notna(value) and '年交SFYP' in str(value):
                    user_input_rows.append(i)

            print(f"\n需要用户输入的行: {len(user_input_rows)} 个")
            for row in user_input_rows:
                policy_num = df.iloc[row]['保单号'] if '保单号' in df.columns else 'N/A'
                print(f"  行{row}: 保单号={policy_num}, 值={chengbao_data[row]}")

        # 检查模拟承保趸期计算会产生多少用户输入项
        print(f"\n\n模拟承保趸期数据计算...")
        print(f"  检查SFYP2(不含短险续保)列...")

        sfyp_col = None
        premium_col = None

        for col in df.columns:
            if 'SFYP2' in col and '不含短险续保' in col:
                sfyp_col = col
            if '首年保费' in col:
                premium_col = col

        if sfyp_col and premium_col:
            print(f"    找到列: {sfyp_col}")
            print(f"    找到列: {premium_col}")

            # 模拟计算：R = SFYP2 / 首年保费
            r_values = []
            for i in range(len(df)):
                sfyp = df.iloc[i][sfyp_col]
                premium = df.iloc[i][premium_col]

                if pd.isna(sfyp) or pd.isna(premium) or premium == 0:
                    r = 0
                else:
                    r = round(sfyp / premium, 3)

                r_values.append(r)
                print(f"    行{i}: SFYP2={sfyp}, 首年保费={premium}, R={r}")

            # 统计R=1的行（需要用户输入）
            r1_rows = [i for i, r in enumerate(r_values) if r == 1]
            print(f"\n  R=1的行（需要用户输入）: {len(r1_rows)} 个")
            for row in r1_rows:
                policy_num = df.iloc[row]['保单号'] if '保单号' in df.columns else 'N/A'
                print(f"    行{row}: 保单号={policy_num}, R={r_values[row]}")

        print("\n" + "=" * 70)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_excel_data()
