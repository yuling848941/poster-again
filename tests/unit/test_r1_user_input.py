#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试R=1情况下的用户输入
"""

import sys
import os
import tempfile
import pandas as pd

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_reader import DataReader
from src.memory_management.data_formatter import DataFormatter

def create_test_data_with_r1():
    """创建包含R=1情况的测试数据"""
    print("=" * 60)
    print("创建包含R=1的测试数据")
    print("=" * 60)

    # 测试数据：包含R=1, R=0.1, R=0.3的情况
    test_data = pd.DataFrame({
        "保单号": ["P001", "P002", "P003", "P004", "P005"],
        "SFYP2(不含短险续保)": [
            1000,   # R=0.1 -> 趸交FYP
            5000,   # R=1 -> 需要用户输入
            3000,   # R=0.3 -> 3年交SFYP
            6000,   # R=1 -> 需要用户输入
            2500    # R=0.5 -> 5年交SFYP
        ],
        "首年保费": [
            10000,  # R=0.1
            5000,   # R=1
            10000,  # R=0.3
            6000,   # R=1
            5000    # R=0.5
        ],
        "其他数据": ["A", "B", "C", "D", "E"]
    })

    print("\n测试数据:")
    print(test_data)

    # 计算预期结果
    print("\n预期结果:")
    for i, row in test_data.iterrows():
        sfyp2 = row["SFYP2(不含短险续保)"]
        premium = row["首年保费"]
        ratio = sfyp2 / premium

        if abs(ratio - 0.1) < 0.001:
            result = "趸交FYP"
        elif abs(ratio - 1.0) < 0.001:
            result = "待用户输入"
        else:
            years = int(round(ratio * 10))
            result = f"{years}年交SFYP"

        print(f"  行{i+1}: R={ratio:.2f} -> {result}")

    # 保存为Excel文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        test_file = tmp.name

    test_data.to_excel(test_file, index=False)
    print(f"\n测试文件已创建: {test_file}")

    return test_file, test_data

def test_with_thousand_separator():
    """测试带千位分隔符的情况"""
    print("\n" + "=" * 60)
    print("测试带千位分隔符的情况")
    print("=" * 60)

    # 创建测试数据
    test_file, original_data = create_test_data_with_r1()

    try:
        # 测试1: 使用DataFormatter直接计算（不使用千位分隔符）
        print("\n测试1: 直接使用DataFormatter计算")
        formatter = DataFormatter()
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            original_data,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费",
            target_column="承保趸期数据"
        )

        print(f"需要用户输入的行数: {len(pending_rows)}")
        if pending_rows:
            print("需要用户输入的行:")
            for row in pending_rows:
                row_idx = row['row_index']
                policy_num = row.get('policy_number', 'N/A')
                print(f"  行{row_idx+1}: 保单号={policy_num}")

        # 显示结果
        print("\n承保趸期数据结果:")
        cols_to_show = ["保单号", "SFYP2(不含短险续保)", "首年保费", "承保趸期数据"]
        print(result_df[cols_to_show].to_string(index=True))

        # 测试2: 使用DataReader加载（使用千位分隔符）
        print("\n测试2: 使用DataReader加载（使用千位分隔符）")
        reader = DataReader()
        success = reader.load_excel(test_file, use_thousand_separator=True, parent_widget=None)

        if success:
            print("DataReader加载成功")
            print(f"数据行数: {len(reader.data)}")

            # 检查承保趸期数据列
            if "承保趸期数据" in reader.data.columns:
                print("\n承保趸期数据列已生成")
                print(reader.data[cols_to_show].to_string(index=True))

                # 统计空值数量
                empty_count = (reader.data["承保趸期数据"] == "").sum()
                print(f"\n空值数量: {empty_count}")

                # 如果有R=1的行，应该有空值
                expected_empty = sum(1 for _, row in original_data.iterrows()
                                    if abs(row["SFYP2(不含短险续保)"] / row["首年保费"] - 1.0) < 0.001)
                print(f"期望空值数量（R=1的行数）: {expected_empty}")

                if empty_count == expected_empty:
                    print("[OK] 空值数量正确")
                else:
                    print(f"[FAIL] 空值数量不匹配，期望={expected_empty}，实际={empty_count}")
            else:
                print("[FAIL] 承保趸期数据列未生成")

        else:
            print("[FAIL] DataReader加载失败")

    finally:
        # 清理临时文件
        try:
            os.unlink(test_file)
            print(f"\n测试文件已清理: {test_file}")
        except:
            pass

def main():
    """主函数"""
    print("承保趸期数据 - R=1用户输入测试")
    print("=" * 60)

    try:
        # 测试带千位分隔符的情况
        test_with_thousand_separator()

        print("\n" + "=" * 60)
        print("测试总结")
        print("=" * 60)
        print("当R=1时，程序应该:")
        print("1. 检测到需要用户输入的行")
        print("2. 在对话框中显示这些行")
        print("3. 用户输入数值后，更新承保趸期数据列")
        print("4. 如果用户取消，允许用户选择是否继续")
        print("\n如果测试结果正确，说明修复有效")
        return 0

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
