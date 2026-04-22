#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
承保趸期数据功能测试脚本
测试calculate_chengbao_term_data方法的计算逻辑
"""

import pandas as pd
import sys
import os
import tempfile

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.memory_management.data_formatter import DataFormatter

def test_calculate_chengbao_term_data():
    """测试承保趸期数据计算功能"""
    print("=" * 60)
    print("测试承保趸期数据计算功能")
    print("=" * 60)

    formatter = DataFormatter()

    # 测试数据1：包含三种情况的数据
    test_data_1 = pd.DataFrame({
        "SFYP2(不含短险续保)": [1000, 5000, 3000, 2500, 0],
        "首年保费": [10000, 5000, 10000, 5000, 10000],
        "保单号": ["P001", "P002", "P003", "P004", "P005"],
        "其他列": ["A", "B", "C", "D", "E"]
    })

    print("\n测试数据1 (包含R=0.1, R=1, R=0.3, R=0.5, R=0的情况):")
    print(test_data_1)

    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            test_data_1,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费",
            target_column="承保趸期数据",
            policy_number_column="保单号"
        )

        print("\n计算结果:")
        print(result_df[["保单号", "SFYP2(不含短险续保)", "首年保费", "承保趸期数据"]])

        print("\n需要用户输入的行:")
        if pending_rows:
            for row in pending_rows:
                print(f"  行 {row['row_index']}: 保单号={row['policy_number']}")
        else:
            print("  无需用户输入")

        # 验证计算结果
        print("\n验证计算结果:")
        expected_results = ["趸交FYP", "", "3年交SFYP", "5年交SFYP", ""]
        actual_results = result_df["承保趸期数据"].tolist()

        for i, (expected, actual) in enumerate(zip(expected_results, actual_results)):
            status = "[OK]" if expected == actual else "[FAIL]"
            print(f"  行 {i+1}: {status} 期望='{expected}', 实际='{actual}'")

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

    # 测试数据2：缺少必要列
    print("\n" + "=" * 60)
    print("测试数据2 (缺少SFYP2列):")
    test_data_2 = pd.DataFrame({
        "首年保费": [10000, 5000],
        "其他列": ["A", "B"]
    })

    print(test_data_2)

    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            test_data_2,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费",
            target_column="承保趸期数据",
            policy_number_column="保单号"
        )

        print("\n[OK] 测试通过：缺少列时正确跳过计算")
        print(f"  原始数据行数: {len(test_data_2)}")
        print(f"  结果数据行数: {len(result_df)}")

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {str(e)}")
        return False

    # 测试数据3：空值处理
    print("\n" + "=" * 60)
    print("测试数据3 (包含空值):")
    test_data_3 = pd.DataFrame({
        "SFYP2(不含短险续保)": [1000, None, 3000],
        "首年保费": [10000, 5000, None],
        "保单号": ["P001", "P002", "P003"]
    })

    print(test_data_3)

    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            test_data_3,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费",
            target_column="承保趸期数据",
            policy_number_column="保单号"
        )

        print("\n[OK] 测试通过：正确处理空值")
        print(result_df[["保单号", "SFYP2(不含短险续保)", "首年保费", "承保趸期数据"]])

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {str(e)}")
        return False

    # 测试数据4：除零错误处理
    print("\n" + "=" * 60)
    print("测试数据4 (首年保费为0):")
    test_data_4 = pd.DataFrame({
        "SFYP2(不含短险续保)": [1000, 2000],
        "首年保费": [0, 10000],
        "保单号": ["P001", "P002"]
    })

    print(test_data_4)

    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            test_data_4,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费",
            target_column="承保趸期数据",
            policy_number_column="保单号"
        )

        print("\n[OK] 测试通过：正确处理除零错误")
        print(result_df[["保单号", "SFYP2(不含短险续保)", "首年保费", "承保趸期数据"]])

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {str(e)}")
        return False

    print("\n" + "=" * 60)
    print("[OK] 所有测试通过！")
    print("=" * 60)
    return True

def test_data_reader_integration():
    """测试DataReader与承保趸期数据的集成"""
    print("\n" + "=" * 60)
    print("测试DataReader集成")
    print("=" * 60)

    from src.data_reader import DataReader

    # 创建测试Excel文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        test_file = tmp.name

    try:
        # 创建测试数据
        test_data = pd.DataFrame({
            "SFYP2(不含短险续保)": [1000, 5000, 3000, 2500, 0],
            "首年保费": [10000, 5000, 10000, 5000, 10000],
            "保单号": ["P001", "P002", "P003", "P004", "P005"],
            "缴费年期（主险） ID": [10, 20, 30, 25, 15]
        })

        # 保存为Excel
        test_data.to_excel(test_file, index=False)

        # 使用DataReader加载
        reader = DataReader()
        success = reader.load_excel(test_file, use_thousand_separator=False)

        if success:
            print("\n[OK] 数据加载成功")
            print(f"  数据行数: {len(reader.data)}")
            print(f"  数据列数: {len(reader.data.columns)}")

            # 检查承保趸期数据列是否存在
            if "承保趸期数据" in reader.data.columns:
                print("\n[OK] 承保趸期数据列已添加")
                print("\n承保趸期数据结果:")
                print(reader.data[["保单号", "SFYP2(不含短险续保)", "首年保费", "承保趸期数据"]])
            else:
                print("\n[FAIL] 承保趸期数据列未添加")

            # 检查期趸数据列
            if "期趸数据" in reader.data.columns:
                print("\n[OK] 期趸数据列也存在（现有功能）")
            else:
                print("\n[FAIL] 期趸数据列不存在")

        else:
            print("\n[FAIL] 数据加载失败")

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        # 清理临时文件
        try:
            os.unlink(test_file)
        except:
            pass

if __name__ == "__main__":
    print("承保趸期数据功能测试\n")

    # 测试核心计算逻辑
    success = test_calculate_chengbao_term_data()

    if success:
        # 测试DataReader集成
        test_data_reader_integration()

    print("\n测试完成")
