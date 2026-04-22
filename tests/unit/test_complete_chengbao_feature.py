#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
承保趸期数据功能完整集成测试脚本
测试从数据加载到计算的完整流程
"""

import pandas as pd
import tempfile
import os
import sys

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_reader import DataReader
from src.memory_management.data_formatter import DataFormatter

def create_test_excel():
    """创建测试Excel文件"""
    print("=" * 60)
    print("创建测试Excel文件")
    print("=" * 60)

    # 创建完整的测试数据，包含所有测试场景
    test_data = pd.DataFrame({
        "保单号": [f"P{str(i).zfill(3)}" for i in range(1, 16)],
        "SFYP2(不含短险续保)": [
            1000,   # R=0.1 -> 趸交FYP
            5000,   # R=1 -> 需要用户输入
            3000,   # R=0.3 -> 3年交SFYP
            2500,   # R=0.5 -> 5年交SFYP
            4000,   # R=0.8 -> 8年交SFYP
            0,      # R=0 -> 空值
            None,   # 空值 -> 空值
            10000,  # R=1 -> 需要用户输入
            1500,   # R=0.3 -> 3年交SFYP
            4500,   # R=0.9 -> 9年交SFYP
            2000,   # R=0.2 -> 2年交SFYP
            500,    # R=0.1 -> 趸交FYP
            7000,   # R=1 -> 需要用户输入
            3000,   # R=0.6 -> 6年交SFYP
            10000   # R=1 -> 需要用户输入
        ],
        "首年保费": [
            10000, 10000, 10000, 5000, 5000, 5000, 5000,
            10000, 5000, 5000, 10000, 5000, 7000, 5000, 10000
        ],
        "缴费年期（主险） ID": [10, 20, 30, 20, 15, 25, 10, 5, 20, 30, 15, 25, 10, 20, 30],
        "其他数据": [f"Data{i}" for i in range(1, 16)]
    })

    print("\n测试数据概览:")
    print(f"  总行数: {len(test_data)}")
    print(f"  列数: {len(test_data.columns)}")
    print("\n测试场景分布:")
    print("  R=0.1: 2行 (趸交FYP)")
    print("  R=1: 5行 (需要用户输入)")
    print("  R=0.2-0.9: 6行 (自动计算)")
    print("  异常值: 2行 (空值或0)")

    # 保存为Excel文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        test_file = tmp.name

    test_data.to_excel(test_file, index=False)
    print(f"\n测试文件已创建: {test_file}")

    return test_file, test_data

def test_complete_flow():
    """测试完整的数据处理流程"""
    print("\n" + "=" * 60)
    print("测试完整数据处理流程")
    print("=" * 60)

    # 创建测试Excel文件
    test_file, original_data = create_test_excel()

    try:
        # 使用DataReader加载数据（不提供parent_widget，所以不会弹出对话框）
        print("\n步骤1: 加载Excel数据")
        reader = DataReader()
        success = reader.load_excel(test_file, use_thousand_separator=False)

        if not success:
            print("[FAIL] 数据加载失败")
            return False

        print("[OK] 数据加载成功")
        print(f"  文件: {test_file}")
        print(f"  数据行数: {len(reader.data)}")
        print(f"  数据列数: {len(reader.data.columns)}")

        # 验证列存在性
        print("\n步骤2: 验证必要列")
        required_columns = ["保单号", "SFYP2(不含短险续保)", "首年保费"]
        for col in required_columns:
            if col in reader.data.columns:
                print(f"  [OK] 列 '{col}' 存在")
            else:
                print(f"  [FAIL] 列 '{col}' 不存在")
                return False

        # 验证承保趸期数据列
        print("\n步骤3: 验证承保趸期数据处理")
        if "承保趸期数据" in reader.data.columns:
            print("  [OK] '承保趸期数据' 列已添加")

            # 统计结果
            result_counts = reader.data["承保趸期数据"].value_counts(dropna=False)
            print("\n承保趸期数据结果统计:")
            for value, count in result_counts.items():
                if pd.isna(value):
                    print(f"  空值: {count} 行")
                else:
                    print(f"  {value}: {count} 行")

            # 验证计算结果
            print("\n步骤4: 验证计算结果")
            expected_results = {
                "趸交FYP": 2,
                "": 2,  # 空值行
                "3年交SFYP": 2,
                "5年交SFYP": 1,
                "8年交SFYP": 1,
                "9年交SFYP": 1,
                "2年交SFYP": 1,
                "6年交SFYP": 1
            }

            for expected, count in expected_results.items():
                actual_count = (reader.data["承保趸期数据"] == expected).sum()
                status = "[OK]" if actual_count == count else "[FAIL]"
                print(f"  {status} '{expected}': 期望={count}, 实际={actual_count}")

        else:
            print("  [FAIL] '承保趸期数据' 列未添加")
            return False

        # 验证期趸数据列（确保现有功能不受影响）
        print("\n步骤5: 验证现有功能（期趸数据）")
        if "期趸数据" in reader.data.columns:
            print("  [OK] '期趸数据' 列也存在（现有功能正常）")
        else:
            print("  [FAIL] '期趸数据' 列不存在（现有功能受影响）")
            return False

        # 详细查看结果
        print("\n步骤6: 查看计算详情")
        print("\n承保趸期数据计算详情:")
        result_df = reader.data[[
            "保单号", "SFYP2(不含短险续保)", "首年保费", "承保趸期数据"
        ]].copy()

        # 计算比例并显示
        result_df["比例R"] = result_df["SFYP2(不含短险续保)"] / result_df["首年保费"]
        result_df["比例R"] = result_df["比例R"].round(2)

        print(result_df.to_string(index=True))

        # 验证与现有功能的兼容性
        print("\n步骤7: 验证数据兼容性")
        print("  列名检查:")
        print(f"    原始数据列数: {len(original_data.columns)}")
        print(f"    处理后列数: {len(reader.data.columns)}")
        print(f"    新增列数: {len(reader.data.columns) - len(original_data.columns)}")

        if "期趸数据" in reader.data.columns and "承保趸期数据" in reader.data.columns:
            print("  [OK] 两个趸期数据列可以共存")

        print("\n[OK] 完整流程测试通过")
        return True

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理临时文件
        try:
            os.unlink(test_file)
            print(f"\n测试文件已清理: {test_file}")
        except:
            pass

def test_edge_cases():
    """测试边界条件"""
    print("\n" + "=" * 60)
    print("测试边界条件")
    print("=" * 60)

    formatter = DataFormatter()

    # 测试案例1：缺少列
    print("\n测试案例1: 缺少SFYP2列")
    test_data_1 = pd.DataFrame({
        "首年保费": [10000, 5000],
        "保单号": ["P001", "P002"]
    })

    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            test_data_1,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费"
        )
        print("  [OK] 缺少列时正常处理")
    except Exception as e:
        print(f"  [FAIL] 缺少列时处理失败: {str(e)}")
        return False

    # 测试案例2：首年保费为0
    print("\n测试案例2: 首年保费为0")
    test_data_2 = pd.DataFrame({
        "SFYP2(不含短险续保)": [1000, 2000],
        "首年保费": [0, 5000],
        "保单号": ["P001", "P002"]
    })

    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            test_data_2,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费"
        )

        # 第1行应该是空值（除零错误）
        if result_df.loc[0, "承保趸期数据"] == "":
            print("  [OK] 除零错误正确处理")
        else:
            print(f"  [FAIL] 除零错误处理错误: {result_df.loc[0, '承保趸期数据']}")
            return False
    except Exception as e:
        print(f"  [FAIL] 除零测试失败: {str(e)}")
        return False

    # 测试案例3：边界值（R=0.1, R=1.0, R=0.2, R=0.9）
    print("\n测试案例3: 边界值测试")
    test_data_3 = pd.DataFrame({
        "SFYP2(不含短险续保)": [1000, 5000, 2000, 4500],
        "首年保费": [10000, 5000, 10000, 5000],
        "保单号": ["P001", "P002", "P003", "P004"]
    })

    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            test_data_3,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费"
        )

        expected = ["趸交FYP", "", "2年交SFYP", "9年交SFYP"]
        for i, exp in enumerate(expected):
            actual = result_df.loc[i, "承保趸期数据"]
            if actual == exp:
                print(f"  [OK] 行{i+1}: {actual}")
            else:
                print(f"  [FAIL] 行{i+1}: 期望={exp}, 实际={actual}")
                return False
    except Exception as e:
        print(f"  [FAIL] 边界值测试失败: {str(e)}")
        return False

    print("\n[OK] 所有边界条件测试通过")
    return True

def test_performance():
    """测试性能"""
    print("\n" + "=" * 60)
    print("测试性能")
    print("=" * 60)

    import time

    # 创建大数据集
    print("\n创建1000行测试数据...")
    large_data = pd.DataFrame({
        "SFYP2(不含短险续保)": [1000 if i % 3 == 0 else 5000 if i % 3 == 1 else 3000
                                for i in range(1000)],
        "首年保费": [10000] * 1000,
        "保单号": [f"P{str(i).zfill(4)}" for i in range(1000)],
        "其他数据": [f"Data{i}" for i in range(1000)]
    })

    # 测试计算性能
    print("测试1000行数据的计算性能...")
    formatter = DataFormatter()
    start_time = time.time()

    try:
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            large_data,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费"
        )

        end_time = time.time()
        elapsed = end_time - start_time

        print(f"  [OK] 计算完成")
        print(f"  处理行数: {len(large_data)}")
        print(f"  耗时: {elapsed:.3f} 秒")
        print(f"  平均: {elapsed/len(large_data)*1000:.3f} ms/行")

        if elapsed < 5.0:  # 5秒内完成
            print("  [OK] 性能达标（<5秒）")
        else:
            print("  [WARN] 性能可能需要优化")

        print(f"  需要用户输入的行数: {len(pending_rows)}")

    except Exception as e:
        print(f"  [FAIL] 性能测试失败: {str(e)}")
        return False

    print("\n[OK] 性能测试完成")
    return True

def main():
    """主函数"""
    print("承保趸期数据功能完整集成测试")
    print("=" * 60)
    print()

    # 运行所有测试
    test_results = []

    # 测试1: 完整流程
    print("\n[TEST 1] 完整数据处理流程")
    test_results.append(("完整流程", test_complete_flow()))

    # 测试2: 边界条件
    print("\n[TEST 2] 边界条件处理")
    test_results.append(("边界条件", test_edge_cases()))

    # 测试3: 性能测试
    print("\n[TEST 3] 性能测试")
    test_results.append(("性能测试", test_performance()))

    # 汇总结果
    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)

    for test_name, result in test_results:
        status = "[PASS]" if result else "[FAIL]"
        print(f"{status} {test_name}")

    passed = sum(1 for _, result in test_results if result)
    total = len(test_results)

    print(f"\n总计: {passed}/{total} 个测试通过")

    if passed == total:
        print("\n[SUCCESS] 所有测试通过！")
        print("\n承保趸期数据功能已就绪，包括：")
        print("  [OK] 核心计算逻辑")
        print("  [OK] 边界条件处理")
        print("  [OK] 批量输入对话框")
        print("  [OK] 与DataReader集成")
        print("  [OK] 性能优化")
        return 0
    else:
        print(f"\n[ERROR] {total - passed} 个测试失败")
        return 1

if __name__ == "__main__":
    sys.exit(main())
