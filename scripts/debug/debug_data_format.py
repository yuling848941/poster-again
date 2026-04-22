#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
数据格式调试脚本
检查Excel数据格式和计算问题
"""

import pandas as pd
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def debug_excel_data(excel_file):
    """调试Excel数据格式"""
    print("=" * 60)
    print("Excel数据格式调试")
    print("=" * 60)

    # 读取Excel文件
    try:
        df = pd.read_excel(excel_file)
        print(f"\n[OK] 成功读取Excel文件: {excel_file}")
        print(f"数据行数: {len(df)}")
        print(f"数据列数: {len(df.columns)}")

        # 显示列名
        print("\n列名列表:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1}. '{col}'")

        # 检查是否有SFYP2和首年保费列
        sfyp2_col = "SFYP2(不含短险续保)"
        premium_col = "首年保费"

        print(f"\n检查必要列:")
        has_sfyp2 = sfyp2_col in df.columns
        has_premium = premium_col in df.columns
        print(f"  SFYP2列存在: {has_sfyp2}")
        print(f"  首年保费列存在: {has_premium}")

        if has_sfyp2 and has_premium:
            # 显示数据类型
            print(f"\n数据类型:")
            print(f"  SFYP2列类型: {df[sfyp2_col].dtype}")
            print(f"  首年保费列类型: {df[premium_col].dtype}")

            # 显示前5行数据
            print(f"\n前5行数据:")
            print(df[[sfyp2_col, premium_col]].head())

            # 检查数值类型
            print(f"\n数值类型检查:")

            # 检查SFYP2
            sfyp2_values = df[sfyp2_col]
            print(f"  SFYP2列:")
            print(f"    非空值数量: {sfyp2_values.count()}/{len(sfyp2_values)}")
            print(f"    是否全为数值: {pd.api.types.is_numeric_dtype(sfyp2_values)}")

            # 转换为数值类型（如果需要）
            try:
                sfyp2_numeric = pd.to_numeric(sfyp2_values, errors='coerce')
                print(f"    转换后非空值: {sfyp2_numeric.count()}")
                print(f"    转换后无效值: {sfyp2_numeric.isna().sum()}")
            except:
                print(f"    转换失败")

            # 检查首年保费
            premium_values = df[premium_col]
            print(f"  首年保费列:")
            print(f"    非空值数量: {premium_values.count()}/{len(premium_values)}")
            print(f"    是否全为数值: {pd.api.types.is_numeric_dtype(premium_values)}")

            # 转换为数值类型（如果需要）
            try:
                premium_numeric = pd.to_numeric(premium_values, errors='coerce')
                print(f"    转换后非空值: {premium_numeric.count()}")
                print(f"    转换后无效值: {premium_numeric.isna().sum()}")
            except:
                print(f"    转换失败")

            # 计算比例R = SFYP2 / 首年保费
            print(f"\n比例计算:")
            try:
                # 转换为数值类型
                sfyp2_num = pd.to_numeric(df[sfyp2_col], errors='coerce')
                premium_num = pd.to_numeric(df[premium_col], errors='coerce')

                # 计算比例
                ratio = sfyp2_num / premium_num

                # 显示前5行的比例
                print(f"前5行的比例计算:")
                for i in range(min(5, len(df))):
                    sfyp2 = sfyp2_num.iloc[i]
                    premium = premium_num.iloc[i]
                    r = ratio.iloc[i]

                    print(f"  行{i+1}: SFYP2={sfyp2}, 首年保费={premium}, R={r:.3f}")

                # 统计比例分布
                print(f"\n比例分布统计:")
                print(f"  总行数: {len(ratio)}")
                print(f"  有效比例: {ratio.notna().sum()}")
                print(f"  无效比例: {ratio.isna().sum()}")

                # 查找R≈0.5的行
                print(f"\n查找R≈0.5的行:")
                near_half = abs(ratio - 0.5) < 0.01
                if near_half.any():
                    print(f"  找到 {near_half.sum()} 行R≈0.5:")
                    for i in ratio[near_half].index:
                        print(f"    行{i+1}: R={ratio[i]:.3f}, SFYP2={sfyp2_num[i]}, 首年保费={premium_num[i]}")
                else:
                    print(f"  未找到R≈0.5的行")
                    # 显示最接近0.5的值
                    if ratio.notna().any():
                        closest_idx = (ratio - 0.5).abs().idxmin()
                        print(f"  最接近的行: 行{closest_idx+1}, R={ratio[closest_idx]:.3f}")

            except Exception as e:
                print(f"  计算失败: {str(e)}")
                import traceback
                traceback.print_exc()

        else:
            print(f"\n[FAIL] 缺少必要列，无法继续调试")
            return False

        return True

    except Exception as e:
        print(f"\n[FAIL] 读取Excel文件失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def test_formatter_with_real_data(excel_file):
    """使用真实数据测试DataFormatter"""
    print("\n" + "=" * 60)
    print("使用DataFormatter测试")
    print("=" * 60)

    from src.memory_management.data_formatter import DataFormatter

    try:
        # 读取数据
        df = pd.read_excel(excel_file)

        # 创建formatter
        formatter = DataFormatter()

        # 计算承保趸期数据
        result_df, pending_rows = formatter.calculate_chengbao_term_data(
            df,
            sfyp2_column="SFYP2(不含短险续保)",
            premium_column="首年保费",
            target_column="承保趸期数据",
            policy_number_column="保单号"
        )

        print(f"\n计算结果:")
        print(f"  处理行数: {len(result_df)}")
        print(f"  需要用户输入的行数: {len(pending_rows)}")

        if "承保趸期数据" in result_df.columns:
            # 显示前10行的结果
            print(f"\n前10行结果:")
            cols_to_show = ["SFYP2(不含短险续保)", "首年保费", "承保趸期数据"]
            if "保单号" in result_df.columns:
                cols_to_show.insert(0, "保单号")

            print(result_df[cols_to_show].head(10).to_string(index=True))

            # 统计结果
            result_counts = result_df["承保趸期数据"].value_counts(dropna=False)
            print(f"\n结果统计:")
            for value, count in result_counts.items():
                if pd.isna(value) or value == "":
                    print(f"  空值: {count} 行")
                else:
                    print(f"  {value}: {count} 行")
        else:
            print(f"\n[FAIL] 未生成承保趸期数据列")

    except Exception as e:
        print(f"\n[FAIL] 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()

def main():
    """主函数"""
    # 获取Excel文件路径（从命令行参数或默认路径）
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        # 尝试查找用户提到的文件
        excel_file = "承保快报.xlsx"

    if not os.path.exists(excel_file):
        print(f"错误：找不到Excel文件 '{excel_file}'")
        print(f"\n请将Excel文件放在当前目录，或使用完整路径：")
        print(f"  python debug_data_format.py <Excel文件路径>")
        print(f"\n当前目录文件:")
        for f in os.listdir('.'):
            if f.endswith(('.xlsx', '.xls')):
                print(f"  - {f}")
        return 1

    # 调试数据格式
    debug_excel_data(excel_file)

    # 测试DataFormatter
    test_formatter_with_real_data(excel_file)

    return 0

if __name__ == "__main__":
    sys.exit(main())
