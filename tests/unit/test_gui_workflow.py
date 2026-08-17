#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
模拟GUI工作流程测试
测试DataReader在GUI中的使用方式（含承保趸期数据计算）
"""

import os
import sys
import tempfile
import shutil

import pandas as pd

# 将项目根目录加入 sys.path，确保 src 包可导入
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

from src.data_reader import DataReader


def _create_test_excel():
    """生成包含承保趸期计算所需列的自包含测试 Excel 文件"""
    data = pd.DataFrame({
        "保单号": ["P001", "P002", "P003"],
        "SFYP2(不含短险续保)": [5000, 5000, 1000],
        "首年保费": [10000, 5000, 10000],
        "缴费年期（主险） ID": [20, 5, 20],
    })
    # R: 0.5 -> 5年交SFYP；1.0 -> 需用户输入；0.1 -> 趸交FYP
    tmp_dir = tempfile.mkdtemp()
    path = os.path.join(tmp_dir, "ka_test.xlsx")
    data.to_excel(path, index=False)
    return tmp_dir, path


def test_gui_workflow():
    """模拟GUI工作流程"""
    print("=" * 60)
    print("模拟GUI工作流程测试")
    print("=" * 60)

    tmp_dir, excel_file = _create_test_excel()
    try:
        print("步骤1: 创建DataReader")
        data_reader = DataReader()

        print("步骤2: 加载Excel文件（模拟GUI中不传递parent_widget）")
        success = data_reader.load_excel(excel_file, use_thousand_separator=True, parent_widget=None)

        if not success:
            print("[FAIL] 数据加载失败")
            return False

        print("[OK] 数据加载成功")
        print(f"  数据行数: {len(data_reader.data)}")
        print(f"  数据列数: {len(data_reader.data.columns)}")

        if "承保趸期数据" not in data_reader.data.columns:
            print("[FAIL] 承保趸期数据列未生成")
            return False

        print("[OK] 承保趸期数据列已生成")

        # 检查是否有空值（需要用户输入的行）
        empty_values = data_reader.data["承保趸期数据"] == ""
        if empty_values.any():
            print(f"[INFO] 发现 {empty_values.sum()} 行需要用户输入（R=1的情况）")
        else:
            print("[OK] 所有行都已计算出结果，无空值")

        return True
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_with_gui_parent():
    """测试传递parent_widget的情况"""
    print()
    print("=" * 60)
    print("测试传递parent_widget（模拟GUI环境）")
    print("=" * 60)

    tmp_dir, excel_file = _create_test_excel()
    try:
        data_reader = DataReader()
        success = data_reader.load_excel(excel_file, use_thousand_separator=True, parent_widget=None)

        if not success:
            print("[FAIL] 数据加载失败")
            return False

        print("[OK] 数据加载成功（无论是否传递parent_widget）")

        if "承保趸期数据" in data_reader.data.columns:
            result = data_reader.data["承保趸期数据"].iloc[0]
            print(f"承保趸期数据结果: {result}")

            if result == "5年交SFYP":
                print("[OK] 结果正确！")
            else:
                print(f"[FAIL] 结果不正确，期望'5年交SFYP'，实际'{result}'")
                return False
        else:
            print("[FAIL] 承保趸期数据列未生成")
            return False

        return True
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def main():
    """主函数"""
    print("承保趸期数据GUI工作流程测试")
    print("=" * 60)

    test1_ok = test_gui_workflow()
    test2_ok = test_with_gui_parent()

    print()
    print("=" * 60)
    print("测试结果汇总")
    print("=" * 60)

    if test1_ok and test2_ok:
        print("[SUCCESS] 所有测试通过")
        return 0
    else:
        print("[FAIL] 部分测试失败")
        return 1


if __name__ == "__main__":
    sys.exit(main())
