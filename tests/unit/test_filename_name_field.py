"""
验证文件名生成逻辑：智能查找含"姓名"的列

用真实数据文件测试，确认：
  1. 递交文件（有"营销员或推广员 姓名"列）→ 文件名含真实姓名
  2. 承保文件（无姓名列）→ 文件名 fallback 到 output
  3. NaN/空值 → fallback 到 output
"""
import sys
import os
import pandas as pd

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))


def generate_name_for_row(row):
    """复刻 ppt_generator.py 里的文件名逻辑，独立测试"""
    name_value = 'output'
    if '姓名' in row:
        name_value = row['姓名']
    else:
        for col in row.index:
            if isinstance(col, str) and '姓名' in col:
                name_value = row[col]
                break
    if name_value is None or (isinstance(name_value, float) and pd.isna(name_value)) or str(name_value).strip() == '':
        name_value = 'output'
    else:
        name_value = str(name_value).strip()
    return name_value


def test_excel_file(filename):
    """测试一个 Excel 文件的文件名生成"""
    if not os.path.exists(filename):
        print(f"[跳过] {filename} 不存在")
        return None

    df = pd.read_excel(filename, nrows=3)
    name_cols = [c for c in df.columns if isinstance(c, str) and '姓名' in c]
    print(f"\n=== {filename} ===")
    print(f"  含\"姓名\"的列: {name_cols if name_cols else '（无）'}")

    all_pass = True
    for i, row in df.iterrows():
        name = generate_name_for_row(row)
        expected_is_output = len(name_cols) == 0
        actual_is_output = (name == 'output')
        ok = (expected_is_output == actual_is_output)
        status = 'PASS' if ok else 'FAIL'
        print(f"  [{status}] 行 {i}: 文件名 = {i+1}_{name}")
        if not ok:
            all_pass = False
    return all_pass


def test_nan_and_empty():
    """测试 NaN 和空值的 fallback"""
    print("\n=== NaN/空值测试 ===")
    df = pd.DataFrame({
        '营销员 姓名': ['正常', None, float('nan'), '  ', '张三'],
    })
    expected = ['正常', 'output', 'output', 'output', '张三']
    all_pass = True
    for i, row in df.iterrows():
        name = generate_name_for_row(row)
        ok = (name == expected[i])
        status = 'PASS' if ok else 'FAIL'
        print(f"  [{status}] 输入 {repr(df['营销员 姓名'][i])!s:20s} → {name} (期望 {expected[i]})")
        if not ok:
            all_pass = False
    return all_pass


def main():
    print("=" * 64)
    print("验证文件名姓名字段智能查找")
    print("=" * 64)

    results = []

    # 测试真实数据文件
    root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    r = test_excel_file(os.path.join(root, 'KA递交模板.xlsx'))
    if r is not None:
        results.append(r)

    r = test_excel_file(os.path.join(root, 'KA承保快报+-+2025-10-23T164946.667.xlsx'))
    if r is not None:
        results.append(r)

    # 测试边界情况
    results.append(test_nan_and_empty())

    print()
    print("=" * 64)
    passed = sum(results)
    total = len(results)
    print(f"结果: {passed}/{total} 通过")
    print("=" * 64)
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()
