#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
创建测试用Excel文件
"""

import pandas as pd

# 创建测试数据
data = {
    "保单号": ["A001", "A002", "A003", "A004"],
    "SFYP2(不含短险续保)": [100, 200, 300, 400],
    "首年保费": [1000, 2000, 3000, 4000]
}

df = pd.DataFrame(data)

# 保存到Excel文件
excel_file = "test_chengbao.xlsx"
df.to_excel(excel_file, index=False)

print(f"已创建测试Excel文件: {excel_file}")
print("\n文件内容:")
print(df)
