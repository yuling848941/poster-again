#!/usr/bin/env python3
"""
批量替换 main_window.py 中的 self.log_text.append 为 print + logger
"""

import re

file_path = "src/gui/main_window.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 正则表达式匹配 self.log_text.append("...")
pattern = r'self\.log_text\.append\((f?".*?")\)'

def replace_match(match):
    arg = match.group(1)
    # 如果是 f-string，保留 f
    return f'print(f"[LOG] {{{arg.lstrip("f")}}}"); logger.info({arg})'

# 替换所有匹配
new_content = re.sub(pattern, replace_match, content)

# 写入文件
with open(file_path, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("替换完成！")

# 检查是否还有剩余的 log_text 引用
remaining = re.findall(r'self\.log_text', new_content)
if remaining:
    print(f"警告：还有 {len(remaining)} 处 log_text 引用")
else:
    print("所有 log_text 引用已清除")
