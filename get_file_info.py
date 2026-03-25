import os
import pandas as pd
from datetime import datetime

project_path = r'D:\申银万国期货\3.档案管理\2、业务档案\2. 未归档项目-2025\毛喜瑞交易咨询项目'

print("=== 项目文件夹文件分析 ===")

# 获取所有文件
files = os.listdir(project_path)
print(f"总文件数: {len(files)}")

# 筛选阿拉伯数字开头的文件
digit_files = []
for file in files:
    if file and file[0].isdigit():  # 以数字开头
        # 检查是否是以年份开头（如2024、2025），按你要求忽略
        if len(file) >= 4 and file[:4].isdigit() and 2000 <= int(file[:4]) <= 2100:
            continue  # 跳过以年份开头的文件
        digit_files.append(file)

print(f"\n以阿拉伯数字开头的文件: {len(digit_files)}")
for file in digit_files:
    full_path = os.path.join(project_path, file)
    # 获取修改时间
    mod_time = os.path.getmtime(full_path)
    mod_date = datetime.fromtimestamp(mod_time).strftime('%Y%m%d')
    print(f"  {file} -> 修改日期: {mod_date}")

# 按数字排序
digit_files_sorted = sorted(digit_files, key=lambda x: int(x.split('、')[0]) if '、' in x else int(x.split('.')[0]))
print(f"\n排序后的文件列表 (按数字顺序):")
for i, file in enumerate(digit_files_sorted, 1):
    full_path = os.path.join(project_path, file)
    mod_time = os.path.getmtime(full_path)
    mod_date = datetime.fromtimestamp(mod_time).strftime('%Y%m%d')
    print(f"  {i:04d}. {file} -> {mod_date}")