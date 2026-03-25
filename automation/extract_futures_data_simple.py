import pandas as pd
import json
import os
from datetime import datetime

# 简单测试
print("Testing...")

try:
    # 读取Excel
    df = pd.read_excel('data_source/futures_data.xlsx', sheet_name=0, header=None)
    print(f"Data shape: {df.shape}")
    
    # 打印前几行数据
    for i in range(3):
        print(f"Row {i}: {df.iloc[i].tolist()}")
        
    print("Test successful!")
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()