import pandas as pd
import os

# 读取台账文件
file_path = r'D:\申银万国期货\3.档案管理\2、业务档案\3. 台账相关-每年随档案报送\期货交易咨询跟踪台账-魏子骥-20251218.xlsx'

try:
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # 查找"毛喜瑞"相关记录
    print("=== 查找客户'毛喜瑞' ===")
    
    # 查找包含"毛喜瑞"的客户名称
    mask = df['客户名称'].astype(str).str.contains('毛喜瑞', na=False)
    if mask.any():
        mao_client = df[mask]
        print(f"找到 {len(mao_client)} 条匹配记录:")
        print(mao_client[['客户名称', '产品名称', '正确代码', '资金代码', '承揽业务单位', '业务人员', '项目负责人']].to_string(index=False))
    else:
        print("未找到'毛喜瑞'，尝试查找类似客户...")
        
        # 显示所有客户名称供参考
        print("\n所有客户名称:")
        clients = df['客户名称'].dropna().unique()
        for i, client in enumerate(clients[:50], 1):
            print(f"{i:3}. {client}")
        
        # 搜索部分匹配
        print("\n=== 搜索部分匹配 ===")
        for client in clients:
            if '毛' in str(client) or '喜' in str(client) or '瑞' in str(client):
                print(f"可能匹配: {client}")
    
    # 检查数据完整性
    print("\n=== 数据完整性检查 ===")
    print(f"总记录数: {len(df)}")
    print(f"客户名称非空: {df['客户名称'].notna().sum()}")
    print(f"产品名称非空: {df['产品名称'].notna().sum()}")
    
except Exception as e:
    print(f"错误: {e}")