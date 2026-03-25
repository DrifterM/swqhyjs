import pandas as pd
import os

# 文件路径 - 即使控制台显示乱码，Python也能正确识别
file_path = r'D:\申银万国期货\3.档案管理\2、业务档案\3. 台账相关-每年随档案报送\期货交易咨询跟踪台账-魏子骥-20251218.xlsx'

print(f"尝试读取文件: {file_path}")
print(f"文件是否存在: {os.path.exists(file_path)}")

try:
    # 读取Excel文件
    df = pd.read_excel(file_path, engine='openpyxl')
    
    print("\n=== 成功读取台账文件 ===")
    print(f"文件形状: {df.shape} (行数: {df.shape[0]}, 列数: {df.shape[1]})")
    print("\n列名列表:")
    for i, col in enumerate(df.columns, 1):
        print(f"{i}. {col}")
    
    print("\n=== 前5行数据 ===")
    print(df.head())
    
    # 查找需要的列
    print("\n=== 关键字段检查 ===")
    required_cols = ['正确代码', '资金代码', '客户名称', '产品名称', '承揽业务单位', '业务人员', '项目负责人']
    
    for col in required_cols:
        if col in df.columns:
            print(f"✓ '{col}' 存在")
        else:
            # 尝试相似匹配
            found = False
            for actual_col in df.columns:
                if col in str(actual_col):
                    print(f"→ '{col}' 可能对应 '{actual_col}'")
                    found = True
                    break
            if not found:
                print(f"✗ '{col}' 未找到")
    
    # 数据概况
    print("\n=== 数据概况 ===")
    if '客户名称' in df.columns:
        print(f"客户数量: {df['客户名称'].nunique()}")
        print("\n前10个客户:")
        print(df['客户名称'].dropna().head(10).tolist())
    
    if '正确代码' in df.columns and '资金代码' in df.columns:
        print(f"\n正确代码数量: {df['正确代码'].notna().sum()}")
        print(f"资金代码数量: {df['资金代码'].notna().sum()}")
        print(f"两者都有值的数量: {df[df['正确代码'].notna() & df['资金代码'].notna()].shape[0]}")
        print(f"两者都为空的数量: {df[df['正确代码'].isna() & df['资金代码'].isna()].shape[0]}")
    
except Exception as e:
    print(f"\n❌ 读取文件时出错: {type(e).__name__}")
    print(f"错误详情: {e}")
    print("\n尝试修复路径...")
    
    # 尝试通配符查找
    import glob
    folder = r'D:\申银万国期货\3.档案管理\2、业务档案\3. 台账相关-每年随档案报送'
    xlsx_files = glob.glob(os.path.join(folder, "*.xlsx"))
    print(f"\n文件夹中的xlsx文件:")
    for f in xlsx_files:
        print(f"  - {f}")