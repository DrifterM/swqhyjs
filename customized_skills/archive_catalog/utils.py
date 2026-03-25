import os
import pandas as pd
from pathlib import Path

def find_latest_ledger(base_path_pattern):
    """
    根据路径模式（含*）查找最新的台账文件
    """
    path = Path(base_path_pattern.replace('*', '[0-9]*'))
    files = list(path.parent.glob(path.name))
    if not files:
        return None
    # 按文件名排序，取最新
    latest = max(files, key=lambda f: f.name)
    return str(latest)

def get_client_code_from_ledger(ledger_path, client_name):
    """
    从 Excel 台账中提取客户代码
    优先使用“正确代码”，其次“资金代码”
    """
    try:
        df = pd.read_excel(ledger_path, dtype=str)
        # 查找客户行
        matched = df[df['客户名称'].str.contains(client_name, na=False)]
        if len(matched) == 0:
            return None
        row = matched.iloc[0]
        # 优先“正确代码”，其次“资金代码”
        code = row.get('正确代码') or row.get('资金代码')
        return code.strip() if code else None
    except Exception as e:
        print(f"Error reading ledger {ledger_path}: {e}")
        return None