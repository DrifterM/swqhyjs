import pandas as pd
import json
import os
from datetime import datetime

def extract_futures_data():
    # 读取Excel文件
    df = pd.read_excel('data_source/futures_data.xlsx', sheet_name=0, header=None)
    
    # 获取标题行（第1行）
    headers = df.iloc[1].tolist()
    
    # 找到各列索引
    col_indices = {}
    for i, header in enumerate(headers):
        if isinstance(header, str):  # 只处理字符串类型的标题
            if '前日收盘价' in header:
                col_indices['prev_close'] = i
            elif '昨日收盘价' in header:
                col_indices['yest_close'] = i
            elif '今日收盘价' in header:
                col_indices['today_close'] = i
            elif '昨日涨跌幅' in header:
                col_indices['yest_change'] = i
            elif '今日涨跌幅' in header:
                col_indices['today_change'] = i
            elif '昨日成交量' in header:
                col_indices['yest_volume'] = i
            elif '今日成交量' in header:
                col_indices['today_volume'] = i
    
    # 判断是否收盘
    now = datetime.now()
    market_closed = (now.hour > 15) or (now.hour == 15 and now.minute >= 30)
    
    # 选择数据列
    if market_closed:
        change_col = col_indices.get('today_change')
        volume_col = col_indices.get('today_volume')
        date_used = '今日'
    else:
        change_col = col_indices.get('yest_change')
        volume_col = col_indices.get('yest_volume')
        date_used = '昨日'
    
    # 提取品种数据
    commodities = []
    for idx in range(2, len(df)):  # 从第2行开始是数据
        row = df.iloc[idx]
        symbol = row[1]  # 符号在第1列
        name = row[2]   # 名称在第2列
        
        if pd.isna(symbol) or pd.isna(name):
            continue
        
        # 处理名称和符号
        symbol_str = str(symbol).strip()
        if '.' in symbol_str:
            symbol_str = symbol_str.split('.')[0]
        
        name_str = str(name).strip()
        # 去除合约号（数字）
        name_str = ''.join([c for c in name_str if not c.isdigit()])
        
        # 提取数据
        try:
            change_val = float(row[change_col]) * 100 if not pd.isna(row[change_col]) else 0.0
            volume_val = int(row[volume_col]) / 10000 if not pd.isna(row[volume_col]) else 0
        except Exception as e:
            print(f"Error extracting data: {e}")
            change_val = 0.0
            volume_val = 0
        
        # 分类
        category = classify_commodity(name_str)
        
        commodities.append({
            "name": name_str,
            "symbol": symbol_str,
            "change": round(change_val, 2),
            "volume": round(volume_val, 1),
            "category": category
        })
    
    # 创建输出数据
    output_data = {
        "last_updated": datetime.now().isoformat(),
        "date_used": date_used,
        "market_closed": market_closed,
        "commodities": commodities
    }
    
    # 保存到JSON文件
    output_dir = "prototype_v2_expanded"
    os.makedirs(output_dir, exist_ok=True)
    
    output_path = os.path.join(output_dir, 'data.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    
    print("数据提取完成！")
    print(f"共处理 {len(commodities)} 个品种")
    print(f"使用 {date_used} 数据")
    print(f"已保存到: {output_path}")
    
    return output_data

def classify_commodity(name):
    """分类品种"""
    black_list = ['螺纹钢', '铁矿石', '焦炭', '热卷', '硅铁', '锰硅', '玻璃', '纯碱']
    nonferrous_list = ['铜', '铝', '锌', '铅', '镍', '锡', '氧化铝', '工业硅']
    energy_list = ['原油', '燃油', '低硫燃料油', '沥青', 'PTA', '甲醇', '塑料', 'PP', 'LPG', '乙二醇', '苯乙烯']
    precious_list = ['黄金', '白银']
    agricultural_list = ['豆一', '豆二', '豆粕', '豆油', '棕榈油', '菜粕', '菜油', '花生', '玉米', '淀粉', '鸡蛋', '生猪', '棉花', '棉纱']
    financial_list = ['股指', '国债', '沪深300', '上证50', '中证500', '中证1000', '10年期国债']
    
    if any(keyword in name for keyword in black_list):
        return '黑色系'
    elif any(keyword in name for keyword in nonferrous_list):
        return '有色金属'
    elif any(keyword in name for keyword in energy_list):
        return '能源化工'
    elif any(keyword in name for keyword in precious_list):
        return '贵金属'
    elif any(keyword in name for keyword in agricultural_list):
        return '农产品'
    elif any(keyword in name for keyword in financial_list):
        return '金融指数'
    else:
        return '其他'

if __name__ == "__main__":
    extract_futures_data()