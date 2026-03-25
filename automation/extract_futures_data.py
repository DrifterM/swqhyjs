"""
从futures_data.xlsx提取期货行情数据
根据当前时间决定使用昨日或今日数据
"""

import pandas as pd
import json
from datetime import datetime, timedelta
import os
import sys
import io

# 设置GBK编码环境
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gbk')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='gbk')

def extract_futures_data(excel_path: str, output_dir: str = "../prototype_v2_expanded"):
    """
    从Excel文件中提取期货数据并生成data.json
    
    参数:
        excel_path: Excel文件路径
        output_dir: 输出目录
    """
    # 读取Excel文件，不使用表头
    df = pd.read_excel(excel_path, sheet_name=0, header=None)
    
    # 获取列信息
    headers = df.iloc[1].tolist()  # 第二行是标题
    dates = df.iloc[0].tolist()     # 第一行是日期
    
    # 找到关键列的索引
    col_indices = {}
    for i, header in enumerate(headers):
        header_str = str(header)
        if isinstance(header, pd.Timestamp) or isinstance(header, pd.NaTType):
            continue
        if '前日收盘价' in header_str:
            col_indices['prev_close'] = i
        elif '昨日收盘价' in header_str:
            col_indices['yest_close'] = i
        elif '今日收盘价' in header_str:
            col_indices['today_close'] = i
        elif '昨日涨跌幅' in header_str:
            col_indices['yest_change'] = i
        elif '今日涨跌幅' in header_str:
            col_indices['today_change'] = i
        elif '昨日成交量' in header_str:
            col_indices['yest_volume'] = i
        elif '今日成交量' in header_str:
            col_indices['today_volume'] = i
    
    # 当前时间判断
    now = datetime.now()
    market_closed = (now.hour > 15) or (now.hour == 15 and now.minute >= 30)
    
    # 确定使用哪一列数据
    if market_closed:
        # 已收盘，使用今日数据
        change_col = col_indices.get('today_change')
        volume_col = col_indices.get('today_volume')
        date_used = '今日'
    else:
        # 未收盘，使用昨日数据
        change_col = col_indices.get('yest_change')
        volume_col = col_indices.get('yest_volume')
        date_used = '昨日'
    
    # 提取品种数据
    commodities = []
    for idx, row in df.iterrows():
        if idx < 2:  # 跳过前两行（标题行）
            continue
        
        symbol_cell = row[0]
        name_cell = row[2]
        
        # 跳过空行
        if pd.isna(symbol_cell) or pd.isna(name_cell):
            continue
        
        # 提取符号和名称
        symbol = str(symbol_cell).strip()
        if '.' in symbol:
            symbol = symbol.split('.')[0]
        
        name = str(name_cell).strip()
        if '260' in name or '2' in name:
            # 去除合约号
            name = ''.join([c for c in name if not c.isdigit()])
        
        # 提取涨跌幅和成交量
        try:
            change_val = float(row[change_col]) * 100 if not pd.isna(row[change_col]) else 0.0
            volume_val = int(row[volume_col]) / 10000 if not pd.isna(row[volume_col]) else 0  # 万手
        except Exception as e:
            change_val = 0.0
            volume_val = 0
        
        # 品种分类
        category = classify_commodity(name)
        
        commodities.append({
            "name": name,
            "symbol": symbol,
            "change": round(change_val, 2),
            "volume": round(volume_val, 1),
            "category": category
        })
    
    # 创建输出数据结构
    output_data = {
        "last_updated": datetime.now().isoformat(),
        "date_used": date_used,
        "market_closed": market_closed,
        "commodities": commodities
    }
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 写入JSON文件
    output_path = os.path.join(output_dir, 'data.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    
    print("✅ 数据提取完成！")
    print(f"📊 共处理 {len(commodities)} 个品种")
    print(f"📅 使用 {date_used} 数据")
    print(f"⏰ 市场状态：{'已收盘' if market_closed else '未收盘'}")
    print(f"💾 已保存到: {output_path}")
    
    return output_data

def classify_commodity(name: str) -> str:
    """根据品种名称分类"""
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
    # 设置文件路径
    excel_file = r'C:\Users\H3C\.copaw\data_source\futures_data.xlsx'
    
    # 执行数据提取
    try:
        result = extract_futures_data(excel_file)
        print("✅ 执行成功！")
    except Exception as e:
        import traceback
        print("❌ 执行失败:")
        traceback.print_exc()
        sys.exit(1)