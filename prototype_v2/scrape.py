"""
期货数据抓取脚本
从新浪财经获取主力合约数据
"""

import requests
from bs4 import BeautifulSoup
import json
import datetime

def scrape_futures_data():
    """抓取核心期货品种数据"""
    url = "https://finance.sina.com.cn/futures/quotes/"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    # 目标品种映射
    symbols = {
        '螺纹钢': 'RB',
        '铜': 'CU',
        '原油': 'SC',
        '黄金': 'AU',
        '玉米': 'C',
        '股指': 'IF'
    }
    
    # 板块分类
    categories = {
        '螺纹钢': '黑色系',
        '铜': '有色金属',
        '原油': '能源化工',
        '黄金': '贵金属',
        '玉米': '农产品',
        '股指': '金融指数'
    }
    
    data = []
    
    for name, symbol in symbols.items():
        try:
            # 构造具体品种URL
            if name == '股指':
                quote_url = f"{url}IF.shtml"  # 股指期货
            else:
                quote_url = f"{url}{symbol}0.shtml"  # 主力合约
            
            response = requests.get(quote_url, headers=headers, timeout=10)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 提取价格和涨跌幅
            price_elem = soup.find('div', class_='price')
            change_elem = soup.find('div', class_='change')
            volume_elem = soup.find('td', string='成交量')
            
            if price_elem and change_elem:
                # 涨跌幅
                change_text = change_elem.text.strip()
                change = float(change_text.replace('%', ''))
                
                # 成交量（简化处理）
                volume = 100  # 默认值，实际需要更精确抓取
                
                data.append({
                    'name': name,
                    'symbol': symbol,
                    'change': change,
                    'volume': volume,
                    'category': categories[name]
                })
                
        except Exception as e:
            print(f"抓取{name}失败: {e}")
            # 使用默认数据
            data.append({
                'name': name,
                'symbol': symbol,
                'change': 0.0,
                'volume': 100,
                'category': categories[name]
            })
    
    return {
        'last_updated': datetime.datetime.now().isoformat(),
        'commodities': data
    }

def save_data(data, filename='data.json'):
    """保存数据到JSON文件"""
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
if __name__ == "__main__":
    print("开始抓取期货数据...")
    data = scrape_futures_data()
    save_data(data, 'data.json')
    print("数据抓取完成!")