"""
创建符合正式规范的卷内文件目录Word文档
格式要求：
1. 第一行：黑体三号居左"附件2"
2. 第二行：黑体小二居中"期货交易咨询业务档案卷内文件目录"
3. 空一行
4. 单个统一表格，结构如下：
   行1: 档号
   行2: 案卷题名  
   行3: 客户名称 + 客户代码
   行4-11: 文件清单表头及内容
5. 表格内文字：宋体小四
6. 文件题名列加宽
"""

import os
from datetime import datetime

# 项目信息
PROJECT_NAME = "毛喜瑞交易咨询项目"
CLIENT_CODE = "8005108018"
PRODUCT_NAME = "毛喜瑞交易咨询项目"
YEAR = "2025"
PROJECT_NUM = "0001"
ARCHIVE_CODE = f"SWQH-YW·{YEAR}-D30-3-1009-{PROJECT_NUM}"

# 文件列表
file_list = [
    {"序号": "1", "文件题名": "1、申报OA-毛喜瑞.pdf", "日期": "20251110"},
    {"序号": "2", "文件题名": "2、业务（含产品、服务）洗钱风险评估表-毛喜瑞交易咨询项目.xlsx", "日期": "20251110"},
    {"序号": "3", "文件题名": "3、合规审核意见截图.png", "日期": "20251111"},
    {"序号": "4", "文件题名": "4、交易咨询部-魏子骥-期货交易咨询业务风险等级评估表.pdf", "日期": "20251114"},
    {"序号": "5", "文件题名": "5、毛喜瑞伟交易咨询项目评审表决票.pdf", "日期": "20251110"},
    {"序号": "6", "文件题名": "6、期货交易咨询业务内部评审小组立项讨论会纪要-毛喜瑞交易咨询项目.pdf", "日期": "20251114"},
    {"序号": "7", "文件题名": "7、用印OA-毛喜瑞与申银万国期货有限公司之期货交易咨询协议.pdf", "日期": "20260127"},
    {"序号": "8", "文件题名": "8、毛喜瑞与申银万国期货有限公司之期货交易咨询协议.pdf", "日期": "20251231"},
]

# 清理之前的测试文件
test_files = [
    "卷内文件目录_预览.html",
    "卷内文件目录_文本版.txt", 
    "卷内文件目录_数据.csv",
    "卷内文件目录_标准版.txt"
]

project_path = r'D:\申银万国期货\3.档案管理\2、业务档案\2. 未归档项目-2025\毛喜瑞交易咨询项目'
for file in test_files:
    try:
        os.remove(os.path.join(project_path, file))
        print(f"已清理: {file}")
    except:
        pass

# 创建符合规范的HTML预览（便于查看格式）
html_content = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>卷内文件目录 - 规范版预览</title>
    <style>
        @page {
            size: A4;
            margin: 2cm;
        }
        body {
            font-family: 'SimSun', '宋体', sans-serif;
            margin: 0;
            padding: 0;
        }
        .container {
            width: 21cm;
            margin: 0 auto;
            padding: 1cm;
        }
        .attachment {
            font-family: 'SimHei', '黑体', sans-serif;
            font-size: 16pt;
            font-weight: bold;
            text-align: left;
            margin-bottom: 10px;
        }
        .main-title {
            font-family: 'SimHei', '黑体', sans-serif;
            font-size: 18pt;
            font-weight: bold;
            text-align: center;
            margin-bottom: 30px;
        }
        .spacer {
            height: 20px;
        }
        .catalog-table {
            width: 100%;
            border-collapse: collapse;
            border: 2px solid #000;
            font-family: 'SimSun', '宋体', sans-serif;
            font-size: 12pt;
        }
        .catalog-table th, .catalog-table td {
            border: 1px solid #000;
            padding: 8px 5px;
            text-align: center;
            vertical-align: middle;
        }
        .catalog-table th {
            background-color: #f0f0f0;
            font-weight: bold;
        }
        .file-title-cell {
            text-align: left;
            min-width: 400px;
        }
        .archive-code-row {
            font-weight: bold;
        }
        .product-title-row {
            font-weight: bold;
        }
        .client-info-row td {
            padding: 10px 5px;
        }
        .table-header-row {
            background-color: #e0e0e0;
        }
        .generated-time {
            text-align: right;
            margin-top: 30px;
            font-size: 10pt;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- 第一行：附件2 -->
        <div class="attachment">附件2</div>
        
        <!-- 第二行：主标题 -->
        <div class="main-title">期货交易咨询业务档案卷内文件目录</div>
        
        <!-- 空一行 -->
        <div class="spacer"></div>
        
        <!-- 单个统一表格 -->
        <table class="catalog-table">
            <!-- 第一行：档号 -->
            <tr class="archive-code-row">
                <td colspan="8">''' + ARCHIVE_CODE + '''</td>
            </tr>
            
            <!-- 第二行：案卷题名 -->
            <tr class="product-title-row">
                <td colspan="8">''' + PRODUCT_NAME + '''</td>
            </tr>
            
            <!-- 第三行：客户名称 + 客户代码 -->
            <tr class="client-info-row">
                <td colspan="4" style="text-align: left; padding-left: 20px;">客户名称：''' + PROJECT_NAME + '''</td>
                <td colspan="4" style="text-align: left; padding-left: 20px;">客户代码：''' + CLIENT_CODE + '''</td>
            </tr>
            
            <!-- 第四行：表头 -->
            <tr class="table-header-row">
                <th width="60">序号</th>
                <th width="450">文件题名</th>
                <th width="80">责任者</th>
                <th width="80">经办人</th>
                <th width="100">日期</th>
                <th width="100">材料性质</th>
                <th width="100">起止页号</th>
                <th width="150">备注</th>
            </tr>
'''

# 添加文件行
for file_item in file_list:
    html_content += f'''
            <tr>
                <td>{file_item['序号']}</td>
                <td class="file-title-cell">{file_item['文件题名']}</td>
                <td></td>
                <td></td>
                <td>{file_item['日期']}</td>
                <td></td>
                <td></td>
                <td></td>
            </tr>'''

html_content += f'''
        </table>
        
        <div class="generated-time">
            生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        </div>
    </div>
</body>
</html>'''

# 保存HTML预览
html_path = os.path.join(project_path, "卷内文件目录_规范预览.html")
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"✅ 规范预览文件已生成: {html_path}")
print("\n=== 文档结构预览 ===")
print("1. 附件2 (黑体三号，居左)")
print("2. 期货交易咨询业务档案卷内文件目录 (黑体小二，居中)")
print("3. 空一行")
print("4. 单个统一表格，包含：")
print("   - 第1行：档号 (SWQH-YW·2025-D30-3-1009-0001)")
print("   - 第2行：案卷题名 (毛喜瑞交易咨询项目)")
print("   - 第3行：客户名称 + 客户代码")
print("   - 第4行：表头 (序号、文件题名、责任者、经办人、日期、材料性质、起止页号、备注)")
print("   - 第5-12行：8个文件清单")
print("\n5. 表格内文字：宋体小四")
print("6. 文件题名列已加宽")
print(f"\n档号格式: {ARCHIVE_CODE}")

# 同时创建纯文本版便于快速查看
text_content = f"""附件2

期货交易咨询业务档案卷内文件目录

档号：{ARCHIVE_CODE}
案卷题名：{PRODUCT_NAME}
客户名称：{PROJECT_NAME}    客户代码：{CLIENT_CODE}

序号  文件题名                                                  责任者 经办人 日期     材料性质 起止页号 备注
-------------------------------------------------------------------------------------------------------------------
1     1、申报OA-毛喜瑞.pdf                                      -     -     20251110 -       -       -
2     2、业务（含产品、服务）洗钱风险评估表-毛喜瑞交易咨询项目.xlsx    -     -     20251110 -       -       -
3     3、合规审核意见截图.png                                    -     -     20251111 -       -       -
4     4、交易咨询部-魏子骥-期货交易咨询业务风险等级评估表.pdf          -     -     20251114 -       -       -
5     5、毛喜瑞伟交易咨询项目评审表决票.pdf                        -     -     20251110 -       -       -
6     6、期货交易咨询业务内部评审小组立项讨论会纪要-毛喜瑞交易咨询项目.pdf  -     -     20251114 -       -       -
7     7、用印OA-毛喜瑞与申银万国期货有限公司之期货交易咨询协议.pdf      -     -     20260127 -       -       -
8     8、毛喜瑞与申银万国期货有限公司之期货交易咨询协议.pdf            -     -     20251231 -       -       -

说明：此为纯文本预览，正式版为Word文档格式。
生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"""

text_path = os.path.join(project_path, "卷内文件目录_格式说明.txt")
with open(text_path, 'w', encoding='utf-8') as f:
    f.write(text_content)

print(f"\n✅ 格式说明文件已生成: {text_path}")