import os
from datetime import datetime

# 常量
YEAR = "2025"  # 归档年份
PROJECT_NUM = "0001"  # 毛喜瑞项目分配为0001
PROJECT_NAME = "毛喜瑞交易咨询项目"
CLIENT_CODE = "8005108018"  # 正确代码
PRODUCT_NAME = "毛喜瑞交易咨询项目"  # 产品名称

# 档号格式
ARCHIVE_CODE = f"SWQH-YW·{YEAR}-D30-3-1009-{PROJECT_NUM}"

# 文件列表 (从之前的脚本获取)
file_list = [
    {"序号": "0001", "文件题名": "1、申报OA-毛喜瑞.pdf", "日期": "20251110"},
    {"序号": "0002", "文件题名": "2、业务（含产品、服务）洗钱风险评估表-毛喜瑞交易咨询项目.xlsx", "日期": "20251110"},
    {"序号": "0003", "文件题名": "3、合规审核意见截图.png", "日期": "20251111"},
    {"序号": "0004", "文件题名": "4、交易咨询部-魏子骥-期货交易咨询业务风险等级评估表.pdf", "日期": "20251114"},
    {"序号": "0005", "文件题名": "5、毛喜瑞伟交易咨询项目评审表决票.pdf", "日期": "20251114"},
    {"序号": "0006", "文件题名": "6、期货交易咨询业务内部评审小组立项讨论会纪要-毛喜瑞交易咨询项目.pdf", "日期": "20251114"},
    {"序号": "0007", "文件题名": "7、用印OA-毛喜瑞与申银万国期货有限公司之期货交易咨询协议.pdf", "日期": "20260127"},
    {"序号": "0008", "文件题名": "8、毛喜瑞与申银万国期货有限公司之期货交易咨询协议.pdf", "日期": "20251231"},
]

# 生成HTML格式用于查看
html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>卷内文件目录 - {PROJECT_NAME}</title>
    <style>
        body {{ font-family: 'SimSun', sans-serif; margin: 40px; }}
        h1 {{ text-align: center; margin-bottom: 30px; }}
        .info-table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; }}
        .info-table td {{ padding: 8px; border: 1px solid #ccc; }}
        .info-table .label {{ width: 20%; background-color: #f5f5f5; font-weight: bold; }}
        .file-table {{ width: 100%; border-collapse: collapse; }}
        .file-table th, .file-table td {{ padding: 8px; border: 1px solid #ccc; text-align: center; }}
        .file-table th {{ background-color: #e0e0e0; font-weight: bold; }}
        .file-table tr:nth-child(even) {{ background-color: #f9f9f9; }}
        .timestamp {{ margin-top: 40px; text-align: right; color: #666; font-size: 14px; }}
    </style>
</head>
<body>
    <h1>卷内文件目录</h1>
    
    <table class="info-table">
        <tr>
            <td class="label">客户名称</td>
            <td>{PROJECT_NAME}</td>
            <td class="label">档号</td>
            <td>{ARCHIVE_CODE}</td>
        </tr>
        <tr>
            <td class="label">案卷题名</td>
            <td>{PRODUCT_NAME}</td>
            <td class="label">客户代码</td>
            <td>{CLIENT_CODE}</td>
        </tr>
    </table>
    
    <table class="file-table">
        <thead>
            <tr>
                <th width="80">序号</th>
                <th>文件题名</th>
                <th width="120">日期</th>
                <th width="120">责任者</th>
                <th width="120">经办人</th>
                <th width="100">起止页号</th>
                <th width="150">备注</th>
            </tr>
        </thead>
        <tbody>
"""

for file_item in file_list:
    html_content += f"""
            <tr>
                <td>{file_item['序号']}</td>
                <td style="text-align: left;">{file_item['文件题名']}</td>
                <td>{file_item['日期']}</td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
            </tr>"""

html_content += f"""
        </tbody>
    </table>
    
    <div class="timestamp">
        生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br>
        归档年份: {YEAR}
    </div>
</body>
</html>"""

# 保存为HTML文件便于预览
output_path = r"D:\申银万国期货\3.档案管理\2、业务档案\2. 未归档项目-2025\毛喜瑞交易咨询项目\卷内文件目录_预览.html"
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"HTML预览文件已生成: {output_path}")

# 同时生成纯文本版本
txt_content = f"""卷内文件目录

客户名称：{PROJECT_NAME}
档号：{ARCHIVE_CODE}
案卷题名：{PRODUCT_NAME}
客户代码：{CLIENT_CODE}

序号    文件题名                                                日期      责任者 经办人 起止页号 备注
{"-"*120}"""

for file_item in file_list:
    filename = file_item['文件题名']
    # 格式化对齐
    seq = file_item['序号']
    date = file_item['日期']
    txt_content += f"\n{seq:4}    {filename:60}    {date}        -       -       -       -"

txt_output = r"D:\申银万国期货\3.档案管理\2、业务档案\2. 未归档项目-2025\毛喜瑞交易咨询项目\卷内文件目录_文本版.txt"
with open(txt_output, 'w', encoding='utf-8') as f:
    f.write(txt_content)

print(f"文本版文件已生成: {txt_output}")

# 生成CSV版本便于导入Excel
csv_content = "序号,文件题名,日期,责任者,经办人,起止页号,备注\n"
for file_item in file_list:
    csv_content += f"{file_item['序号']},{file_item['文件题名']},{file_item['日期']},,,,\n"

csv_output = r"D:\申银万国期货\3.档案管理\2、业务档案\2. 未归档项目-2025\毛喜瑞交易咨询项目\卷内文件目录_数据.csv"
with open(csv_output, 'w', encoding='utf-8') as f:
    f.write(csv_content)

print(f"CSV数据文件已生成: {csv_output}")

print(f"\n=== 关键信息汇总 ===")
print(f"档号: {ARCHIVE_CODE}")
print(f"客户: {PROJECT_NAME}")
print(f"代码: {CLIENT_CODE}")
print(f"文件数量: {len(file_list)}")
print(f"时间范围: {file_list[0]['日期']} 至 {file_list[-1]['日期']}")