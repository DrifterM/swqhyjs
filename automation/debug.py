import pandas as pd

df = pd.read_excel('data_source/futures_data.xlsx', sheet_name=0, header=None)

print("Excel data preview:")
print("First 10 rows:")
for i in range(10):
    row = df.iloc[i]
    print(f"Row {i}: {row.tolist()}")

print("\n\nData analysis:")
print(f"Total rows: {len(df)}")
print(f"Columns: {len(df.columns)}")

# 检查标题行
print("\nTitle row (row 1):")
headers = df.iloc[1].tolist()
for i, h in enumerate(headers):
    print(f"  Col {i}: {h} (type: {type(h)})")

# 检查几个数据行
print("\nData rows:")
for i in range(2, 7):
    symbol = df.iloc[i, 1]
    name = df.iloc[i, 2]
    print(f"Row {i}: Symbol='{symbol}', Name='{name}' (type: {type(name)})")