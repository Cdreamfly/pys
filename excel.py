# 导入pandas库，简写为pd
import pandas as pd

# 读取表格,使用openpyxl引擎，获取表名为表1的内容

df = pd.read_excel("data.xlsx",index_col=0)

# 读取指定行，读取第一行数据，表头外的第一行（pandas读取表格默认不读取表头，即第一行）
row_data = df.loc[[71]].values
for cel in row_data:
    print(cel)

address = cel[4]
name = cel[10]
p = cel[14]
area = cel[16]
year = cel[17]

print(address,name,p,area,year)