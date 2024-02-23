import pandas as pd
from utils import list_files
# df = pd.read_excel('.\欧莱雅染发甄选\【抖音看板】店铺销售概览_2024-01-01_2024-01-31.xlsx',sheet_name='店铺销售概览')

fns = list_files(path='.\欧莱雅染发甄选', include=['自营', '构成', '销售概览'], is_prt=False)
dfs = []
for fn in fns:
    df = pd.read_excel(fn)
    if 'Date' in df.columns:
        df = df.set_index('日期')
    dfs.append(df)
# 横向拼接所有sheet的数据
result = pd.concat(dfs, axis=1)

# 打印拼接后的数据的前几行
# print(result.head())
result.to_excel('.\欧莱雅染发甄选\[合并]销售概览_2024-01-01_2024-01-31.xlsx', index=False)

# print(df)
