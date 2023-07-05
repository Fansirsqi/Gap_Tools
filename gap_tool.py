# _*_coding:utf-8_*_

import pandas as pd

path = './植村秀-6.1-6.30品牌数据Gap.xlsx'

df = pd.read_excel(path)
excel_file = pd.ExcelFile(path)

brand = excel_file.parse('brand')
_t = brand.iloc[0].tolist()#获取第二行数据

brand1 = excel_file.parse('brand1')
_t1 = brand1.columns.to_list()

# s = brand['素材名称'].tolist()

column_mapping = {col: _index for _index, col in enumerate(_t, start=0)}

for col in _t1:
    if col in column_mapping:
        brand1[col] = brand.iloc[:, column_mapping[col]].tolist()

# 保存源文件
output_path = './modified_植村秀-6.1-6.30品牌数据Gap.xlsx'
with pd.ExcelWriter(output_path) as writer:
    writer.book = excel_file.book
    writer.sheets = excel_file.sheets
    brand1.to_excel(writer, sheet_name='brand1', index=False)

print("源文件保存成功！")