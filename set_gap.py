import openpyxl
from openpyxl import load_workbook
from utils import *
wb = load_workbook('./植村秀-6.1-6.30品牌数据Gap.xlsx')

for i in wb:
    print(i.title)
    
ws_brand = wb['brand1']
ws_system = wb['system']
ws_gap = wb.create_sheet('gap')
#取最大列，最大行
max_col = max(ws_brand.max_column, ws_system.max_column)
max_row = max(ws_brand.max_row, ws_system.max_row)
sr = set_az()
system_title = [cell.value for cell in ws_system[1]]
brand_title = [cell.value for cell in ws_brand[2]]
print(len(sr),system_title,'\n',brand_title)
for i in range(1,max_col*3,3):#此处i代表新表中的每一列的开头一列
    # print(i)
    brand_col = i
    system_col = i+1
    gap_col = i+2
    _index = i//3
    quote_col_name = sr[_index] #需要引入的列名
    
    for irow in range(1,max_row+1):
        _col_name = f'{quote_col_name}{irow}'
        # 前两列需要引入的公式
        quite_brand = f'={ws_brand.title}!{_col_name}'
        quite_system = f'={ws_system.title}!{_col_name}'
        quite_gap = f'='
        gap_brand_cell = ws_gap.cell(row=irow,column=brand_col)
        gap_system_cell = ws_gap.cell(row=irow,column=system_col)
        gap_gap_cell = ws_gap.cell(row=irow,column=gap_col)
        # ws_gap[gap_brand_cell.coordinate]
        gap_brand_cell.value = quite_brand
        gap_system_cell.value = quite_system
        gap_gap_cell.value = quite_gap
print(len(sr))