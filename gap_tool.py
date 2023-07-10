# _*_coding:utf-8_*_
import os
import pandas as pd
from tqdm import tqdm
import pandas as pd


def set_brand_true(path,brand_name='brand'):
    with pd.ExcelFile(path) as excel_file:
        brand = excel_file.parse(f'{brand_name}')
        _t = brand.iloc[0].tolist()  # 获取第二行数据,表头默认为索引
        system = excel_file.parse('system')
        _t1 = system.columns.to_list()
        brand1 = pd.DataFrame(columns=_t1)
        column_mapping = {col: _index for _index, col in enumerate(_t, start=0)}
        for tcol in _t1:# 将TS表头匹配品牌表头，如果有值，就拿到列名称
            if tcol in column_mapping:
                _brand_index = column_mapping[tcol] #对应列在品牌表中的索引  
                for i in range(brand.shape[0]):
                    relative_reference = f"='{brand_name}'!R{i+1}C{_brand_index+1}"  # 相对单元格引用
                    brand1.at[i, tcol] = relative_reference
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            brand1.to_excel(writer, sheet_name='brand1', index=False)
    print("源文件保存成功！")


path = './植村秀-6.1-6.30品牌数据Gap.xlsx'
# set_brand1_true(path)
