# _*_coding:utf-8_*_

import pandas as pd
from tqdm import tqdm
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def set_brand1_true(path):
    # path = './植村秀-6.1-6.30品牌数据Gap.xlsx'
    with pd.ExcelFile(path) as excel_file:
        brand = excel_file.parse('brand')
        _t = brand.iloc[0].tolist()  # 获取第二行数据
        system = excel_file.parse('system')
        _t1 = system.columns.to_list()
        column_mapping = {col: _index for _index, col in enumerate(_t, start=0)}
        # 创建 brand1 表
        brand1 = pd.DataFrame(columns=_t1)
        for tcol in _t1:
            if tcol in column_mapping:
                # 从 brand 取出对应列数据，并存入 brand1
                brand_col_data = brand.iloc[:, column_mapping[tcol]].tolist()
                # brand1[col] = 
                # print(column_mapping[tcol],tcol,len(brand_col_data))
                for i in tqdm(range(len(brand_col_data)),desc=f'{tcol}列',ncols=120):
                    _target = f'={system}!{tcol}{i}'
                    brand1.at[i+1,tcol] = _target
        # 保存到新表 brand1
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            brand1.to_excel(writer, sheet_name='brand1', index=False)
    print("源文件保存成功！")


path = './植村秀-6.1-6.30品牌数据Gap.xlsx'

set_brand1_true(path)
