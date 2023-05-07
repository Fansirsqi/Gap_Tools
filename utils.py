# _*_coding:utf-8_*_
import json
import pandas as pd

# @PROJECT : Gap_Tools
# @Time : 2023/4/24 11:22
# @Author : Byseven
# @File : utils.py
# @SoftWare:
class GapUtils:

    @staticmethod
    def read_local_json(file_path):
        """
        读取本地json
        @param file_path:
        @return: dict
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            json_str = f.read()
        return json.loads(json_str)
    
    @staticmethod
    def mapping_title(file_path: str,n: int, data_dict: dict):
        """mapping title

        Args:
            file_path (str): file path
            n (int): colunm
        """
        df = pd.read_excel(file_path,header=n,sheet_name='Sheet1')
        for col in df.columns:
            # print(col)
            if col in data_dict:
                # df.iloc[0, df.columns.get_loc(col)] = data_dict[col]
                print(data_dict[col], end="\t")
            else:
                print('NaN', end="\t")
                # df.iloc[0, df.columns.get_loc(col)] = '未能匹配'

if __name__ == '__main__':
    obj = GapUtils()
    data = obj.read_local_json('ts.json')
    obj.mapping_title(r'C:\Users\admin\Documents\WXWork\1688855561385475\Cache\File\2023-05\vs.xlsx',1,data)
    