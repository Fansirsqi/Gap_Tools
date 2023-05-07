import os

import pandas as pd
from openpyxl.reader.excel import load_workbook
import utils


class GapTools:
    def __init__(self, excel_path: str):
        """
        传入表格路径
        @param excel_path:
        """
        if os.path.exists(excel_path):
            self.excel_path = excel_path
            self.wb = load_workbook(self.excel_path)
            self.pwb = pd.ExcelFile(self.excel_path)
            print('已成功读取表格对象')
        else:
            print(f'{excel_path}路径文件不存在！')


if __name__ == '__main__':
    ut = utils.GapUtils()
    dj = ut.read_local_json('baiku.json')
    print(dj)
    # path = r'C:\Users\admin\WorkDate\2023\04-23-巴欧分钟级4期Gap表完善\分钟监测-UAT-客户与导出数据GAP_4.23-V1.5.xlsx'
    # obj = GapTools(path)
    # df = obj.pwb.parse('原因排查3.28')
    # print(df)
    # print(df.loc[0])
    # for line in df.values:
    #     print(line)
