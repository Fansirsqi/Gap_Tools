# _*_coding:utf-8_*_
import os
import time

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, numbers
from tqdm import tqdm


# @PROJECT : DY_SCRM_TEST_PROJECT
# @Time : 2023/3/24 12:24
# @Author : Byseven
# @File : gaps.py
# @SoftWare:

class Gap:
    def __init__(self, excel_path: str):
        """
        传入表格路径
        @param excel_path:
        """
        if os.path.exists(excel_path):
            self.excel_path = excel_path
            self.wb = load_workbook(self.excel_path)
            self.pwb = pd.ExcelFile(self.excel_path)
            print('已成功读取表格')

        else:
            print(f'{excel_path}路径文件不存在！')
        # 用于处理批注导出的场景内表格
        self.comment = {
            '店铺销售概览': 'W',  # x+1
            '终端构成': 'F',
            '账号构成': 'O',
            '自营账号成交构成': 'U',
            '自营账号店铺直播成交概况': 'N',
            '店播日报': 'AX',
            '店播流量渠道概览': 'CT',
            '店播付费流量投放总览': 'K',
            '店播人群': 'AH',
        }

    @staticmethod
    def auto_save(func):
        def saver(*args, **kwargs):
            data = func(*args, **kwargs)
            try:
                args[0].wb.save(args[0].excel_path)
                print(f'🟢保存成功🟢')
            except IOError as e:
                print(f'🔴保存失败🔴==>{e}')
            return data

        return saver

    @staticmethod
    def get_filed_index_columns_as_pd(dfs, sheet_name: str, column_name: str, find_str: str) -> list:
        """
        获取指定列的所有单元格数据,查找指定字段所在单元格的位置坐标
        @param sheet_name: sheet名
        @param dfs:all表格对象
        @param column_name:指定列名
        @param find_str:查找字段
        @return:
        """
        # 获取所有sheet名称列表
        sheet_names = dfs.sheet_names

        # 读取指定的sheet数据
        df = dfs.parse(sheet_name)
        column_data = df[column_name].values.tolist()
        if find_str in column_data:
            cell_index = column_data.index(find_str)
            row_index = df.index[cell_index]
            col_index = df.columns.get_loc(column_name)
            return [row_index, col_index]
        else:
            return ['can find nothing']

    @staticmethod
    def get_file_list(path: str) -> list:
        """
        遍历返回子文件
        :param path:
        :return:
        """
        file_list = []
        for home, dirs, files in os.walk(path):
            for filename in files:
                # 文件名列表，包含完整路径
                file_list.append(os.path.join(home, filename))
                # # 文件名列表，只包含文件名
                # Filelist.append( filename)
        return file_list

    @staticmethod
    def get_column_data(ws, column: str, max_row: int):
        """获取某一列数据,从第二行开始,到X行结束"""
        data = []
        print(f'🟡正在获取第{column}列数据🟡')
        for i in range(2, max_row + 1):
            _data = ws[f'{column}{i}'].value
            data.append(_data)
        print('✅获取完毕✅')
        return data

    def get_row_data_comment(self, worksheet, row_num: int, max_column_name: str):
        """
        横向读取表格批注，将批注写入下方单元格，装饰器用于保存工作表
        @param worksheet: sheet对象
        @param row_num: 读取第num行
        @param max_column_name:最大列数的名称
        @return:
        """
        print(f'🍵开始写入批注🍵')
        alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        data = {}
        check_send = 0
        print(f'🌱开始处理：{worksheet.title}A-->Z列🌱')
        for pointer_index, pointer_column in enumerate(alphabet):
            if pointer_column == max_column_name:
                check_send = 1
                break
            coordinate = f'{pointer_column}{row_num}'
            field = worksheet[coordinate].value
            if isinstance(field, str) and field != '日期':
                field_comment = str(worksheet[coordinate].comment).replace('Comment: ', '').replace(' by Author',
                                                                                                    '').replace(' ', '')
                worksheet[f'{pointer_column}{row_num + 1}'].value = field_comment
                comments = field_comment.split('\n')
                comment_dict = {comment.split('：')[0]: comment.split('：').pop() for comment in comments}
                data.update({field: comment_dict})
        if check_send != 1:
            print(f'🍀开始处理：{worksheet.title}AA-->ZZ列🍀')
            for pointer_index1 in range(len(alphabet)):
                for pointer_index2 in range(len(alphabet)):
                    pointer_column = f'{alphabet[pointer_index1]}{alphabet[pointer_index2]}'
                    if pointer_column == max_column_name:
                        break
                    coordinate = f'{pointer_column}{row_num}'
                    field = worksheet[coordinate].value
                    field_comment = str(worksheet[coordinate].comment).replace('Comment: ', '').replace(' by Author',
                                                                                                        '').replace(' ',
                                                                                                                    '')
                    worksheet[f'{pointer_column}{row_num + 1}'].value = field_comment
                    comments = field_comment.split('\n')
                    comment_dict = {comment.split('：')[0]: comment.split('：').pop() for comment in comments}
                    data.update({field: comment_dict})
        print('🟢处理完成🟢')
        try:
            with open(self.excel_path, 'w') as f:
                self.wb.save(f)
            print(f'🟢保存成功🟢')
        except IOError as e:
            print(f'🔴保存失败🔴==>{e}')
        return data

    @staticmethod
    def paser_filed_data(data):
        """
        解析获取到的批注字段字典，转换成二维数组
        @param data: dict_by filed comment
        @return:
        """
        print('⚙️解析批注字典⚙️')
        all_row = []
        for filed, keys in data.items():
            rows_list = [filed]
            for i in keys.values():
                rows_list.append(i)
            if len(rows_list) < 5:
                for i in range(5 - len(rows_list)):
                    rows_list.append('')
            all_row.append(rows_list)
        print('🟢解析完成🟢️')
        return all_row

    @staticmethod
    def add_new_ws_by_pd(excel_path, all_row, ws_name):
        """
        在表格里新创建一张sheet,提取批注,至新sheet
        用于处理批注提取
        @param excel_path: 表格路径
        @param all_row: 二维数组
        @param ws_name: 创建sheet的名字
        @return:
        """
        print('🟡开始创建解析字典表格🟡')
        # df = pd.read_excel(excel_path)  # sheet_name='Sheet1'不添加将返回第一个
        # 创建新的 Excel 文件
        writer = pd.ExcelWriter(excel_path, mode='a')
        title = ['-', '字段定义', '字段来源', '滚动更新频次', '统计维度']
        df = pd.DataFrame(all_row, columns=title)
        # 将 DataFrame 写入新创建的 Excel 文件的工作表中
        df.to_excel(writer, sheet_name=f'{ws_name}', index=False)
        try:
            writer.close()
            print('🟢创建提取批注sheet完成🟢')
        except IOError as e:
            print(f'{e}')

    @staticmethod
    def set_sl():
        """
        生成一个列表[A-ZZ]
        @return: [A-ZZ]
        """
        sr = 'ABCDEFGHIJKLMNOPGRSTUVWXYZ'
        sr = list(sr)
        sl = []
        for i in sr:
            sl.append(i)
        for i in sr:
            for j in sr:
                k = i + j
                sl.append(k)
        # print(len(sl))
        return sl

    @auto_save
    def set_gap_sheet(self, dev_ws_name, uat_ws_name):
        """
        生成Gap_sheet
        @param dev_ws_name: ts表面
        @param uat_ws_name: 对比表名
        @return:
        """
        print(f'🟡开始生成Gap_sheet🟡')
        dev_ws = self.wb[dev_ws_name]
        uat_ws = self.wb[uat_ws_name]
        wb = dev_ws.parent
        gap_ws = wb.create_sheet('gap')
        sr = self.set_sl()
        max_col = min(dev_ws.max_column, uat_ws.max_column)
        max_row = max(dev_ws.max_row, uat_ws.max_row)
        print(f'🚀最大列{max_col}\n🚀最大行{max_row}')
        for gap_dev_col in tqdm(range(1, max_col * 3, 3)):
            gap_uat_col = gap_dev_col + 1
            gap_col = gap_dev_col + 2
            col_index = gap_dev_col // 3
            print(f'🚀当前遍历列数{sr[col_index]}')
            for irow in range(1, max_row + 1):
                # print(f'🚀当前遍历行数{irow}')
                dev_filed = f"='{dev_ws.title}'!{sr[col_index]}{irow}"
                uat_filed = f"='{uat_ws.title}'!{sr[col_index]}{irow}"
                gap_dev_cell = gap_ws.cell(row=irow, column=gap_dev_col)
                gap_uat_cell = gap_ws.cell(row=irow, column=gap_uat_col)
                gap_gap_cell = gap_ws.cell(row=irow, column=gap_col)
                # 将ws.cell()对象转换为ws[]对象
                gap_ws[gap_dev_cell.coordinate] = dev_filed  # 引用部分
                gap_ws[gap_uat_cell.coordinate] = uat_filed  # 引用部分
                dev_value = dev_ws[f'{sr[col_index]}{irow}'].value  # 计算值
                uat_value = uat_ws[f'{sr[col_index]}{irow}'].value  # 计算值
                if dev_value is None:
                    dev_value = 0
                if uat_value is None:
                    uat_value = 0
                if not isinstance(dev_ws[f'{sr[col_index]}{irow}'].value, (int, float)):
                    gap_ws[gap_gap_cell.coordinate] = f'=EXACT("{dev_value}","{uat_value}")'
                else:
                    if uat_value != 0:
                        gap_ws[gap_gap_cell.coordinate] = f'=({dev_value}-{uat_value})/{uat_value}'
                        # 设置单元格格式
                        gap_ws[gap_gap_cell.coordinate].number_format = numbers.FORMAT_PERCENTAGE_00
                        # 单元格条件
                        # print(dev_value, uat_value)
                        result = (dev_value - uat_value) / uat_value
                        if result < -0.005:
                            gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='DC143C', end_color='DC143C',
                                                                               fill_type='solid')
                        elif result > 0.005:
                            gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='4169E1', end_color='4169E1',
                                                                               fill_type='solid')
                    elif uat_value == 0:
                        gap_ws[gap_gap_cell.coordinate].value = 0
                        # 设置单元格格式
                        gap_ws[gap_gap_cell.coordinate].number_format = numbers.FORMAT_PERCENTAGE_00
                if irow == 1:  # 处理头部
                    dev_title = f'{dev_filed}&"-小时"'
                    uat_title = f'{uat_filed}&"-天"'
                    # print(f'dev-title-->{dev_title}')
                    # print(f'uat-title-->{uat_title}')
                    gap_ws[gap_dev_cell.coordinate].value = dev_title  # dev头
                    gap_ws[gap_uat_cell.coordinate].value = uat_title  # uat头
                    gap_ws[gap_gap_cell.coordinate].value = 'Gap'

                    gap_ws[gap_dev_cell.coordinate].font = Font(size=10, bold=True, color='000000')
                    gap_ws[gap_uat_cell.coordinate].font = Font(size=10, bold=True, color='000000')
                    gap_ws[gap_gap_cell.coordinate].font = Font(size=10, bold=True, color='000000')

                    gap_ws[gap_dev_cell.coordinate].fill = PatternFill(start_color='00a67d', end_color='00a67d',
                                                                       fill_type='solid')
                    gap_ws[gap_uat_cell.coordinate].fill = PatternFill(start_color='FF7F50', end_color='FF7F50',
                                                                       fill_type='solid')
                    gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='df3079', end_color='df3079',
                                                                       fill_type='solid')

        print('🎁Gap_Sheet生成成功🎁')

    @auto_save
    def get_data_by_row(self, ws_name, ws_name2):
        """
        匹配两个表头，将能匹配上的，数据传入另一个表头下方
        @param ws_name:
        @param ws_name2:
        """
        print(f'🟡开始匹配表头🟡{ws_name}-{ws_name2}')
        ws = self.wb[ws_name]
        ws2 = self.wb[ws_name2]
        _keys = []
        keys = []
        _keys2 = []
        keys2 = []
        # print(ws.max_column)
        for i in range(1, ws.max_column):
            key = ws.cell(row=1, column=i)
            _keys.append(key)
            keys.append(key.value)
        for j in range(1, ws2.max_column):
            key = ws2.cell(row=1, column=j)
            _keys2.append(key)
            keys2.append(key.value)
        # print(ws[first[0].coordinate].value, first[0].coordinate)
        # print(ws2[secend[0].coordinate].value, secend[0].coordinate)
        for c in tqdm(keys2):
            time.sleep(0.01)
            if c in keys:
                n1 = keys.index(c)  # 欲遍历的字段，在匹配字段列表中的索引
                v_cell = _keys[n1]  # 根据索引取出sheet中单元格cell对象
                # print(f'当前指向值===>{c}')
                # print(f'对应数据  {v_cell.coordinate}')
                n0 = keys2.index(c)
                v0_cell = _keys2[n0]
                # print(f'原始表格中的对应数据  {v0_cell.coordinate}')
                to_ws2 = ws2.cell(row=v0_cell.row + 1, column=v0_cell.column)
                to_ws2.value = f'={ws.title}!{v_cell.coordinate}'
                # print(ws[key3.coordinate].value, key3.coordinate)
        print('✨表头匹配完成✨')

    @auto_save
    def replace_column_data(self, ws, column, max_row: int, config_dict: dict):
        """
        替换某一列数据，配置好映射字典即可，本函数不会保存表格，需要手动保存
        @param config_dict: 配置替换字典
        @param ws: 表名
        @param column:列名称
        @param max_row: 最大行号
        @return:
        """
        print('🟡开始替换🟡')
        for i in tqdm(range(2, max_row + 1)):
            _data = ws[f'{column}{i}'].value
            if _data in config_dict.keys():
                ws[f'{column}{i}'].value = config_dict[_data]
        print('🟢替换完成🟢')


if __name__ == '__main__':
    path = r'C:\Users\admin\WorkDate\2023\04-23-巴欧分钟级4期Gap表完善\分钟监测-UAT-客户与导出数据GAP_4.23-V1.5.xlsx'
    a = Gap(path)

    df = a.pwb.parse('原因排查3.28')
    p = df.loc[2]
    print(type(p))
    # column_data = df['报表展示字段'].values.tolist()
    # for filed in column_data:
    #     # print(filed)
    #     b = a.get_filed_index_columns_as_pd(
    #         dfs=a.pwb,
    #         sheet_name='原因排查3.28',
    #         column_name='报表展示字段',
    #         find_str=filed
    #     )
    #     print(filed, b)
