# _*_coding:utf-8_*_
import os
import time
from log import do_loger
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, numbers
from tqdm import tqdm
from typing import List, Union
from utils import *


class Gap:
    """包含wb[openpyxl],pw[pandas]
    """
    def __init__(self, excel_path: str):
        """
        传入表格路径
        @param excel_path:
        """
        print('开始读取表格，请耐心等待')
        if os.path.exists(excel_path):
            self.excel_path = excel_path
            self.wb = load_workbook(self.excel_path)
            self.pwb = pd.ExcelFile(self.excel_path)
            print(f'已成功读取表格: {excel_path}')
        else:
            print(f'{excel_path}路径文件不存在！')

def auto_save(func):
    def saver(*args, **kwargs):
        data = func(*args, **kwargs)
        try:
            print(f'开始数据保存，请耐心等待。。')
            args[0].wb.save(args[0].excel_path)
            print(f'🟢保存成功🟢')
        except IOError as e:
            print(f'🔴保存失败🔴==>{e}')
        return data
    return saver

@auto_save
def get_row_data_comment(gap: Gap, row_num: int, max_column_name: str):
    """
    横向读取表格批注，将批注写入下方单元格
    @param ws: sheet对象
    @param row_num: 读取第num行
    @param max_column_name:最大列数的名称
    @return:
    """
    _val = input_selector('请分别输入【表名 提取批注所在行 欲读取最大列（大写字母）】')
    ws_name = _val[:3]
    ws = gap.wb[ws_name]
    print(f'🍵开始写入批注🍵')
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    data = {}
    check_send = 0
    print(f'🌱开始处理：{ws.title}A-->Z列🌱')
    for pointer_index, pointer_column in enumerate(alphabet):
        if pointer_column == max_column_name:
            check_send = 1
            break
        coordinate = f'{pointer_column}{row_num}'
        field = ws[coordinate].value
        if isinstance(field, str) and field != '日期':
            field_comment = str(ws[coordinate].comment).replace('Comment: ', '').replace(' by Author',
                                                                                                '').replace(' ', '')
            ws[f'{pointer_column}{row_num + 1}'].value = field_comment
            comments = field_comment.split('\n')
            comment_dict = {comment.split('：')[0]: comment.split('：').pop() for comment in comments}
            data.update({field: comment_dict})
    if check_send != 1:
        print(f'🍀开始处理：{ws.title}AA-->ZZ列🍀')
        for pointer_index1 in range(len(alphabet)):
            for pointer_index2 in range(len(alphabet)):
                pointer_column = f'{alphabet[pointer_index1]}{alphabet[pointer_index2]}'
                if pointer_column == max_column_name:
                    break
                coordinate = f'{pointer_column}{row_num}'
                field = ws[coordinate].value
                field_comment = str(ws[coordinate].comment).replace('Comment: ', '').replace(' by Author',
                                                                                                    '').replace(' ',
                                                                                                                '')
                ws[f'{pointer_column}{row_num + 1}'].value = field_comment
                comments = field_comment.split('\n')
                comment_dict = {comment.split('：')[0]: comment.split('：').pop() for comment in comments}
                data.update({field: comment_dict})
    print('🟢处理完成🟢')
    return data

@auto_save
def set_gap_sheet(gap: Gap):
    """生产Gap表
    Args:
        gap (Gap): _传入一个gap实例对象_
    """
    _values = input_selector('请分别输入【系统表名 品牌表名】')
    dev_ws_name,uat_ws_name = _values[:2]
    # do_loger(f'🟡开始生成Gap_sheet🟡')
    dev_ws = gap.wb[dev_ws_name]
    uat_ws = gap.wb[uat_ws_name]
    wb = dev_ws.parent
    gap_ws = wb.create_sheet('set_gap')
    sr = set_az()
    max_col = min(dev_ws.max_column, uat_ws.max_column)
    max_row = max(dev_ws.max_row, uat_ws.max_row)
    # do_loger(f'🚀最大列{max_col}\n🚀最大行{max_row}')
    for gap_dev_col in tqdm(range(1, max_col * 3, 3),desc='处理进度'):
        gap_uat_col = gap_dev_col + 1
        gap_col = gap_dev_col + 2
        col_index = gap_dev_col // 3
        sheet_title = dev_ws[sr[col_index]+'1'].value
        # do_loger(f'🚀当前遍历列数{sr[col_index]},{sheet_title}')
        for irow in range(1, max_row + 1):
            # print(f'🚀当前遍历行数{irow}')
            dev_filed = f"='{dev_ws.title}'!{sr[col_index]}{irow}"#DEV表的原始引用公式
            uat_filed = f"='{uat_ws.title}'!{sr[col_index]}{irow}"#UAT表的原始引用公式
            gap_dev_cell = gap_ws.cell(row=irow, column=gap_dev_col)
            gap_uat_cell = gap_ws.cell(row=irow, column=gap_uat_col)
            gap_gap_cell = gap_ws.cell(row=irow, column=gap_col)
            # 将ws.cell()对象转换为ws[]对象
            gap_ws[gap_dev_cell.coordinate] = dev_filed  # 引用部分(将公式写入字段)
            gap_ws[gap_uat_cell.coordinate] = uat_filed  # 引用部分(将公式写入字段)
            dev_value = dev_ws[f'{sr[col_index]}{irow}'].value  # 纳入计算的值
            uat_value = uat_ws[f'{sr[col_index]}{irow}'].value  # 纳入计算的值
            # do_loger(f'\t🚀当前遍历行数{irow},dev:{dev_value},uat:{uat_value}')
            #先判断对比表，如果对比表没有值，就直接不计算Gap,但是前两个单元格仍旧引用数据
            if uat_value is None:#品牌字段为空
                gap_ws[gap_gap_cell.coordinate] = None#当对比品牌数据为空时，不进行Gap，将三列单元格全部置灰
                gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='BFBFBF',end_color='BFBFBF',fill_type='solid')
                gap_ws[gap_dev_cell.coordinate].fill = PatternFill(start_color='BFBFBF',end_color='BFBFBF',fill_type='solid')
                gap_ws[gap_uat_cell.coordinate].fill = PatternFill(start_color='BFBFBF',end_color='BFBFBF',fill_type='solid')
            else:# 品牌字段非空
                if not isinstance(dev_ws[f'{sr[col_index]}{irow}'].value, (int, float)):#处理字符串字段
                    # 非计算字段采用特殊公式标记成False,此行是为了写入表格，但是颜色需要单独处理
                    gap_ws[gap_gap_cell.coordinate] = f'=EXACT("{dev_value}","{uat_value}")'
                    # 处理颜色
                    if dev_value != uat_value:#标记False 为红色
                        gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FF7C80',end_color='FF7C80', fill_type='solid')
                    else:# True
                        #548235-font C6E0B4-bg
                        # gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='C6E0B4',end_color='C6E0B4', fill_type='solid')
                        pass
                else:# 省下的都是可计算值
                    dev_value = float(dev_value)
                    uat_value = float(uat_value)
                    if isinstance(dev_ws[f'{sr[col_index]}{irow}'].value, (int, float)) and uat_value !=0:#处理能计算的字段
                        #当品牌值不为零，且不等于TS值
                        gap_ws[gap_gap_cell.coordinate] = f'=IF({dev_value}=0,IF({uat_value}=0,0,IF({dev_value}>{uat_value},({uat_value}-{dev_value})/{dev_value},({uat_value}-{dev_value})/{uat_value})),IF({uat_value}=0,0,IF({dev_value}>{uat_value},({uat_value}-{dev_value})/{dev_value},({uat_value}-{dev_value})/{uat_value})))'
                        # 设置单元格格式
                        gap_ws[gap_gap_cell.coordinate].number_format = numbers.FORMAT_PERCENTAGE_00
                        # 根据计算结果-设背景颜色
                        try:
                            result = (uat_value - dev_value) / dev_value if dev_value != 0 else (0 if uat_value == 0 else ((uat_value - dev_value) / dev_value if dev_value > uat_value else (uat_value - dev_value) / uat_value)) if uat_value != 0 else 0
                            # result = (float(dev_value) - float(uat_value)) / float(uat_value)
                        except ZeroDivisionError:
                            result = 0.0  # 或者其他你认为合适的默认值
                        if result < -0.005:#小于使用黄色
                            gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FFCC00', end_color='FFCC00', fill_type='solid')
                        elif result > 0.005:#大于使用红色
                            gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FF7C80', end_color='FF7C80', fill_type='solid')
                    else:#品牌值=0，
                        gap_ws[gap_gap_cell.coordinate] = f'=IF({dev_value}=0,IF({uat_value}=0,0,IF({dev_value}>{uat_value},({uat_value}-{dev_value})/{dev_value},({uat_value}-{dev_value})/{uat_value})),IF({uat_value}=0,0,IF({dev_value}>{uat_value},({uat_value}-{dev_value})/{dev_value},({uat_value}-{dev_value})/{uat_value})))'
                        # 设置单元格格式
                        gap_ws[gap_gap_cell.coordinate].number_format = numbers.FORMAT_PERCENTAGE_00
                        try:
                            result = (uat_value - dev_value) / dev_value if dev_value != 0 else (0 if uat_value == 0 else ((uat_value - dev_value) / dev_value if dev_value > uat_value else (uat_value - dev_value) / uat_value)) if uat_value != 0 else 0
                            # result = (float(dev_value) - float(uat_value)) / float(uat_value)
                        except ZeroDivisionError:
                            result = 0.0  # 或者其他你认为合适的默认值
                        if result < -0.005:#小于使用黄色
                            gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FFCC00', end_color='FFCC00', fill_type='solid')
                        elif result > 0.005:#大于使用红色
                            gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FF7C80', end_color='FF7C80', fill_type='solid')
            # 处理非表头字体
            gap_ws[gap_dev_cell.coordinate].font = Font(size=8, bold=False, color='000000')
            gap_ws[gap_uat_cell.coordinate].font = Font(size=8, bold=False, color='000000')
            gap_ws[gap_gap_cell.coordinate].font = Font(size=8, bold=False, color='000000')#调整Gap单元格字体size
            if irow == 1:  # 处理头部
                dev_title = f'{dev_filed}&"-系统"'
                uat_title = f'{uat_filed}&"-品牌"'
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
def get_data_by_row_title(gap: Gap):
    ColorPrint.print("""
在这之前！你需要将品牌表头第2行，改成TS对应的mapping值！！这很重要
Mapping from sheet title data
匹配两个表头，将能匹配上的，数据传入另一个表头下方
@param ws_name:品牌sheetName
@param _is_reference: 1/0 (是/否开启全部匹配)
传参方式 品牌表名 1/0
    """, color='random')
    _values = input_selector('请分别输入【品牌表名  ？】')
    ws_name,_is_reference = _values[:2]
    _is_reference = int(_is_reference)
    if _is_reference:
        print('开启！')
    else:
        print('未开启')
        
    print(f'🟡开始匹配表头🟡{ws_name}->{ws_name}1')
    brand_sheet = gap.wb[ws_name]
    brand_max_row = brand_sheet.max_row
    system_sheet = gap.wb['system']
    #创建新表
    wb = brand_sheet.parent
    ts_sheet = wb.create_sheet(f'{ws_name}1')
    # 获取品牌表第二行数据
    brand_title = [cell.value for cell in brand_sheet[2]]
    # 获取系统表头字段
    system_title = [cell.value for cell in system_sheet[1]]
    for _title in tqdm(system_title, desc='匹配进度'):
        time.sleep(0.01)
        index_system = system_title.index(_title)# Bxx,Axx的数据 eg:A1
        ts_cell = ts_sheet.cell(row=1, column=index_system + 1)
        ts_cell.value = f'={system_sheet.title}!{ts_cell.coordinate}'
        
        if _title in brand_title:# 如果系统表头在品牌第二行
            index_brand = brand_title.index(_title)#索引-品牌
            if not _is_reference:
                _row = 2
                brand_cell = brand_sheet.cell(row=_row, column=index_brand + 1)#品牌第二行数据cell对象,kais
                # print(brand_cell.coordinate,brand_max_row)
                ts_cell_2 = ts_sheet.cell(row=_row, column=index_system + 1)
                ts_cell_2.value = f'={brand_sheet.title}!{brand_cell.coordinate}'
                print('执行高级代码')
            else:
                for _row in range(2,brand_max_row+1):
                    brand_cell = brand_sheet.cell(row=_row, column=index_brand + 1)#品牌第row行数据cell对象,kais
                    # print(brand_cell.coordinate,brand_max_row)
                    ts_cell_2 = ts_sheet.cell(row=_row, column=index_system + 1)
                    ts_cell_2.value = f'={brand_sheet.title}!{brand_cell.coordinate}'
                    print('执行垃圾代码')
        else:# 如果系统表头NOT 在品牌第二行
            ts_cell_2 = ts_sheet.cell(row=2, column=index_system + 1)
            ts_cell_2.value = '/'
    print('✨匹配完成✨')


@auto_save
def contrast_sheets(gap: Gap):
    """标注两列数据的差异
    Args:
        sheet_name1 (str): sheet1名称
        ts_column (str): sheet1_列名
        sheet_name2 (str): sheet2名称
        p_column (str): sheet2列名
    """
    ColorPrint.print("注意!!操作的数据列请确保格式完全一致，请排除空格引号等问题！\n"*3,color='random')
    _val1 = input_selector('请分别输入【表名1 匹配列名1（需要大写字母）】')
    try:
        sheet_name1,column_1 = _val1[:2]
        print('载入表1，成功')
    except Exception as e:
        print("载入表1错误",e)
        exit(1)
    _val2 = input_selector('请分别输入【表名2 匹配列名2（需要大写字母）】')
    try:
        sheet_name2,column_2 = _val2[:2]
        print('载入表2，成功')
    except Exception as e:
        print("载入表2错误",e)
        exit(1)
    fill = PatternFill("solid", fgColor="1874CD")  # 蓝色样式标记-品牌
    fill2 = PatternFill("solid", fgColor="FF0000")  # 红色样式标记-TS
    ws = gap.wb[sheet_name1]
    wp = gap.wb[sheet_name2]
    t_data = set(ws[column_1][2:ws.max_row])
    p_data = set(wp[column_2][2:wp.max_row])
    # print(f'🚀正在标注  {sheet_name1}  与  {sheet_name2}  差异！')
    for sty, tid in tqdm(enumerate(ws[column_1][2:ws.max_row], start=2)):
        if tid.value not in p_data:
            ws[f'{column_1}{sty}'].fill = fill2
    # print(f'🚀正在标注  {sheet_name2}  与  {sheet_name1}  差异！')
    for py, pid in tqdm(enumerate(wp[column_2][2:wp.max_row], start=2)):
        if pid.value not in t_data:
            wp[f'{column_2}{py}'].fill = fill


def set_gap_by_vlookup(gap: Gap):
    _values = input_selector('请分别输入【系统表名 品牌表名】')
    dev_ws_name,uat_ws_name = _values[:2]
    do_loger(f'🟡开始生成Gap_sheet🟡')
    dev_ws = gap.wb[dev_ws_name]
    uat_ws = gap.wb[uat_ws_name]
    wb = dev_ws.parent
    gap_ws = wb.create_sheet('set_gap')
    sr = set_sl()
    max_col = min(dev_ws.max_column, uat_ws.max_column)
    max_row = max(dev_ws.max_row, uat_ws.max_row)
    
    pass


get_row_data_comment.__name__='提取标题批注'
get_data_by_row_title.__name__='快速引用品牌列数据'
contrast_sheets.__name__='对比两列数据并标记'
set_gap_sheet.__name__='生成GapSheet[openpyxl]'
# set_gap_sheet1.__name__='生成GapSheet[pandas]'

fundict = {
    '1':get_row_data_comment,
    '2':get_data_by_row_title,
    '3':contrast_sheets,
    '4':set_gap_sheet,
    # '5':set_gap_sheet1,
}

def function_list(obj: Gap):
    ColorPrint.print('    ','='*10,'功能列表','='*10,color='yellow')
    for i in fundict:
        print()
        ColorPrint.print("          ",i,fundict[i].__name__,color='yellow')
    print()
    ColorPrint.print('    ','='*10,'功能列表','='*10,color='yellow')
    _x = input_selector("选择功能：")
    x = _x[:1][0]
    if x not in fundict.keys():
        clear()
        print('输入有误！，请重新选择')
        return function_list()
    else:
        try:
            return fundict[x](obj)
        except Exception as e:
            print('error-[也许你应该传入表名+（空格），再回车]:', e)
            exit(1)

if __name__ == '__main__':

    path = file_selector()
    print(path)
    obj = Gap(path)
    function_list(obj)
    
    
    
    # # obj.contrast_sheets(sheet_name1='品牌',ts_column='B',sheet_name2='TS',p_column='A')
    # get_data_by_row(obj,'品牌','品牌toTS')
    # set_gap_sheet(obj,'TS','品牌toTS1')
