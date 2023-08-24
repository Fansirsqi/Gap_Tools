# _*_coding:utf-8_*_
import os
import time
from datetime import datetime
import re
from log import *
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, numbers, NamedStyle
from collections import defaultdict
from Sheetstyle import *
from utils import *
from rich.progress import track
import logging

datetime_style = NamedStyle(name='datetime_style', number_format='YYYY-MM-DD')


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


def parse_value(value_str):
    if isinstance(value_str, (datetime, )):
        logging.debug(f'{value_str}是日期格式')
        do_logger(f'{value_str}是日期格式')
        return value_str
    # 判断是否为百分比形式，例如 "2.4%, 百分数%"
    if "%" in value_str:
        cleaned_value = value_str.replace('%', '').strip()
        try:
            converted_value = float(cleaned_value) / 100.0
            return converted_value
        except ValueError:
            return value_str
    # 判断是否为包含逗号的数字形式，例如 "178,692.4"
    if "," in value_str:
        cleaned_value = float(value_str.replace(',', ''))
        return cleaned_value
    try:
        converted_value = float(value_str)
        return converted_value
    except ValueError:
        return value_str


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
            field_comment = str(ws[coordinate].comment).replace('Comment: ', '').replace(' by Author', '').replace(' ', '')
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
                field_comment = str(ws[coordinate].comment).replace('Comment: ', '').replace(' by Author', '').replace(' ', '')
                ws[f'{pointer_column}{row_num + 1}'].value = field_comment
                comments = field_comment.split('\n')
                comment_dict = {comment.split('：')[0]: comment.split('：').pop() for comment in comments}
                data.update({field: comment_dict})
    print('🟢处理完成🟢')
    return data


@auto_save
def set_gap_sheet(gap: Gap):
    """生产Gap表,要求两个sheet数据绝对的对其
    Args:
        gap (Gap): _传入一个gap实例对象_
    """
    _values = input_selector('请分别输入【系统表名 品牌表名】')
    dev_ws_name, uat_ws_name = _values[:2]
    do_logger(f'🟡开始生成Gap_sheet🟡')
    dev_ws = gap.wb[dev_ws_name]
    uat_ws = gap.wb[uat_ws_name]
    wb = dev_ws.parent
    gap_ws = wb.create_sheet('set_gap')
    sr = set_az()
    new_rule = True  # 新对数规则
    secend_check = True  # 第二次检查
    max_col = min(dev_ws.max_column, uat_ws.max_column)
    max_row = max(dev_ws.max_row, uat_ws.max_row)
    do_logger(f'🚀最大列{max_col}\n🚀最大行{max_row}')
    for gap_dev_col in track(range(1, max_col * 3, 3), description='处理进度'):
        if gap_dev_col == 1 and secend_check:  # 如果当前列号=1,就新增一行
            # 新增一行空白行
            gap_ws.insert_rows(1)
            secend_check = False
        gap_uat_col = gap_dev_col + 1
        gap_col = gap_dev_col + 2
        col_index = gap_dev_col // 3
        sheet_title = dev_ws[sr[col_index] + '1'].value
        do_logger(f'🚀当前遍历列数{sr[col_index]},表头->{sheet_title}')
        for irow in range(1, max_row + 1):
            # do_logger(f'🚀当前遍历行数{irow}')
            # DEV表的原始引用公式
            dev_filed = f"='{dev_ws.title}'!{sr[col_index]}{irow}"
            # UAT表的原始引用公式
            uat_filed = f"='{uat_ws.title}'!{sr[col_index]}{irow}"
            gap_dev_cell = gap_ws.cell(row=irow, column=gap_dev_col)
            gap_uat_cell = gap_ws.cell(row=irow, column=gap_uat_col)
            gap_gap_cell = gap_ws.cell(row=irow, column=gap_col)
            # 将ws.cell()对象转换为ws[]对象
            gap_ws[gap_dev_cell.coordinate] = dev_filed  # 引用部分(将公式写入字段)
            gap_ws[gap_uat_cell.coordinate] = uat_filed  # 引用部分(将公式写入字段)
            dev_value = dev_ws[f'{sr[col_index]}{irow}'].value  # 纳入计算的值
            uat_value = uat_ws[f'{sr[col_index]}{irow}'].value  # 纳入计算的值

            if irow == 1:  # 处理头部
                if new_rule:
                    dev_title = f'{dev_filed}'
                    # uat_title = f'{uat_filed}&"-品牌"'
                    # gap_ws[gap_dev_cell.coordinate].value = dev_title  # dev头
                    gap_ws.cell(row=irow, column=gap_dev_col).value = dev_title
                    gap_ws.cell(row=irow + 1, column=gap_dev_col).value = "品牌数据"
                    gap_ws.cell(row=irow + 1, column=gap_uat_col).value = "系统数据"
                    gap_ws.cell(row=irow + 1, column=gap_col).value = "Gap"
                    # gap_ws[gap_uat_cell.coordinate].value = uat_title  # uat头
                    # gap_ws[gap_gap_cell.coordinate].value = 'Gap'

                    gap_ws.cell(row=irow + 1, column=gap_dev_col).font = Font(size=10, bold=False, color='000000')
                    gap_ws.cell(row=irow + 1, column=gap_uat_col).font = Font(size=10, bold=False, color='000000')
                    gap_ws.cell(row=irow + 1, column=gap_col).font = Font(size=10, bold=False, color='000000')
                    gap_ws[gap_dev_cell.coordinate].font = Font(size=10, bold=True, color='000000')
                    gap_ws[gap_uat_cell.coordinate].font = Font(size=10, bold=True, color='000000')
                    gap_ws[gap_gap_cell.coordinate].font = Font(size=10, bold=True, color='000000')
                else:
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
                    gap_ws[gap_dev_cell.coordinate].fill = PatternFill(start_color='00a67d', end_color='00a67d', fill_type='solid')
                    gap_ws[gap_uat_cell.coordinate].fill = PatternFill(start_color='FF7F50', end_color='FF7F50', fill_type='solid')
                    gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='df3079', end_color='df3079', fill_type='solid')
            else:  #不是第一行
                # 先判断对比表，如果对比表没有值，就直接不计算Gap,但是前两个单元格仍旧引用数据
                if uat_value is None:  # 品牌字段为空
                    # 当对比品牌数据为空时，不进行Gap，将三列单元格全部置灰
                    gap_ws[gap_gap_cell.coordinate] = None
                    gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')
                    gap_ws[gap_dev_cell.coordinate].fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')
                    gap_ws[gap_uat_cell.coordinate].fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')
                else:  # 品牌字段非空
                    dev_value = parse_value(dev_value)
                    uat_value = parse_value(uat_value)
                    do_logger(f"\t🚀当前遍历行数{irow}, dev:{dev_value} type:{type(dev_value)}, uat:{uat_value} type:{type(uat_value)}")
                    if not isinstance(uat_value, (
                        int,
                        float,
                        )):  # 处理字符串字段(不可计算)
                        if isinstance(dev_value, (datetime)) or isinstance(uat_value, (datetime, )):
                            gap_dev_cell.style = datetime_style  # 添加检测单元格日期格式并且设置格式
                            gap_uat_cell.style = datetime_style
                        # 非计算字段采用特殊公式标记成False,此行是为了写入表格，但是颜色需要单独处理
                        gap_ws[gap_gap_cell.coordinate] = f'=EXACT("{dev_value}","{uat_value}")'
                        # 处理颜色
                        if dev_value != uat_value:  # 标记False 为红色
                            gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FF7C80', end_color='FF7C80', fill_type='solid')
                        else:  # True
                            # 548235-font C6E0B4-bg
                            # gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='C6E0B4',end_color='C6E0B4', fill_type='solid')
                            pass
                    else:  # 省下的都是可计算值
                        dev_value = float(dev_value)
                        uat_value = float(uat_value)
                        # 处理能计算的字段
                        if isinstance(dev_ws[f'{sr[col_index]}{irow}'].value, (int, float)) and uat_value != 0:
                            # 当品牌值不为零，且不等于TS值
                            gap_ws[
                                gap_gap_cell.coordinate
                                ] = f'=IF({dev_value}=0,IF({uat_value}=0,0,IF({dev_value}>{uat_value},({uat_value}-{dev_value})/{dev_value},({uat_value}-{dev_value})/{uat_value})),IF({uat_value}=0,0,IF({dev_value}>{uat_value},({uat_value}-{dev_value})/{dev_value},({uat_value}-{dev_value})/{uat_value})))'
                            # 设置单元格格式
                            gap_ws[gap_gap_cell.coordinate].number_format = numbers.FORMAT_PERCENTAGE_00
                            # 根据计算结果-设背景颜色
                            try:
                                result = (uat_value - dev_value) / dev_value if dev_value != 0 else (
                                    0 if uat_value == 0 else ((uat_value - dev_value) / dev_value if dev_value > uat_value else (uat_value - dev_value) / uat_value)
                                    ) if uat_value != 0 else 0
                                # result = (float(dev_value) - float(uat_value)) / float(uat_value)
                            except ZeroDivisionError:
                                result = 0.0  # 或者其他你认为合适的默认值
                            if result < -0.005:  # 小于使用黄色
                                gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FFCC00', end_color='FFCC00', fill_type='solid')
                            elif result > 0.005:  # 大于使用红色
                                gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FF7C80', end_color='FF7C80', fill_type='solid')
                        else:  # 品牌值=0，
                            gap_ws[
                                gap_gap_cell.coordinate
                                ] = f'=IF({dev_value}=0,IF({uat_value}=0,0,IF({dev_value}>{uat_value},({uat_value}-{dev_value})/{dev_value},({uat_value}-{dev_value})/{uat_value})),IF({uat_value}=0,0,IF({dev_value}>{uat_value},({uat_value}-{dev_value})/{dev_value},({uat_value}-{dev_value})/{uat_value})))'
                            # 设置单元格格式
                            gap_ws[gap_gap_cell.coordinate].number_format = numbers.FORMAT_PERCENTAGE_00
                            try:
                                result = (uat_value - dev_value) / dev_value if dev_value != 0 else (
                                    0 if uat_value == 0 else ((uat_value - dev_value) / dev_value if dev_value > uat_value else (uat_value - dev_value) / uat_value)
                                    ) if uat_value != 0 else 0
                                # result = (float(dev_value) - float(uat_value)) / float(uat_value)
                            except ZeroDivisionError:
                                result = 0.0  # 或者其他你认为合适的默认值
                            if result < -0.005:  # 小于使用黄色
                                gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FFCC00', end_color='FFCC00', fill_type='solid')
                            elif result > 0.005:  # 大于使用红色
                                gap_ws[gap_gap_cell.coordinate].fill = PatternFill(start_color='FF7C80', end_color='FF7C80', fill_type='solid')
                    # 处理非表头字体
                    gap_ws[gap_dev_cell.coordinate].font = Font(size=8, bold=False, color='000000')
                    gap_ws[gap_uat_cell.coordinate].font = Font(size=8, bold=False, color='000000')
                    gap_ws[gap_gap_cell.coordinate].font = Font(size=8, bold=False, color='000000')  # 调整Gap单元格字体size

    print('🎁Gap_Sheet生成成功🎁')


@auto_save
def get_data_by_row_title(gap: Gap):
    ColorPrint.print(
        """
在这之前！你需要将品牌表头第2行，改成TS对应的mapping值！！这很重要
Mapping from sheet title data
匹配两个表头，将能匹配上的，数据传入另一个表头下方
@param ws_name:品牌sheetName
@param _is_reference: 1/0 (是/否开启全部匹配)
传参方式 品牌表名 1/0
    """,
        color='random'
        )
    _values = input_selector('请分别输入【品牌表名  ？】')
    ws_name, _is_reference = _values[:2]
    _is_reference = int(_is_reference)
    if _is_reference:
        print('开启！')
    else:
        print('未开启')
    print(f'🟡开始匹配表头🟡{ws_name}->{ws_name}1')
    brand_sheet = gap.wb[ws_name]
    brand_max_row = brand_sheet.max_row
    system_sheet = gap.wb['system']
    # 创建新表
    wb = brand_sheet.parent
    ts_sheet = wb.create_sheet(f'{ws_name}1')
    # 获取品牌表第二行数据
    brand_title = [cell.value for cell in brand_sheet[2]]
    # 获取系统表头字段
    system_title = [cell.value for cell in system_sheet[1]]
    # 统计每个系统表头字段的出现次数
    counter = defaultdict(int)
    for cell in system_sheet[1]:
        counter[cell.value] += 1
    # 遍历system_title，并为重复的字段添加后缀
    unique_system_title = []
    for title in system_title:
        if counter[title] > 1:
            suffix = str(counter[title])
            counter[title] -= 1
            title += suffix
        unique_system_title.append(title)
    for _title in track(unique_system_title, description='匹配进度'):  # 遍历系统表头
        time.sleep(0.01)
        index_system = unique_system_title.index(_title)  # Bxx,Axx的数据 eg:A1
        ts_cell = ts_sheet.cell(row=1, column=index_system + 1)
        # --》第一行引入TS表头
        ts_cell.value = f'={system_sheet.title}!{ts_cell.coordinate}'
        v = re.sub(r"\d", "", _title)
        ts_sheet.cell(row=2, column=index_system + 1).value = v  # 第二行，表头，文本内容
        if _title in brand_title:  # 如果系统表头在品牌第二行,这里默认只匹配第一个字段，如果有相同字段
            index_brand = brand_title.index(_title)  # 索引-品牌
            _row = 3  # 第三行
            brand_cell = brand_sheet.cell(row=_row, column=index_brand + 1)  # 品牌第3行数据cell对象,kais
            print(brand_cell.coordinate, brand_max_row)
            ts_cell_2 = ts_sheet.cell(row=_row, column=index_system + 1)
            ts_cell_2.value = f'={brand_sheet.title}!{brand_cell.coordinate}'
            do_logger(f'匹配到 {_title} 在 {ws_name}表头中')
        else:  # 如果系统表头NOT 在品牌第二行
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
    ColorPrint.print("注意!!操作的数据列请确保格式完全一致，请排除空格引号等问题！\n" * 3, color='random')
    _val1 = input_selector('请分别输入【表名1 匹配列名1（需要大写字母）】')
    try:
        sheet_name1, column_1 = _val1[:2]
        print('载入表1，成功')
    except Exception as e:
        print("载入表1错误", e)
        exit(1)
    _val2 = input_selector('请分别输入【表名2 匹配列名2（需要大写字母）】')
    try:
        sheet_name2, column_2 = _val2[:2]
        print('载入表2，成功')
    except Exception as e:
        print("载入表2错误", e)
        exit(1)
    fill = PatternFill("solid", fgColor="1874CD")  # 蓝色样式标记-品牌
    fill2 = PatternFill("solid", fgColor="FF0000")  # 红色样式标记-TS
    ws = gap.wb[sheet_name1]
    wp = gap.wb[sheet_name2]
    t_data = set(ws[column_1][2:ws.max_row])
    _tdata = [cell.value for cell in t_data]
    p_data = set(wp[column_2][2:wp.max_row])
    _pdata = [cell.value for cell in p_data]
    # print(f'🚀正在标注  {sheet_name1}  与  {sheet_name2}  差异！')
    for sty, tid in track(enumerate(t_data, start=2)):
        # print(sty,tid.value)
        if tid.value not in _pdata:
            ws[f'{column_1}{sty}'].fill = fill2
    # print(f'🚀正在标注  {sheet_name2}  与  {sheet_name1}  差异！')
    for py, pid in track(enumerate(p_data, start=2)):
        if pid.value not in _tdata:
            wp[f'{column_2}{py}'].fill = fill


@auto_save
def set_gap_by_vlookup(gap: Gap):
    wb = gap.wb
    _value = input_selector("请输入【系统表名 系统表头->品牌数据表名】")
    ws_name_sys, ws_name_brand = _value[:2]
    print(ws_name_sys, ws_name_brand)
    ws_brand = wb[ws_name_brand]
    ws_system = wb[ws_name_sys]
    ws_gap = wb.create_sheet('gap')
    # 取最大列，最大行
    max_col = max(ws_brand.max_column, ws_system.max_column)
    max_row = max(ws_brand.max_row, ws_system.max_row)
    min_row = min(ws_brand.max_row, ws_system.max_row)
    sr = set_az()
    system_title: dict = {cell.value: _i for _i, cell in enumerate(ws_system[1], start=1)}  # 此处将字段作为键，索引作为值
    for i in track(range(1, max_col * 3, 3), description=f'匹配进度列'):  # 此处i代表新表中的每一列的开头一列
        # print(i)
        brand_col = i
        system_col = i + 1
        gap_col = i + 2
        _index = i // 3
        quote_col_name = sr[_index]  # 需要引入的列名
        # 当前列title,这里需要直接取brand字段
        title = ws_brand[f'{quote_col_name}2'].value
        for irow in range(1, min_row + 1):
            _brand = ws_brand[f'{quote_col_name}{irow}']
            _system = ws_system[f'{quote_col_name}{irow}']
            # 标注出3个cell对象
            gap_brand_cell = ws_gap.cell(row=irow, column=brand_col)
            gap_system_cell = ws_gap.cell(row=irow, column=system_col)
            gap_gap_cell = ws_gap.cell(row=irow, column=gap_col)
            # 注意不用区拿标题，所以当irow=1时，需要过滤
            F1 = gap_system_cell.coordinate
            G1 = gap_brand_cell.coordinate
            H1 = gap_gap_cell.coordinate
            letter_part = H1.rstrip('0123456789')
            use_letter = f'{letter_part}:{letter_part}'
            # 前两列需要引入的公式
            _col_name = f'{quote_col_name}{irow}'
            if irow == 1:  # 如果是首行
                quite_brand = f'={ws_brand.title}!{_col_name}&"-品牌"'  # 品牌表头
                quite_system = f'={ws_system.title}!{_col_name}&"-系统"'  # 系统表头
                quite_gap = f'={ws_system.title}!{_col_name}&"-GAP"'
                gap_brand_cell.fill = brand_fill
                gap_system_cell.fill = system_fill
                gap_gap_cell.fill = gap_fill
                gap_brand_cell.font = title_font
                gap_system_cell.font = title_font
                gap_gap_cell.font = title_font
            else:  # 非首行
                if irow == 2:
                    quite_gap = title
                    gap_gap_cell.value = quite_gap
                    gap_gap_cell.fill = nothing_fill  # 标注无需gap字段
                    gap_brand_cell.fill = nothing_fill
                    gap_system_cell.fill = nothing_fill
                    continue
                if title in system_title:  # 这里默认会检索key
                    _vindex = system_title[title]
                else:
                    _vindex = 'no'
                if _vindex == 'no':
                    quite_gap = None  # Gap
                    quite_system = None  # 系统
                    quite_brand = None  # Brand
                    gap_gap_cell.fill = nothing_fill  # 标注无需gap字段
                    gap_brand_cell.fill = nothing_fill
                    gap_system_cell.fill = nothing_fill
                else:
                    gap_brand_cell.font = gap_font
                    gap_system_cell.font = gap_font
                    gap_gap_cell.font = gap_font
                    if isinstance(_brand.value, (int, float)) and isinstance(_system.value, (int, float)):
                        # =IF(A1=0,IF(B1=0,0,IF(A1>B1,(B1-A1)/A1,(B1-A1)/B1)),IF(B1=0,0,IF(A1>B1,(B1-A1)/A1,(B1-A1)/B1)))
                        # gap
                        quite_gap = f'=IF({F1}=0,IF({G1}=0,0,IF({F1}>{G1},({G1}-{F1})/{F1},({G1}-{F1})/{G1})),IF({G1}=0,0,IF({F1}>{G1},({G1}-{F1})/{F1},({G1}-{F1})/{G1})))'
                        gap_gap_cell.number_format = numbers.FORMAT_PERCENTAGE_00  # 设置百分比格式？
                    else:
                        quite_gap = f'=EXACT({F1},{G1})'
                        if gap_system_cell.value != gap_brand_cell.value:
                            gap_gap_cell.fill = gap_fill_false  # 字符串不等，标记红色
                    # 系统
                    quite_system = f'=VLOOKUP(A{irow},{ws_system.title}!A:ZZ,{_vindex},0)'
                    quite_brand = f'={ws_brand.title}!{_col_name}'  # 品牌
            gap_brand_cell.value = quite_brand
            gap_system_cell.value = quite_system
            gap_gap_cell.value = quite_gap
    wrokgap = wb.copy_worksheet(ws_gap)
    wrokgap.title = 'wrokgap'
    wb.close()
    print('🎁Gap_Sheet生成成功🎁，等待保存')


get_row_data_comment.__name__ = '提取标题批注'
get_data_by_row_title.__name__ = '快速引用品牌列数据【在执行4前应该先执行这个】'
contrast_sheets.__name__ = '对比两列数据并标记'
set_gap_sheet.__name__ = '生成GapSheet[openpyxl]'
set_gap_by_vlookup.__name__ = '生成GapSheet[需要先手动打开在保存一次再使用]'

fundict = {
    '1': get_row_data_comment,
    '2': get_data_by_row_title,
    '3': contrast_sheets,
    '4': set_gap_by_vlookup,
    '5': set_gap_sheet,
    }


def function_list(obj: Gap):
    ColorPrint.print('    ', '=' * 10, '功能列表', '=' * 10, color='yellow')
    for i in fundict:
        print()
        ColorPrint.print("          ", i, fundict[i].__name__, color='yellow')
    print()
    ColorPrint.print('    ', '=' * 10, '功能列表', '=' * 10, color='yellow')
    _x = input_selector("选择功能：")
    x = _x[:1][0]
    if x not in fundict.keys():
        clear()
        print('输入有误！，请重新选择')
        return function_list()
    else:
        # try:
        return fundict[x](obj)
        # except Exception as e:
        #     Log.error(e)
        #     exit(1)


if __name__ == '__main__':

    path = file_selector()
    print(path)
    obj = Gap(path)
    function_list(obj)
