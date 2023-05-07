# _*_coding:utf-8_*_

# @PROJECT : DY_SCRM_TEST_PROJECT
# @Time : 2022/12/6 22:56
# @Author : Byseven
# @File : gap_tool.py
# @SoftWare:

import os
import re

import openpyxl
from openpyxl.styles import PatternFill

fill = PatternFill("solid", fgColor="1874CD")
fill2 = PatternFill("solid", fgColor="FF0000")


def get_filelist(path: str) -> list:
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


def get_name_info(file_path: str) -> list:
    """获取文件名，文件名后缀
    :param file_path: 文件路径
    :return: [文件名，后缀]
    """
    if os.path.isfile(file_path):
        file_name = file_path.split('\\').pop()
        file_name_last = file_name.split('.').pop()
    else:
        return ['！！ERROR！！传入的路径并不是文件路径！！']
    return [file_name, file_name_last]


def cancel_merge(excel_path):
    """
    取消全表单元格合并居中
    :param excel_path: 路径
    :return:
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    print('开始处理合并单元格')
    for s in wb:
        ws = wb[s.title]
        # 列出所有的合并单元格的索引信息
        merge_list = [str(item) for item in ws.merged_cells.ranges]
        # 批量取消合并,并且将原来合并的区域填充上原来的数据(因为本来也许就是空的,也可以不填)
        for item_merge in merge_list:
            # 左上,右下角值坐标
            top_left, bot_right = item_merge.split(':')
            top_left_col, top_left_row = ws[top_left].column, ws[top_left].row
            bot_right_col, bot_right_row = ws[bot_right].column, ws[
                bot_right].row
            # 记下该合并单元格的值
            cell_value = ws[top_left].value
            # 取消合并单元格
            ws.unmerge_cells(item_merge)
            # 批量给子单元格赋值
            # 遍历列
            for col_idx in range(top_left_col, bot_right_col + 1):
                # 遍历行
                for row_idx in range(top_left_row + 1, bot_right_row + 1):
                    ws[f"{chr(col_idx + 64)}{row_idx}"] = cell_value
        # 记得保存
        wb.save(excel_path)
        print(f'{s.title}合并单元格处理完成')
    print('已处理完成本表所有合并单元格')


def contrast_sheets(excel_path: str, ts_sheet_name: str, ts_column: str, ts_max: int, p_sheet_name: str, p_column: str,
                    pin_max: int):
    """
    ！！请确保两列数据格式的一致性！！
    对比两个sheet中的B列【原本用于匹配千川视频id】，将两者不包含的分别标注
    :param excel_path: 表格路径
    :param ts_sheet_name: tssheet名称
    :param ts_column: ts列
    :param ts_max: ts最大行数
    :param p_sheet_name: 品牌sheet名称
    :param p_column: 品牌列
    :param pin_max: 品牌最大行数
    :return:
    """
    fill = PatternFill("solid", fgColor="1874CD")  # 蓝色样式标记-品牌
    fill2 = PatternFill("solid", fgColor="FF0000")  # 红色样式标记-TS
    wb = openpyxl.load_workbook(excel_path)
    print('读取表格成功！请中途勿关闭程序！\n--会导致表格损坏！！-')
    ws = wb[ts_sheet_name]
    print(f'读取sheet:{ts_sheet_name}成功')
    tv_ids = []
    for i in range(2, ts_max + 1):
        v_id = ws[f'{ts_column}{i}'].value
        print(f'读取TS第{i}行 - |{v_id}|')
        tv_ids.append(v_id)

    wp = wb[p_sheet_name]
    print(f'读取sheet:{p_sheet_name}成功！')
    pv_ids = []
    for j in range(2, pin_max + 1):
        vp_id = wp[f'{p_column}{j}'].value
        print(f'读取PIN第{j}行 - |{vp_id}|')
        pv_ids.append(vp_id)

    print(f'正在标注  {p_sheet_name}  与  {ts_sheet_name}  差异！')
    for pid in pv_ids:  # 遍历品牌数据
        if pid not in tv_ids:  # 如果值不在TS列表中
            ty = pv_ids.index(pid) + 2
            wp[f'{p_column}{ty}'].fill = fill

    print(f'正在标注  {ts_sheet_name}  与  {p_sheet_name}  差异！')
    for tid in tv_ids:
        if tid not in pv_ids:
            sty = tv_ids.index(tid) + 2
            ws[f'{ts_column}{sty}'].fill = fill2

    wb.save(excel_path)
    print('处理完成')
    return tv_ids, pv_ids


def get_all_title(excel_path, table_title, l):
    file_name = get_name_info(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    sr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    print(f'开始处理:{file_name[0]}')
    for s in wb:
        sheet_name = s.title
        ws = wb[sheet_name]
        for i in range(len(sr)):  # 遍历A~Z
            pti = ws[f'{sr[i]}1'].value
            if i != 'A':
                pti_last = ws[f'{sr[i - 1]}1'].value
                if pti == pti_last:
                    print(pti, pti_last, f"{file_name[0]}\t{sr[i]}列跳出循环")
                    break
            l.append(sr[i])
            table_title.append(str(pti) + '\t' + str(excel_path) + '\t' + str(sr[i]))
        for j in range(len(sr)):  # j 控制第一个字母
            for k in range(len(sr)):  # k 控制第二个字母
                pti = ws[f'{sr[j]}{sr[k]}1'].value
                if k != 'A':
                    pti_last = ws[f'{sr[j]}{sr[k - 1]}1'].value
                    if pti == pti_last:
                        print(pti, pti_last, f"{file_name[0]}\t{sr[j]}{sr[k]}列跳出循环")
                        break
                l.append(sr[j] + sr[k])
                table_title.append(str(pti) + '\t' + str(excel_path) + '\t' + str(sr[j]) + str(sr[k]))
    print('处理完成！')


def exact_title(ws, x: int):
    """
    读取A-ZZ列，x行的数据，返回数据列表，和数据列表对应列号
    :param x:
    :param ws: sheet
    :return:[]
    """
    key_list = []  # 记录字段
    key_list_index = []  # 记录字段对应的列
    sr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    for i in sr:
        key = ws[f'{i}{x}'].value
        if key is None or key == '':
            print(f'{ws}表 {i}列 值 {x} 行为空 跳出')
            break
        print(f'{ws}表 {i}列 {x}行 值{key}')
        key_list_index.append(i)
        key_list.append(key)
    for j in sr:
        for k in sr:
            key = ws[f'{j}{k}{x}'].value
            if key is None or key == '':
                print(f'{ws}表 {j}{k}列 {x}行 值为空 跳出')
                break
            print(f'{ws}表 {j}{k}列 {x}行 值{key}')
            key_list_index.append(j + k)
            key_list.append(key)
    return [key_list, key_list_index]


def ext(excel_path, ts_name, p_name, np_name):
    """
    处理表头
    :param excel_path:
    :return:
    :param np_name:
    :param p_name:
    :param ts_name:
    """
    wb = openpyxl.load_workbook(excel_path)
    print('开始处理并标记')
    ws = wb[p_name]  # 读取品牌表
    pin_data = exact_title(ws, 1)
    wst = wb[ts_name]  # 读取TS表
    ts_data = exact_title(wst, 1)  # ts_sheet 第二行，数值，对应索引
    to_do_list = []  # 创建一个需要工作的任务列表
    _to_do = []  # 这是需要放在函数里的值，也建表与上面好对应
    for pi in pin_data[0]:  # 这是存放第x行数据的列表
        if pi in ts_data[0]:  # 这是存放第x行数据的列表
            pin_column = pin_data[0].index(pi)  # 在品牌数据表的下标
            ts_column = ts_data[0].index(pi)  # 在TS数据表的下标
            ts_lie = ts_data[1][ts_column]  # ts列
            p_lie = pin_data[1][pin_column]  # 品牌列

            to_do_list.append(ts_lie)  # 这是目标列
            _to_do.append(p_lie)  # 对应目标列的函数内容，部分
            print(ts_lie, pi, p_lie)
    print(to_do_list)
    print(_to_do)
    ows = wb[np_name]  # 读取转换表，需要手动创建，否则报错
    print(f'开始写入{p_name}')
    for c, z in zip(to_do_list, _to_do):
        print(f'写入品牌{z}列至{c}列')
        ows[f'{c}1'] = f'={p_name}!{z}1'  # 函数引用品牌表头
        ows[f'{c}2'].value = z
        ows[f'{c}3'] = f'={p_name}!{z}2'  # 函数
    wb.save(excel_path)
    print('写入完成！')
    exit()


def heng(excel_path, ksy):
    """
    返回第一行名称为ksy的列号
    :param excel_path:
    :return:
    """
    file_name = get_name_info(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    sr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    print(f'开始处理:{file_name[0]}中寻找存在{ksy}的列号')
    for s in wb:
        sheet_name = s.title
        ws = wb[sheet_name]
        for i in range(len(sr)):  # 遍历A~Z
            pti = ws[f'{sr[i]}1'].value
            if i != 'A':
                pti_last = ws[f'{sr[i - 1]}1'].value
                if pti == pti_last:
                    print(pti, pti_last, f"{file_name[0]}\t{sr[i]}列跳出循环")
                    break
            # l.append(sr[i])
            if pti == ksy:
                print(f'存在{ksy}的列{sr[i]}')
                return sr[i]
            # table_title.append(str(pti) + '\t' + str(excel_path) + '\t' + str(sr[i]))
        for j in range(len(sr)):  # j 控制第一个字母
            for k in range(len(sr)):  # k 控制第二个字母
                pti = ws[f'{sr[j]}{sr[k]}1'].value
                if k != 'A':
                    pti_last = ws[f'{sr[j]}{sr[k - 1]}1'].value
                    if pti == pti_last:
                        print(pti, pti_last, f"{file_name[0]}\t{sr[j]}{sr[k]}列跳出循环")
                        break
                # l.append(sr[j] + sr[k])
                if pti == ksy:
                    print(f'存在{ksy}的列{sr[j]}{sr[k]}')
                    return sr[j] + sr[k]
                # table_title.append(str(pti) + '\t' + str(excel_path) + '\t' + str(sr[j]) + str(sr[k]))
    print(f'寻列{ksy}完成！')


def vertical_read(excel_path, sheet_name, lie, max_read_line):
    """
    从第一行往下读取数据
    :param sheet_name: sheet名
    :param max_read_line: 读取最大行数
    :param excel_path:
    :param ksy: 需过滤的字符串
    :param lie: 列号
    :return:
    """
    ls: list = []
    file_name = get_name_info(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[sheet_name]
    print('读取成功')
    nws = wb['new']
    for i in range(2, max_read_line):
        key = ws[f'{lie}{i}'].value
        ls.append(key)
    print('已加载完毕所有类目')
    for i in range(2, max_read_line):
        key = ws[f'{lie}{i}'].value
        print(f'{key},{ls.index(key) + 2}')


def to_dict_table(excel_path):
    """
    处理数据字典【日报】
    :param excel_path:
    :return:
    """
    # cancel_merge(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    sr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    ts_f = []
    gk = []
    print('开始处理')
    ws = wb['罗盘爬数对应点位']
    for i in range(1, 73):
        keys = ws[f'Q{i}'].value
        if keys == '原始值':
            gksys = ws[f'G{i}'].value
            gksys = re.sub("[\u4e00-\u9fa5\0-9\,\。]", "", gksys)  # 提取英文
            ts_filed = ws[f'H{i}'].value
            ts_f.append(ts_filed)  # 导出展示字段
            gk.append(gksys)  # 字典爬数
            print(gksys, ts_filed)
    return [ts_f, gk]


if __name__ == '__main__':
    t, p = contrast_sheets(
        r'D:\DY_SCRM_TEST_PROJECT\DEV_CODE\chatGPT_tools\直播明细_GAp.xlsx',
        '（直播明细_0(2(1）品牌维度_销售额',
        "A",
        494,
        '（直播明细_0）品牌维度_销售额',
        "A",
        315)
