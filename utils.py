from datetime import datetime, timedelta
import json
import os
from time import sleep
import glob
import zipfile
from tqdm import tqdm
import platform
import subprocess
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import time
import os
import random

class ColorPrint:
    COLORS = {
        'black': '30',
        'red': '31',
        'green': '32',
        'yellow': '33',
        'blue': '34',
        'magenta': '35',
        'cyan': '36',
        'white': '37'
    }

    @staticmethod
    def print(*args, color=None, font_weight="bold", end='\n'):
        STYLES = {
            'normal': '0',
            'bold': '1',
            'italic': '3',
            'underline': '4',
            'strikethrough': '9'
        }
        if color and color.lower() == 'random':
            color_code = str(random.randint(30, 36))
        elif color and color.lower() in ColorPrint.COLORS:
            color_code = ColorPrint.COLORS[color.lower()]
        else:
            color_code = None
        if color_code is not None:
            text = '\033[{};{}m{}\033[0m'.format(color_code, STYLES[font_weight], ' '.join(map(str, args)))
        else:
            text = ' '.join(map(str, args))
        print(text, end=end)

def trans_str(_str):
        result = str(_str).replace('(','').replace(')','').replace(',','').replace("'",'').replace('"','').replace(r'\n', '\n').replace(r'\t', '\t').replace(r'\r', '\r')
        sleep(0.02)
        return result

class Log:
    """日志打印模块，包含了一个输入获取模块，保持控制台字体一致

    Returns:
        (Any): _description_
    """
    font_yellow = '\033[1;33m'# 黄色
    font_red = '\033[1;31m' # 红色
    font_blue = '\033[1;34m' # 蓝色
    font_gray = '\033[1;30m' # 灰色
    font_green = '\033[1;32m' # 绿色
    font_purple = '\033[1;35m' # 紫色
    font_cyan = '\033[1;36m' # 青色
    font_white = '\033[1;37m' # 白色
    
    bg_red = '\033[41m' # 红色 白字
    bg_green = '\033[42m' # 绿色 深灰字
    bg_yellow = '\033[43m' # 黄色 灰字
    bg_blue = '\033[44m' # 蓝色 白字
    bg_purple = '\033[45m' #紫色 白字
    bg_cyan = '\033[46m' # 青色 深灰字
    bg_gray = '\033[47m' # 灰色 深灰字
    reset = '\033[0m'
    
    @staticmethod
    def warning(*context):
        """打印黄色警告"""
        context = trans_str(context)
        print(f'{Log.font_yellow}⚠️  [WARNING] |\n{context} {Log.reset}') 
    @staticmethod
    def error(*context):
        """打印红色错误警告"""
        context = trans_str(context)
        print(f'{Log.font_red}🔴 [ERROR]   |\n{context} {Log.reset}')
    @staticmethod
    def info(*context):
        """打印蓝色信息"""
        context = trans_str(context)
        print(f'{Log.font_blue}🔵 [INFO]    |\n{context} {Log.reset}')
    @staticmethod
    def success(*context):
        """打印绿色信息"""
        context = trans_str(context)
        print(f'{Log.font_green}🟢 [SUCCESS] |\n{context} {Log.reset}')
    @staticmethod
    def debug(*context):
        """打印灰色信息"""
        context = trans_str(context)
        print(f'{Log.font_gray}⚙️  [DEBUG]   |\n{context} {Log.reset}')
    @staticmethod
    def input(context):
        """获取输入信息"""
        data = input(f'{Log.font_white}✍️  [INPUT]   |\n{context} {Log.reset}')
        return data

class DotDict(dict):
    """将字典数据转换成类的形式，数据可以通过.xx的形式访问

    Args:
        dict (_type_): _description_
    """
    def __init__(self, *args, **kwargs):
        super(DotDict, self).__init__(*args, **kwargs)

    def __getattr__(self, key):
        value = self[key]
        if isinstance(value, dict):
            value = DotDict(value)
        return value

def list_folder_(filter=['.venv','__pycache__','.vscode','.git'],prt=True)-> list:
    """列出当前项目文件夹

    Args:
        filter (list, optional): 过滤的文件夹. Defaults to ['.venv','__pycache__','.vscode','.git'].
        prt (bool, optional): 是否开启打印. Defaults to True.

    Returns:
        list: _description_
    """
    folder_list = [name for name in os.listdir('.') if os.path.isdir(name) and name not in filter]
    if len(folder_list)!=0:
        if prt is True:
            count = 1
            for i in folder_list:
                ColorPrint.print(f'     {count}.{i}',color='blue')
                print()
                count += 1
    return folder_list

def list_files(filter='*.xlsx', prt=True, prefix='~$'):
    """列出文件"""
    # 获取当前工作目录
    current_dir = os.getcwd()
    # 构建要匹配的文件路径模式
    file_pattern = os.path.join(current_dir, filter)
    # 使用glob模块获取匹配的文件列表
    files = glob.glob(file_pattern)
    if prefix:
        # 过滤掉文件名前缀
        files = [file for file in files if not os.path.basename(file).startswith(prefix)]
    if prt is True:
        # 打印文件列表
        count = 1
        ColorPrint.print(f'    ','='*10,'文件列表','='*10, color='magenta')
        print()
        for file in files:
            file_name = os.path.basename(file)
            ColorPrint.print(f'     {count}.{file_name}', color='magenta')
            print()
            count += 1
    return files

def get_file_list(dir):
    search_path = os.path.join(dir, '*')
    # 构建要匹配的文件路径模式
    files = glob.glob(search_path)
    return files

def clear():
    """清屏"""
    # 返回系统平台/OS的名称，如Linux，Windows，Java，Darwin
    system = platform.system()
    if (system == u'Windows'):
        os.system('cls')
    else:
        os.system('clear')

def input_selector(text='请输选择：'):
    """获取输入信息"""
    try:
        user_input = input(f'\033[1;33m{text}\033[0m')
        if user_input:
            return user_input.split(' ')
        else:
            return []
    except KeyboardInterrupt as e:
        ColorPrint.print(f'\n手动终止！{e}', color='red')
        exit(0)

# 调用函数并进行解包操作
# values = input_selector()
# x, y = values[:2]  # 只取前两个值，多余的值会被忽略
# print(x, y)

def file_selector():
    file_list = list_files()
    _select = input_selector()
    select = _select[:1][0]
    # print(select)
    try:
        current = file_list[int(select)-1]
        return (current)
    except Exception as e1:
        ColorPrint.print(f'输入有误！{e1},请重试！')
        return file_selector()
    
def set_az():
    """
    生成一个列表[A-ZZ]
    @return: [A-ZZ]
    """
    sr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    sl = [char for char in sr]
    sl += [char1 + char2 for char1 in sr for char2 in sr]
    # print(len(sl))
    return sl
