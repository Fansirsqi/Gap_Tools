import os
import glob
import platform
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
