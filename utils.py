import glob
import os
import platform
from config import configs


def load_configs():
    """加载配置文件"""
    return configs


def cprint(*args, text_color='white', bg_color='black', is_bold='0', **kwargs):
    text_colors = {'black': '30', 'red': '31', 'green': '32', 'yellow': '33', 'blue': '34', 'magenta': '35', 'cyan': '36', 'white': '37'}
    bg_colors = {'black': '40', 'red': '41', 'green': '42', 'yellow': '43', 'blue': '44', 'magenta': '45', 'cyan': '46', 'white': '47', 'black1': '48'}

    text_color_code = text_colors.get(text_color, '37')  # default to white if color not found
    if text_color_code == '37':
        bg_color_code = bg_colors.get(bg_color, '40')  # default to black if color not found
    else:
        bg_color_code = '40'

    colored_args = [f'\033[{is_bold};{text_color_code};{bg_color_code}m {arg} \033[0m' for arg in args]
    print(*colored_args, **kwargs)


# 调用示例
# cprint('Hello, World!', text_color='red', bg_color='green', is_bold='1')


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


def list_folder_(filter=['.venv', '__pycache__', '.vscode', '.git', '.idea', 'export_人群', 'export_日报'], is_prt=True) -> list:
    """列出当前项目文件夹
    Args:
        filter (list, optional): 过滤的文件夹. Defaults to ['.venv','__pycache__','.vscode','.git'].
        prt (bool, optional): 是否开启打印. Defaults to True.
    Returns:
        list: _description_
    """
    folder_list = [name for name in os.listdir('.') if os.path.isdir(name) and name not in filter]
    if len(folder_list) != 0:
        if is_prt is True:
            count = 1
            cprint('=' * 30, '文件夹列表', '=' * 30)
            for i in folder_list:
                cprint(f'     {count}.{i}')
                count += 1
            cprint('=' * 68)
    return folder_list


def list_files(filter: str = '*.xlsx', is_prt: bool = True, prefix: str = '~$', include: list = [], path: str = ''):
    """
    列出指定路径下匹配特定文件扩展名的文件，并可进行过滤。

    :param filter: 文件扩展名过滤器，默认为 '*.xlsx'
    :param prt: 是否打印文件列表，默认为 True
    :param prefix: 要过滤掉的文件名前缀，默认为 '~$'
    :param include: 包含文件名的关键字列表，默认为 []
    :param path: 指定搜索文件的路径，默认为 ''，即当前工作目录
    :return: 文件列表
    """

    # 设置当前目录
    current_dir = path if path else os.getcwd()

    # 构建文件路径模式
    file_pattern = os.path.join(current_dir, filter)

    # 获取匹配的文件列表
    files = glob.glob(file_pattern)

    # 过滤文件名前缀
    if prefix:
        files = [file for file in files if not os.path.basename(file).startswith(prefix)]

    # 过滤包含的关键字
    if include:
        files = [file for file in files if any(i in file for i in include)]

    # 打印文件列表
    if is_prt:
        count = 1
        cprint('==============================文件列表==============================', text_color='green')
        for file in files:
            file_name = os.path.basename(file)
            cprint(f'{count}. {file_name}', text_color='yellow')
            count += 1
        cprint('====================================================================', text_color='green')

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
    if system == 'Windows':
        os.system('cls')
    else:
        os.system('clear')


def input_selector(text='请输选择：') -> str:
    """获取输入信息"""
    try:
        user_input = input(f'\033[1;33m{text}\033[0m')
        if user_input:
            return user_input.split(' ')
        else:
            return []
    except KeyboardInterrupt as e:
        cprint(f'\n手动终止！{e}')
        exit(0)


# 调用函数并进行解包操作
# values = input_selector()
# x, y = values[:2]  # 只取前两个值，多余的值会被忽略
# cprint(x, y)


def file_selector():
    file_list = list_files()
    _select = input_selector()
    select = _select[:1][0]
    # cprint(select)
    try:
        current = file_list[int(select) - 1]
        return current
    except Exception as e1:
        cprint(f'输入有误！{e1},请重试！')
        return file_selector()


def set_az():
    """
    生成一个列表[A-ZZ]
    @return: [A-ZZ]
    """
    sr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    sl = [char for char in sr]
    sl += [char1 + char2 for char1 in sr for char2 in sr]
    # cprint(len(sl))
    return sl


if __name__ == '__main__':
    ...
    files: list = list_files(path='.\欧莱雅染发甄选', include=['自营', '构成', '销售概览'])
    # cprint(files)
