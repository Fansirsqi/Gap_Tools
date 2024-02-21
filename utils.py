import glob
import os
import platform


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


def list_folder_(filter=['.venv', '__pycache__', '.vscode', '.git', '.idea', 'export_人群', 'export_日报'], prt=True) -> list:
    """列出当前项目文件夹

    Args:
        filter (list, optional): 过滤的文件夹. Defaults to ['.venv','__pycache__','.vscode','.git'].
        prt (bool, optional): 是否开启打印. Defaults to True.

    Returns:
        list: _description_
    """
    folder_list = [name for name in os.listdir('.') if os.path.isdir(name) and name not in filter]
    if len(folder_list) != 0:
        if prt is True:
            count = 1
            print('=' * 30, '文件夹列表', '=' * 30)
            for i in folder_list:
                print(f'     {count}.{i}')
                count += 1
            print('=' * 68)
    return folder_list


def list_files(filter='*.xlsx', prt=True, prefix='~$', path=None):
    """列出文件列表

    Args:
        filter (str, optional): _description_. Defaults to "*.xlsx". 文件类型
        prt (bool, optional): _description_. Defaults to True. 是否打印
        prefix (str, optional): _description_. Defaults to "~$". 过滤前缀
        path (str, optional): _description_. Defaults to None. 输入的路径,如果不输入则默认获取脚本自身所在目录文件,如果输入则获取目录下的内容

    Returns:
        _type_: _description_
    """
    if path is None:
        # 获取当前工作目录
        current_dir = os.getcwd()
    else:
        current_dir = path
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
        print('=' * 30, '文件列表', '=' * 30)
        for file in files:
            file_name = os.path.basename(file)
            print(f' {count}. {file_name}')
            count += 1
        print('=' * 68)
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
        print(f'\n手动终止！{e}')
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
        current = file_list[int(select) - 1]
        return current
    except Exception as e1:
        print(f'输入有误！{e1},请重试！')
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


if __name__ == '__main__':
    files: list = list_files(path=r'C:\Users\HPLaptop-14S\Documents\Gap文档\1', filter='*.csv')
    print(files)
