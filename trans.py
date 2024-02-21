import os
import sys

import pandas as pd


def convert_encoding(folder_path, file_encoding='ansi', tarns_encoding='utf-8'):
    # 获取文件夹中所有文件的列表
    file_list = os.listdir(folder_path)
    # 遍历文件夹中的每个文件
    for file_name in file_list:
        if file_name.endswith('.csv'):
            # 构建完整的文件路径
            file_path = os.path.join(folder_path, file_name)
            # 读取原始编码的CSV文件
            df = pd.read_csv(file_path, encoding=file_encoding)
            
            save_path = file_path.replace(folder_path, f'{folder_path}{tarns_encoding}/')
            # print(save_path)
            os.makedirs(folder_path +  tarns_encoding, exist_ok=True)
            os.makedirs(folder_path + file_encoding, exist_ok=True)
            # 将数据保存为UTF-8编码的CSV文件（覆盖原文件）
            df.to_csv(save_path, index=False, encoding=tarns_encoding)
            # 移动文件到指定文件夹
            os.rename(file_path, file_path.replace(folder_path, f'{folder_path}{file_encoding}/'))
            print(f'文件 {file_name} 已转换为{tarns_encoding}编码并覆盖保存')


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print('Usage: python trans.py ./trans_dir/ <input_encoding utf-8> <output_encoding ansi>')
        exit(1)
    else:
        # 指定包含CSV文件的文件夹路径
        convert_encoding(sys.argv[1], sys.argv[2], sys.argv[3])
