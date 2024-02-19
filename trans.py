import os
import pandas as pd

# 指定包含CSV文件的文件夹路径
folder_path = './csv/'

# 获取文件夹中所有文件的列表
file_list = os.listdir(folder_path)

# 遍历文件夹中的每个文件
for file_name in file_list:
    if file_name.endswith('.csv'):
        # 构建完整的文件路径
        file_path = os.path.join(folder_path, file_name)
        
        # 读取原始编码的CSV文件
        df = pd.read_csv(file_path, encoding='ansi')
        
        # 将数据保存为UTF-8编码的CSV文件（覆盖原文件）
        df.to_csv(file_path, index=False, encoding='utf-8')

        print(f'文件 {file_name} 已转换为UTF-8编码并覆盖保存')
