import os
from sys import stdout

import pandas as pd
from dotenv import load_dotenv
from loguru import logger

load_dotenv(override=True, verbose=True)

IS_DEBUG = os.getenv('IS_DEBUG', 'false').lower() == 'true'
print('denug model', IS_DEBUG)
logger.remove()
logger.add(stdout, level='INFO', colorize=True, format='<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}')

logger.add('日报.log', encoding='utf-8', format='<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}')
_account_name = os.getenv('ACCOUNT_NAME')
"""账号名称"""
_s_date = os.getenv('START_DATE')
"""开始日期"""
_e_date = os.getenv('END_DATE')
"""结束日期"""
csv_data_folder = os.getenv('BASEFLOAD')
"""底表父文件夹"""


class Scrm:
    def __init__(self):
        self.class_name = ''
        self.account_name = _account_name
        if self.account_name is None or self.account_name == '':
            print('账号名称为空或者错误，程序退出')
            exit(1)
        self.start_date = _s_date
        self.end_date = _e_date
        self.csv_data_folder = csv_data_folder
        self.csvPerfix = 'scrm_dy_report_app_fxg_'
        """百库底表前缀"""
        self.export_folder = f'./export_{self.class_name}/{self.account_name}'
        """导出文件夹"""
        os.makedirs(self.export_folder, exist_ok=True)
        self.dfs = {
            'live_detail_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_detail_day.csv'),
            'live_list_details_traffic_time_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_list_details_traffic_time_day.csv'),
            'live_growth_conversion_funnel_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_growth_conversion_funnel_day.csv'),
            'live_list_details_grouping_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_list_details_grouping_day.csv'),
            'live_list_details_indicators_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_list_details_indicators_day.csv'),
            'live_analysis_live_details_all_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_analysis_live_details_all_day.csv'),
            'scrm_ocean_daily': pd.read_csv(f'{self.csv_data_folder}/scrm_ocean_daily.csv'),
            'live_detail_core_interaction_key_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_detail_core_interaction_key_day.csv'),
            'live_crowd_data_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_crowd_data_day.csv'),
            'live_detail_person_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_detail_person_day.csv'),
            'live_list_details_download_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}live_list_details_download_day.csv'),
            'fly_live_detail_first_prchase_day': pd.read_csv(f'{self.csv_data_folder}/{self.csvPerfix}fly_live_detail_first_prchase_day.csv'),
        }
        """csv data"""

        try:  # check csv file
            for csv in self.dfs.keys():
                if os.path.exists(csv):
                    print(f'{csv} 文件不存在')
                    exit(1)
        except Exception as e:
            print(e, '校验 csv data 错误')
