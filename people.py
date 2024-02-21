import os
from sys import stdout

import pandas as pd
from dotenv import load_dotenv
from loguru import logger
from exoprt import down, rq_data, hosts
from functions import set_gap
from openpyxl import load_workbook

load_dotenv(override=True, verbose=True)

IS_DEBUG = os.getenv('IS_DEBUG', 'false').lower() == 'true'
print('debug model', IS_DEBUG)
logger.remove()
logger.add(stdout, level='INFO', colorize=True, format='<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}')

logger.add('流量.log', encoding='utf-8', format='<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}')

_account_name = os.getenv('ACCOUNT_NAME')
"""账号名称"""
_s_date = os.getenv('START_DATE')
"""开始日期"""
_e_date = os.getenv('END_DATE')
"""结束日期"""
BASEFLOAD = os.getenv('BASEFLOAD')


class People:
    class_name = '店播人群'
    current_time = str(pd.Timestamp.now())
    logger.debug('-' * 30 + '| ' + current_time + ' |' + '-' * 30)
    account_name = _account_name
    if account_name is None or account_name == '':
        logger.error('账号名称为空或者错误，程序退出')
        exit()
    start_date = _s_date
    end_date = _e_date

    sheet_date_type = 'date'

    baseFload = BASEFLOAD
    """底表父文件夹"""
    csvPerfix = 'scrm_dy_report_app_fxg_'
    """百库底表前缀"""
    export_folder = f'./export_{class_name}/{account_name}'
    """导出文件夹"""
    os.makedirs(export_folder, exist_ok=True)

    dfs = {
        'live_number_people_covered_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_number_people_covered_day.csv'),
        'live_growth_conversion_funnel_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_growth_conversion_funnel_day.csv'),
        'live_detail_person_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_detail_person_day.csv'),
        'live_crowd_data_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_crowd_data_day.csv'),
        'live_detail_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_detail_day.csv'),
        'live_list_details_grouping_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_list_details_grouping_day.csv'),
        'fly_live_detail_first_prchase_day': pd.read_csv(f'{baseFload}/{csvPerfix}fly_live_detail_first_prchase_day.csv'),
    }
    """csv data"""

    try:  # check csv file
        for csv in dfs.keys():
            if os.path.exists(csv):
                print(f'{csv} 文件不存在')
                exit(1)
    except Exception as e:
        print(e)

    logger.debug('init data import ...')
    for df_key, df in dfs.items():
        logger.debug(f'导入 {df_key}')
        df[sheet_date_type] = pd.to_datetime(df[sheet_date_type])

        if df_key == 'live_analysis_live_details_all_day':
            dfs[df_key] = df[(df['author_nick_name'] == account_name) & (df['biz_date'].between(start_date, end_date))].sort_values(by=['biz_date'])
        elif df_key == 'scrm_ocean_daily':
            dfs[df_key] = df[(df['account_name'] == account_name) & (df['marketing_goal'] == 'LIVE_PROM_GOODS') & (df[sheet_date_type].between(start_date, end_date))].sort_values(by=[sheet_date_type])
        elif df_key == 'fly_live_detail_first_prchase_day':  # 此处关系到粉丝成交GMV相关内容，必须手动补充
            if '圣罗兰' in account_name:
                store_name = 'YSL圣罗兰美妆官方旗舰店'
            elif '兰蔻' in account_name:
                store_name = '兰蔻LANCOME官方旗舰店'
            elif '小美盒' in account_name:
                store_name = '欧莱雅集团小美盒官方旗舰店'
            elif '科颜氏' in account_name:
                store_name = "科颜氏KIEHL'S官方旗舰店"
            elif '碧欧泉' in account_name:
                store_name = '碧欧泉BIOTHERM官方旗舰店'
            elif 'HR赫莲' in account_name:
                store_name = 'HR赫莲娜官方旗舰店'
            elif '植村秀' in account_name:
                store_name = '植村秀shu uemura官方旗舰店'
            elif 'Vichy薇姿' in account_name:
                store_name = 'Vichy薇姿官方旗舰店'
            elif '欧莱雅' in account_name:
                store_name = '欧莱雅官方旗舰店'

            dfs[df_key] = df[
                (df['store_name'] == store_name) & (df['date'].between(start_date, end_date))
                # & (df["account_type"] == "渠道账号")
            ].sort_values(by='date')
        elif df_key == 'live_list_details_traffic_time_day':  # 此处需要过滤部分条件
            df[sheet_date_type] = pd.to_datetime(df[sheet_date_type])
            dfs[df_key] = df[(df['account_name'] == account_name) & (df['flowChannel'] != '整体') & (df['flowChannel'] != '全域推广') & (df['channelName'] != '整体') & (df[sheet_date_type].between(start_date, end_date))].sort_values(by=sheet_date_type)
        else:
            df[sheet_date_type] = pd.to_datetime(df[sheet_date_type])
            dfs[df_key] = df[(df['account_name'] == account_name) & (df[sheet_date_type].between(start_date, end_date))].sort_values(by=sheet_date_type)

    logger.debug('data import success !')

    观看人数 = 'imageCoverageCount'
    成交人数 = 'watchRate'
    成交人群 = 'gmv'
    直播间观看人数 = 'audienceStudio'
    成交人群分析_新老客维度_首购人数 = 'tranAnalysis_customers_firstBuyCount'
    成交人群分析_新老客维度_首购人数_粉丝占比 = 'tranAnalysis_customers_firstBuyCount_fansRate'
    成交人群分析_新老客维度_复购人数 = 'tranAnalysis_customers_secBuyCount'
    成交人群分析_新老客维度_复购人数_粉丝占比 = 'tranAnalysis_customers_secBuyCount_fansRate'
    成交人群分析_粉丝维度_粉丝人数 = 'tranAnalysis_fans_fansCount'
    成交人群分析_粉丝维度_粉丝人数_新增粉丝占比 = 'tranAnalysis_fans_fansCount_newFansRate'
    成交人群分析_互动维度_有互动人数 = 'tranAnalysis_interaction_intrCount'
    粉丝占比 = 'fans'
    直播间成交金额 = 'roomTransactionAmount'
    粉丝成单占比 = 'old_fans_pay_ucnt_ratio'
    累计观看人数 = 'cumulativeAudience'

    def export_import_csv(dfs=dfs, export_folder=export_folder):
        """导出csv文件:
        Args:
            dfs (_type_, optional): 被类处理后的dfs字典. Defaults to dfs.
            export_folder (_type_, optional): 指定的导出文件夹. Defaults to export_folder.
        """
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
        for df_key, df in dfs.items():
            df.to_csv(f'{export_folder}/{df_key}.csv', index=False)

    def get_指定观看人数(df=dfs['live_number_people_covered_day'], sum_filed=观看人数, date_type=sheet_date_type, turnover=None, fans=None, interaction=None, customers=None, watch=None):
        """参数说明
        df: 数据框，默认为 People_df
        sum_filed: 需要汇总的列名，默认为 观看人数
        account_name: 账号名称 从类开头获取
        date_type: 日期类型，默认为 sheet_date_type
        turnover: 成交，默认为 None
        fans: 粉丝，默认为 None
        interaction: 互动，默认为 None
        customers: 新老客，默认为 None
        watch: 观看，默认为 None

        函数功能: 根据传入的参数筛选数据，并对指定列进行汇总
        """
        logger.debug(f'\n{df}')
        ndf = df[
            (df['turnover'] == turnover)  # 成交
            & (df['fans'] == fans)  # 粉丝
            & (df['interaction'] == interaction)  # 互动
            & (df['customers'] == customers)  # 新老客
            & (df['watch'] == watch)  # 观看
        ]
        ndf1 = ndf.groupby([date_type])[sum_filed].sum()  # .sort_index().to_csv("demo.csv")
        return ndf1

    def get_成交人数(df=dfs['live_growth_conversion_funnel_day'], sum_filed=成交人数, date_type=sheet_date_type):
        logger.debug(f'\n{df}')
        ndf = df.groupby([date_type])[sum_filed].sum()
        return ndf

    def get_观看转化率(df1=dfs['live_detail_person_day'], df2=dfs['live_crowd_data_day'], sum_filed1=成交人群, sum_filed2=直播间观看人数, date_type=sheet_date_type):
        logger.debug(f'\n{df1}')
        logger.debug(f'\n{df2}')
        ndf1 = df1.groupby([date_type])[sum_filed1].sum()
        ndf2 = df2.groupby([date_type])[sum_filed2].sum()
        return (ndf1 / ndf2).apply(lambda x: format(x, '.9f'))

    def get_首购人数(df=dfs['live_detail_person_day'], sum_filed=成交人群分析_新老客维度_首购人数, date_type=sheet_date_type):
        logger.debug(f'\n{df}')
        return df.groupby([date_type])[sum_filed].sum()

    def get_首购粉丝数(df=dfs['live_detail_person_day'], sum_filed=[成交人群分析_新老客维度_首购人数, 成交人群分析_新老客维度_首购人数_粉丝占比], date_type=sheet_date_type):
        logger.debug(f'\n{df}')
        ndf1 = df[[sum_filed[0], 'id', date_type]]
        ndf2 = df[[sum_filed[1], 'id']]
        result = pd.merge(ndf1, ndf2, on='id')
        result['首购粉丝'] = result[sum_filed[0]] * result[sum_filed[1]]
        result = result.groupby(date_type)['首购粉丝'].sum()
        return result

    def get_复购粉丝数(df=dfs['live_detail_person_day'], sum_filed=[成交人群分析_新老客维度_复购人数, 成交人群分析_新老客维度_复购人数_粉丝占比], date_type=sheet_date_type):
        ndf1 = df.loc[:, [sum_filed[0], date_type, sum_filed[1], 'id']]
        ndf1.loc[:, '复购粉丝'] = ndf1[sum_filed[0]] * ndf1[sum_filed[1]]
        # result["复购粉丝"] = result[sum_filed1] * result[sum_filed2]
        return ndf1.groupby(date_type)['复购粉丝'].sum().apply(lambda x: format(x, '.0f')).astype('int64')

    def get_复购人数(df=dfs['live_detail_person_day'], sum_filed=成交人群分析_新老客维度_复购人数, date_type=sheet_date_type):
        return df.groupby([date_type])[sum_filed].sum()

    def get_粉丝成交人数(df=dfs['live_detail_person_day'], sum_filed=成交人群分析_粉丝维度_粉丝人数, date_type=sheet_date_type):
        return df.groupby([date_type])[sum_filed].sum()

    def get_粉丝观看人数(df=dfs['live_crowd_data_day'], sum_filed1=直播间观看人数, sum_filed2=粉丝占比, date_type=sheet_date_type):
        ndf1 = df[[sum_filed1, 'id', date_type]]
        ndf2 = df[[sum_filed2, 'id']]
        result = pd.merge(ndf1, ndf2, on='id')
        result['粉丝观看人数'] = result[sum_filed1] * result[sum_filed2] * 0.01
        result = result.groupby(date_type)['粉丝观看人数'].sum().apply(lambda x: format(x, '.2f'))
        return result.astype('float64')

    def get_粉丝成交GMV(df1=dfs['live_detail_day'], df2=dfs['fly_live_detail_first_prchase_day'], sum_filed1=直播间成交金额, sum_filed2=粉丝成单占比):
        ndf1 = df1.set_index('studioId')[[sum_filed1, 'date']]
        ndf2 = df2.set_index('live_room_id')[[sum_filed2]]
        result = pd.concat([ndf1, ndf2], axis=1)
        result['粉丝成交GMV'] = result[sum_filed1] * result[sum_filed2]
        result = result.groupby('date')['粉丝成交GMV'].sum().apply(lambda x: format(x, '.2f'))
        return result.astype('float64')

    def get_累计观看人数(df=dfs['live_list_details_grouping_day'], sum_filed=累计观看人数, date_type=sheet_date_type):
        return df.groupby([date_type])[sum_filed].sum()

    def get_新粉成交人数(df=dfs['live_detail_person_day'], sum_filed1=成交人群分析_粉丝维度_粉丝人数, sum_filed2=成交人群分析_粉丝维度_粉丝人数_新增粉丝占比, date_type=sheet_date_type):
        ndf1 = df.loc[:, [sum_filed1, date_type, sum_filed2, 'id']]
        ndf1.loc[:, '新粉成交人数'] = ndf1[sum_filed1] * ndf1[sum_filed2]
        return ndf1.groupby(date_type)['新粉成交人数'].sum()

    def get_互动成交人数(df=dfs['live_detail_person_day'], sum_filed=成交人群分析_互动维度_有互动人数, date_type=sheet_date_type):
        return df.groupby(date_type)[sum_filed].sum()


def merg_import(folder_path: str = People.export_folder):
    """合并表格

    Args:
        folder_path (str, optional): _description_. Defaults to Daily.export_folder.
    """
    logger.info('开始合并表格')
    output_file = f'{People.export_folder}/【{People.account_name}】{People.class_name}_底表与导出_GAP.xlsx'

    # 读取指定文件夹下的所有csv文件
    try:
        csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]

        # 如果文件不存在，创建一个新文件
        if not os.path.exists(output_file):
            pd.DataFrame().to_excel(output_file, index=False)

        # 将每个csv文件的数据作为sheet，并新增到整个文件中
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for csv_file in csv_files:
                csv_path = os.path.join(folder_path, csv_file)
                temp_df = pd.read_csv(csv_path)
                temp_df.to_excel(writer, sheet_name=csv_file.replace('.csv', '').replace('live_', '').replace('list_', '')[:31], index=False)
            dc = down(host=hosts['uat'], data=rq_data)
            dcf = pd.read_excel(dc)
            dcf['日期'] = pd.to_datetime(dcf['日期'])
            dcf.to_excel(writer, sheet_name='导出', index=False)
        logger.success('合并表格完成')
        gap = load_workbook(output_file)
        set_gap(gap['导出'], gap['底表'], 1)
        gap.save(output_file)
    except Exception as e:
        logger.error(f'合并表格失败\n{e}')


def save():
    观看人数 = People.get_指定观看人数(turnover='不限', fans='不限', interaction='不限', customers='不限', watch='不限')
    成交人数 = People.get_成交人数()
    观看转化率 = People.get_观看转化率()
    首购粉丝 = People.get_首购粉丝数().apply(lambda x: format(x, '.3f')).astype('float')

    首购人数 = People.get_首购人数()
    首购粉丝占比 = (首购粉丝 / 首购人数).apply(lambda x: format(x, '.4f'))

    复购粉丝 = People.get_复购粉丝数()

    复购人数 = People.get_复购人数()

    复购粉丝占比 = (复购粉丝 / 复购人数).apply(lambda x: format(x, '.9f'))
    
    首购占比 = (首购人数 / 成交人数).apply(lambda x: format(x, '.9f'))
    复购占比 = (复购人数 / 成交人数).apply(lambda x: format(x, '.9f'))
    粉丝成交人数 = People.get_粉丝成交人数()

    粉丝观看人数 = People.get_粉丝观看人数()
    粉丝成交GMV = People.get_粉丝成交GMV()
    粉丝GPM = (粉丝成交GMV * 1000 / 粉丝观看人数).apply(lambda x: format(x, '.9f'))
    观看未购粉丝 = (粉丝观看人数 - 粉丝成交人数).apply(lambda x: format(x, '.2f'))
    累计观看人数 = People.get_累计观看人数()  # maybe do not need
    观看人数粉丝占比 = (粉丝观看人数 / 累计观看人数).apply(lambda x: format(x, '.9f'))
    成交人数粉丝占比 = (粉丝成交人数 / 成交人数).apply(lambda x: format(x, '.9f'))
    粉丝观看转化率 = (粉丝成交人数 / 粉丝观看人数).apply(lambda x: format(x, '.9f'))
    新粉观看人数 = People.get_指定观看人数(turnover='不限', fans='新粉丝', interaction='不限', customers='不限', watch='不限')
    观看人数新粉占比 = (新粉观看人数 / 观看人数).apply(lambda x: format(x, '.9f'))
    新粉成交人数 = People.get_新粉成交人数()
    粉丝成交人数新粉占比 = (新粉成交人数 / 粉丝成交人数).apply(lambda x: format(x, '.4f'))
    成交人数新粉占比 = (新粉成交人数 / 成交人数).apply(lambda x: format(x, '.9f'))
    新粉观看转化率 = (新粉成交人数 / 新粉观看人数).apply(lambda x: format(x, '.9f'))
    老粉观看人数 = (粉丝观看人数 - 新粉观看人数).apply(lambda x: format(x, '.2f')).astype('float64')
    观看人数老粉占比 = (老粉观看人数 / 观看人数).apply(lambda x: format(x, '.9f'))
    老粉成交人数 = (粉丝成交人数 - 新粉成交人数).apply(lambda x: format(x, '.0f')).astype('float64')
    成交人数老粉占比 = (老粉成交人数 / 成交人数).apply(lambda x: format(x, '.9f'))
    老粉观看转化率 = (老粉成交人数 / 老粉观看人数).apply(lambda x: format(x, '.9f'))
    互动成交人数 = People.get_互动成交人数()
    互动成交人数占比 = (互动成交人数 / 成交人数).apply(lambda x: format(x, '.9f'))

    title = {
        '观看人数': 观看人数,
        '成交人数': 成交人数,
        '观看转化率': 观看转化率,
        '首购粉丝': 首购粉丝,
        '首购粉丝占比': 首购粉丝占比,
        '复购粉丝': 复购粉丝,
        '复购粉丝占比': 复购粉丝占比,
        '首购人数': 首购人数,
        '首购占比': 首购占比,
        '复购人数': 复购人数,
        '复购占比': 复购占比,
        '观看未购粉丝': 观看未购粉丝,
        '粉丝成交GMV': 粉丝成交GMV,
        '粉丝GPM': 粉丝GPM,
        '粉丝观看人数': 粉丝观看人数,
        '观看人数粉丝占比': 观看人数粉丝占比,
        '粉丝成交人数': 粉丝成交人数,
        '成交人数粉丝占比': 成交人数粉丝占比,
        '粉丝观看转化率': 粉丝观看转化率,
        '新粉观看人数': 新粉观看人数,
        '观看人数新粉占比': 观看人数新粉占比,
        '新粉成交人数': 新粉成交人数,
        '粉丝成交人数新粉占比': 粉丝成交人数新粉占比,
        '成交人数新粉占比': 成交人数新粉占比,
        '新粉观看转化率': 新粉观看转化率,
        '老粉观看人数': 老粉观看人数,
        '观看人数老粉占比': 观看人数老粉占比,
        '老粉成交人数': 老粉成交人数,
        '成交人数老粉占比': 成交人数老粉占比,
        '老粉观看转化率': 老粉观看转化率,
        '互动成交人数': 互动成交人数,
        '互动成交人数占比': 互动成交人数占比,
    }
    export = pd.concat(title.values(), axis=1, keys=title.keys()).fillna(0).replace('nan%', 0).replace('inf%', 0).replace('nan', 0).replace('inf', 0).sort_index()
    try:
        export.to_csv(f'{People.export_folder}/底表.csv', index_label=['日期'])
        logger.success(f'{People.export_folder}/底表.csv 导出成功')
    except Exception as e:
        logger.error(f'{People.export_folder}/底表.csv 导出失败:\n{e}')
    People.export_import_csv()


if __name__ == '__main__':
    People.export_import_csv()
    save()
    merg_import()
    print('需要留意日期是否完整,如果缺失需要手动补充日期填充0\n,全选表格,ctrl+g,选择空的,输入0,按ctrl+enter补充')
    input('按任意键退出')
