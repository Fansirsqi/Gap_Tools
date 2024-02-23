import os
from sys import stdout

import dotenv
import pandas as pd
from loguru import logger
from config import configs

dotenv.load_dotenv(override=True, verbose=True)

IS_DEBUG = os.getenv('IS_DEBUG')
logger.remove()
logger.add(stdout, level='INFO', colorize=True, format='<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}')

logger.add('流量.log', encoding='utf-8', format='<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}')
_account_name = os.getenv('ACCOUNT_NAME')
"""账号名称"""
_s_date = configs.START_DATE
"""开始日期"""
_e_date = configs.END_DATE
"""结束日期"""
BASEFLOAD = configs.BASEFLOAD


class Sales:
    class_name = '销售概览'
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
    export_csv_floader = f'./export_{class_name}'
    """导出文件夹"""

    dfs = {
        'live_list_details_traffic_time_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_list_details_traffic_time_day.csv'),
        'live_detail_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_detail_day.csv'),
        'scrm_ocean_daily': pd.read_csv(f'{baseFload}/scrm_ocean_daily.csv'),
        'live_growth_conversion_funnel_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_growth_conversion_funnel_day.csv'),
        'live_list_details_grouping_day': pd.read_csv(f'{baseFload}/{csvPerfix}live_list_details_grouping_day.csv'),
        'index_business_overview_day': pd.read_csv(f'{baseFload}/{csvPerfix}index_business_overview_day.csv'),
    }
    """需要导出的完整底表"""
    
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
            df['biz_date'] = pd.to_datetime(df['biz_date'])
            dfs[df_key] = df[(df['author_nick_name'] == account_name) & (df['biz_date'].between(start_date, end_date))].sort_values(by=['biz_date'])
        elif df_key == 'scrm_ocean_daily':
            df[sheet_date_type] = pd.to_datetime(df[sheet_date_type])
            dfs[df_key] = df[(df['account_name'] == account_name) & (df['marketing_goal'] == 'LIVE_PROM_GOODS') & (df[sheet_date_type].between(start_date, end_date))].sort_values(by=[sheet_date_type])
        elif df_key == 'fly_live_detail_first_prchase_day' or df_key == 'index_business_overview_day':
            df['date'] = pd.to_datetime(df['date'])
            if '圣罗兰' in account_name:
                account_name = 'YSL圣罗兰美妆官方旗舰店'
            elif '兰蔻' in account_name:
                account_name = '兰蔻LANCOME官方旗舰店'  # 此处的date请不要使用sheet_date_type，以免产生歧义
            elif '小美盒' in account_name:
                account_name = '欧莱雅集团小美盒官方旗舰店'
            elif '科颜氏' in account_name:
                account_name = "科颜氏KIEHL'S官方旗舰店"
            elif '碧欧泉' in account_name:
                account_name = '碧欧泉BIOTHERM官方旗舰店'
            elif 'HR赫莲' in account_name:
                account_name = 'HR赫莲娜官方旗舰店'
            elif '植村秀' in account_name:
                account_name = '植村秀shu uemura官方旗舰店'

            dfs[df_key] = df[
                (df['store_name'] == account_name) & (df['date'].between(start_date, end_date))
                # & (df["account_type"] == "渠道账号")
            ].sort_values(by='date')
        elif df_key == 'live_list_details_traffic_traffic_time_day':  # 此处需要过滤部分条件
            df[sheet_date_type] = pd.to_datetime(df[sheet_date_type])
            dfs[df_key] = df[(df['account_name'] == account_name) & (df['flowChannel'] != '整体') & (df['flowChannel'] != '全域推广') & (df['channelName'] != '整体') & (df[sheet_date_type].between(start_date, end_date))].sort_values(by=sheet_date_type)
        else:
            df[sheet_date_type] = pd.to_datetime(df[sheet_date_type])
            dfs[df_key] = df[(df['account_name'] == account_name) & (df[sheet_date_type].between(start_date, end_date))].sort_values(by=sheet_date_type)
    logger.debug('data import success !')

    直播间成交金额 = 'roomTransactionAmount'
    观看人次 = 'audience'
    成交人数 = 'watchRate'
    直播间曝光人数 = 'exposureCount'
    观看人数 = 'cumulativeAudience'
    商品点击人数 = 'storeClickCount'
    整体视角_成交金额 = 'overall_pay_amt'
    自营视角_成交金额 = 'self_pay_amt'
    整体视角_退款金额 = 'overall_refund_amt'
    整体视角_成交订单数 = 'overall_pay_cnt'
    整体视角_成交人数 = 'overall_pay_ucnt'
    商品曝光次数 = 'overall_product_show_pv'
    自营视角_商品点击pv = 'self_product_click_pv'
    带货视角_商品点击pv = 'commerce_product_click_pv'

    def export_import_csv(dfs=dfs, export_folder=export_csv_floader):
        """导出csv文件:
        Args:
            dfs (_type_, optional): 被类处理后的dfs字典. Defaults to dfs.
            export_folder (_type_, optional): 指定的导出文件夹. Defaults to export_csv_floader.
        """
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
        for df_key, df in dfs.items():
            df.to_csv(f'{export_folder}/{df_key}.csv', index=False)

    def get_5GMV(df=dfs['live_detail_day'], sum_filed=直播间成交金额, date_type=sheet_date_type):
        logger.debug(f'\n{df}')
        return df.groupby(date_type)[sum_filed].sum()

    def get_指定观看人次PV(df=dfs['live_list_details_traffic_time_day'], sum_filed=观看人次, date_type=sheet_date_type, flowChannel=None):
        logger.debug(f'\n{df}')
        if flowChannel is not None:
            df = df[df['flowChannel'] == flowChannel]
        return df.groupby(date_type)[sum_filed].sum()

    def get_成交人数(df=dfs['live_growth_conversion_funnel_day'], sum_filed=成交人数, date_type=sheet_date_type):
        logger.debug(f'\n{df}')
        return df.groupby(date_type)[sum_filed].sum()

    def get_曝光人数(df=dfs['live_growth_conversion_funnel_day'], sum_filed=[直播间曝光人数], date_type=sheet_date_type):
        logger.debug(f'\n{df}')
        return df.groupby(date_type)[sum_filed[0]].sum()

    def get_观看人数(df=dfs['live_list_details_grouping_day'], sum_filed=观看人数, date_type=sheet_date_type):
        logger.debug(f'\n{df}')
        return df.groupby(date_type)[sum_filed].sum()

    def get_外层CTR(df1=dfs['live_growth_conversion_funnel_day'], df2=dfs['live_list_details_grouping_day'], sum_filed=[观看人数, 直播间曝光人数], date_type=sheet_date_type):
        logger.debug(f'\n{df1}')
        logger.debug(f'\n{df2}')
        ndf1 = df2.groupby(date_type)[sum_filed[0]].sum()  #!!!
        ndf2 = df1.groupby(date_type)[sum_filed[1]].sum()  #!!!
        ndf3 = (ndf1 / ndf2).apply(lambda x: format(x, '.11'))
        return ndf3

    def get_商品点击人数(df=dfs['live_growth_conversion_funnel_day'], sum_filed=[商品点击人数], date_type=sheet_date_type):
        logger.debug(f'\n{df}')
        return df.groupby(date_type)[sum_filed[0]].sum()

    def get_sheet1(df=dfs['index_business_overview_day'], sum_filed=None, date_type='date'):
        if sum_filed is None:
            print('请指定维度')
        else:
            return df.groupby(date_type)[sum_filed[0]].sum()

    def get_sheet1_商品点击次数(df=dfs['index_business_overview_day'], sum_fled=[自营视角_商品点击pv, 带货视角_商品点击pv], date_type='date'):
        df['商品_点击次数'] = df[sum_fled[0]] + df[sum_fled[1]]
        return df.groupby(date_type)['商品_点击次数'].sum()

    def get_sheet1_商品点击率_次数_(df=dfs['index_business_overview_day'], sum_fled=[自营视角_商品点击pv, 带货视角_商品点击pv], date_type='date'):
        df['商品_点击次数'] = df[sum_fled[0]] + df[sum_fled[1]]
        ndf = df.groupby(date_type)['overall_product_show_pv'].sum()
        ndf1 = df.groupby(date_type)['商品_点击次数'].sum()
        return (ndf1 / ndf).apply(lambda x: format(x, '.2%'))


def saver_():
    Sales.export_import_csv()
    GMV_ACH_直播期间成交金额_ = Sales.get_5GMV()
    GMV_ACH_直播间成交金额_ = GMV_ACH_直播期间成交金额_
    观看人次PV = Sales.get_指定观看人次PV()
    GPM_直播期间成交金额_ = (GMV_ACH_直播期间成交金额_ * 1000 / 观看人次PV).apply(lambda x: format(x, '.9f'))
    GPM_直播间成交金额_ = (GMV_ACH_直播间成交金额_ * 1000 / 观看人次PV).apply(lambda x: format(x, '.2f'))
    成交人数 = Sales.get_成交人数()
    曝光人数 = Sales.get_曝光人数()
    观看人数UV = Sales.get_观看人数()
    外层CTR = Sales.get_外层CTR()
    商品点击人数 = Sales.get_商品点击人数()
    商品点击率 = (商品点击人数 / 观看人数UV).apply(lambda x: format(x, '.11'))
    点击成交转化率 = (成交人数 / 商品点击人数).apply(lambda x: format(x, '.11'))
    观看转化率 = (成交人数 / 观看人数UV).apply(lambda x: format(x, '.11'))
    title = {
        'GMV_ACH_直播期间成交金额_': GMV_ACH_直播期间成交金额_,
        'GMV_ACH_直播间成交金额_': GMV_ACH_直播间成交金额_,
        'GPM_直播期间成交金额_': GPM_直播期间成交金额_,
        'GPM_直播间成交金额_': GPM_直播间成交金额_,
        '成交人数': 成交人数,
        '曝光人数': 曝光人数,
        '观看人数UV': 观看人数UV,
        '外层CTR': 外层CTR,
        '商品点击人数': 商品点击人数,
        '商品点击率': 商品点击率,
        '点击成交转化率': 点击成交转化率,
        '观看转化率': 观看转化率,
    }
    export = (
        pd.concat(
            title.values(),
            axis=1,
            keys=title.keys(),
        )
        .fillna(0)
        .replace('nan%', 0)
        .replace('nan', 0)
        .replace('inf%', 0)
        .replace('inf', 0)
        .sort_index()
    )
    try:
        export.to_csv(f'{Sales.export_csv_floader}/【{Sales.account_name}】_销售概览底表5.csv', index_label=['日期'])
        logger.success(f'{Sales.export_csv_floader}/【{Sales.account_name}】_销售概览底表5.csv')
    except Exception as e:
        logger.error(f'{Sales.export_csv_floader}/【{Sales.account_name}】_销售概览底表5.csv 导出失败/n:{e}')


# GMV1 = Sales.get_sheet1(sum_filed=[[Sales.整体视角_成交金额]])
# 自营视角_成交金额 = Sales.get_sheet1(sum_filed=[[Sales.自营视角_成交金额]])
# 退款金额_元_ = Sales.get_sheet1(sum_filed=[[Sales.整体视角_退款金额]])
# 成交订单数 = Sales.get_sheet1(sum_filed=[[Sales.整体视角_成交订单数]])
# 成交人数 = Sales.get_sheet1(sum_filed=[[Sales.整体视角_成交人数]])
# 商品曝光次数 = Sales.get_sheet1(sum_filed=[[Sales.商品曝光次数]])
# 商品点击次数 = Sales.get_sheet1_商品点击次数()
# 商品点击率_次数_ = Sales.get_sheet1_商品点击率_次数_()

# print(商品点击率_次数_)

saver_()
