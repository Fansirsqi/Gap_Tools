import os
import pandas as pd
from sys import stdout
from loguru import logger

IS_DEBUG = True
logger.remove()
logger.add(
    stdout,
    level="INFO",
    # encoding="utf-8",
    colorize=True,
    format="<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}",
)

logger.add(
    "流量.log",
    encoding="utf-8",
    format="<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}",
)
if IS_DEBUG:
    _account_name = "欧莱雅集团小美盒"
    """账号名称"""
    _s_date = "2023-07-01"
    """开始日期"""
    _e_date = "2023-09-30"
    """结束日期"""
else:
    try:
        input("欢迎使[店播流量]用取数工具\n按Enter按键继续.....\n")
        print("需要用到的底表如下-请放在csv文件夹下")
        print("scrm_dy_report_app_fxg_live_number_people_covered_day.csv")
        print("scrm_dy_report_app_fxg_live_growth_conversion_funnel_day.csv")
        print("scrm_dy_report_app_fxg_live_detail_person_day.csv")
        print("scrm_dy_report_app_fxg_live_crowd_data_day.csv")
        print("scrm_dy_report_app_fxg_live_detail_day.csv")
        print("scrm_dy_report_app_fxg_fly_live_detail_first_prchase_day.csv")
        print("scrm_dy_report_app_fxg_live_list_details_grouping_day.csv")
        input("确认文件夹下面有以上文件\n按Enter按键继续.....\n")
        _account_name = input("请输入account_name：")
        """账号名称"""
        _s_date = input("请输入开始日期：(eg:2023-08-01)")
        """开始日期"""
        _e_date = input("请输入结束日期：(eg:2023-09-30)")
        """结束日期"""
    except KeyboardInterrupt:
        _account_name = None
        """账号名称"""
        _s_date = None
        """开始日期"""
        _e_date = None
        print("\n用户终止程序")
        input("按任意键退出")


class SalesOverview:
    class_name = "销售概览"
    current_time = str(pd.Timestamp.now())
    logger.debug("-" * 30 + "| " + current_time + " |" + "-" * 30)
    account_name = _account_name
    if account_name is None or account_name == "":
        logger.error("账号名称为空或者错误，程序退出")
        exit()
    start_date = _s_date
    end_date = _e_date
    baiku_date_type = "bizDate"
    """百库底表(大部分)通用日期字段: bizDate"""
    qianchuan_date_type = "date"
    """千川底表(大部分)通用日期字段: date"""
    baseFload = r"C:\Users\Administrator\OneDrive\Gap文档\csv"
    """底表父文件夹"""
    csvPerfix = "scrm_dy_report_app_fxg_"
    """百库底表前缀"""
    export_csv_floader = f"./export_{class_name}"
    """导出文件夹"""
    df1 = pd.read_csv(f"{baseFload}/{csvPerfix}live_list_details_traffic_time_day.csv")
    df2 = pd.read_csv(f"{baseFload}/{csvPerfix}live_detail_day.csv")
    df3 = pd.read_csv(f"{baseFload}/scrm_ocean_daily.csv")
    df4 = pd.read_csv(f"{baseFload}/{csvPerfix}live_growth_conversion_funnel_day.csv")
    df5 = pd.read_csv(f"{baseFload}/{csvPerfix}live_list_details_grouping_day.csv")
    df6 = pd.read_csv(f"{baseFload}/{csvPerfix}index_business_overview_day.csv")
    dfs = {
        "live_list_details_traffic_time_day": df1,
        "live_detail_day": df2,
        "scrm_ocean_daily": df3,
        "live_growth_conversion_funnel_day": df4,
        "live_list_details_grouping_day": df5,
        "index_business_overview_day": df6,
    }
    """需要导出的完整底表"""

    logger.debug("init data import ...")
    for df_key, df in dfs.items():
        logger.debug(f"导入 {df_key}")
        if df_key == "live_analysis_live_details_all_day":
            df["biz_date"] = pd.to_datetime(df["biz_date"])
            dfs[df_key] = df[(df["author_nick_name"] == account_name) & (df["biz_date"].between(start_date, end_date))].sort_values(by=["biz_date"])
        elif df_key == "scrm_ocean_daily":
            df[qianchuan_date_type] = pd.to_datetime(df[qianchuan_date_type])
            dfs[df_key] = df[(df["account_name"] == account_name) & (df["marketing_goal"] == "LIVE_PROM_GOODS") & (df[qianchuan_date_type].between(start_date, end_date))].sort_values(
                by=[qianchuan_date_type]
            )
        elif df_key == "fly_live_detail_first_prchase_day" or df_key == "index_business_overview_day":
            df["date"] = pd.to_datetime(df["date"])
            if "圣罗兰" in account_name:
                account_name = "YSL圣罗兰美妆官方旗舰店"
            elif "兰蔻" in account_name:
                account_name = "兰蔻LANCOME官方旗舰店"  # 此处的date请不要使用qianchuan_date_type，以免产生歧义
            elif "小美盒" in account_name:
                account_name = "欧莱雅集团小美盒官方旗舰店"
            elif "科颜氏" in account_name:
                account_name = "科颜氏KIEHL'S官方旗舰店"
            elif "碧欧泉" in account_name:
                account_name = "碧欧泉BIOTHERM官方旗舰店"
            elif "HR赫莲" in account_name:
                account_name = "HR赫莲娜官方旗舰店"
            elif "植村秀" in account_name:
                account_name = "植村秀shu uemura官方旗舰店"

            dfs[df_key] = df[
                (df["store_name"] == account_name) & (df["date"].between(start_date, end_date))
                # & (df["account_type"] == "渠道账号")
            ].sort_values(by="date")
        elif df_key == "live_list_details_traffic_traffic_time_day":  # 此处需要过滤部分条件
            df[baiku_date_type] = pd.to_datetime(df[baiku_date_type])
            dfs[df_key] = df[
                (df["account_name"] == account_name)
                & (df["flowChannel"] != "整体")
                & (df["flowChannel"] != "全域推广")
                & (df["channelName"] != "整体")
                & (df[baiku_date_type].between(start_date, end_date))
            ].sort_values(by=baiku_date_type)
        else:
            df[baiku_date_type] = pd.to_datetime(df[baiku_date_type])
            dfs[df_key] = df[(df["account_name"] == account_name) & (df[baiku_date_type].between(start_date, end_date))].sort_values(by=baiku_date_type)
    logger.debug("data import success !")

    直播间成交金额 = "roomTransactionAmount"
    观看人次 = "audience"
    成交人数 = "watchRate"
    直播间曝光人数 = "exposureCount"
    观看人数 = "cumulativeAudience"
    商品点击人数 = "storeClickCount"
    整体视角_成交金额 = "overall_pay_amt"
    自营视角_成交金额 = "self_pay_amt"
    整体视角_退款金额 = "overall_refund_amt"
    整体视角_成交订单数 = "overall_pay_cnt"
    整体视角_成交人数 = "overall_pay_ucnt"
    商品曝光次数 = "overall_product_show_pv"
    自营视角_商品点击pv = "self_product_click_pv"
    带货视角_商品点击pv = "commerce_product_click_pv"

    def export_import_csv(dfs=dfs, export_folder=export_csv_floader):
        """导出csv文件:
        Args:
            dfs (_type_, optional): 被类处理后的dfs字典. Defaults to dfs.
            export_folder (_type_, optional): 指定的导出文件夹. Defaults to export_csv_floader.
        """
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
        for df_key, df in dfs.items():
            df.to_csv(f"{export_folder}/{df_key}.csv", index=False)

    def get_5GMV(df=dfs["live_detail_day"], sum_filed=直播间成交金额, date_type=baiku_date_type):
        logger.debug(f"\n{df}")
        return df.groupby(date_type)[sum_filed].sum()

    def get_指定观看人次PV(df=dfs["live_list_details_traffic_time_day"], sum_filed=观看人次, date_type=baiku_date_type, flowChannel=None):
        logger.debug(f"\n{df}")
        if flowChannel is not None:
            df = df[df["flowChannel"] == flowChannel]
        return df.groupby(date_type)[sum_filed].sum()

    def get_成交人数(df=dfs["live_growth_conversion_funnel_day"], sum_filed=成交人数, date_type=baiku_date_type):
        logger.debug(f"\n{df}")
        return df.groupby(date_type)[sum_filed].sum()

    def get_曝光人数(df=dfs["live_growth_conversion_funnel_day"], sum_filed=[直播间曝光人数], date_type=baiku_date_type):
        logger.debug(f"\n{df}")
        return df.groupby(date_type)[sum_filed[0]].sum()

    def get_观看人数(df=dfs["live_list_details_grouping_day"], sum_filed=观看人数, date_type=baiku_date_type):
        logger.debug(f"\n{df}")
        return df.groupby(date_type)[sum_filed].sum()

    def get_外层CTR(df1=dfs["live_growth_conversion_funnel_day"], df2=dfs["live_list_details_grouping_day"], sum_filed=[观看人数, 直播间曝光人数], date_type=baiku_date_type):
        logger.debug(f"\n{df1}")
        logger.debug(f"\n{df2}")
        ndf1 = df2.groupby(date_type)[sum_filed[0]].sum()  #!!!
        ndf2 = df1.groupby(date_type)[sum_filed[1]].sum()  #!!!
        ndf3 = (ndf1 / ndf2).apply(lambda x: format(x, ".11"))
        return ndf3

    def get_商品点击人数(df=dfs["live_growth_conversion_funnel_day"], sum_filed=[商品点击人数], date_type=baiku_date_type):
        logger.debug(f"\n{df}")
        return df.groupby(date_type)[sum_filed[0]].sum()

    def get_sheet1(df=dfs["index_business_overview_day"], sum_filed=None, date_type="date"):
        if sum_filed is None:
            print("请指定维度")
        else:
            return df.groupby(date_type)[sum_filed[0]].sum()

    def get_sheet1_商品点击次数(df=dfs["index_business_overview_day"], sum_fled=[自营视角_商品点击pv, 带货视角_商品点击pv], date_type="date"):
        df["商品_点击次数"] = df[sum_fled[0]] + df[sum_fled[1]]
        return df.groupby(date_type)["商品_点击次数"].sum()
    
    def get_sheet1_商品点击率_次数_(df=dfs["index_business_overview_day"], sum_fled=[自营视角_商品点击pv, 带货视角_商品点击pv], date_type="date"):
        df["商品_点击次数"] = df[sum_fled[0]] + df[sum_fled[1]]
        ndf = df.groupby(date_type)['overall_product_show_pv'].sum()
        ndf1 = df.groupby(date_type)["商品_点击次数"].sum()
        return (ndf1 / ndf).apply(lambda x: format(x, ".2%"))


def saver_():
    SalesOverview.export_import_csv()
    GMV_ACH_直播期间成交金额_ = SalesOverview.get_5GMV()
    GMV_ACH_直播间成交金额_ = GMV_ACH_直播期间成交金额_
    观看人次PV = SalesOverview.get_指定观看人次PV()
    GPM_直播期间成交金额_ = (GMV_ACH_直播期间成交金额_ * 1000 / 观看人次PV).apply(lambda x: format(x, ".9f"))
    GPM_直播间成交金额_ = (GMV_ACH_直播间成交金额_ * 1000 / 观看人次PV).apply(lambda x: format(x, ".2f"))
    成交人数 = SalesOverview.get_成交人数()
    曝光人数 = SalesOverview.get_曝光人数()
    观看人数UV = SalesOverview.get_观看人数()
    外层CTR = SalesOverview.get_外层CTR()
    商品点击人数 = SalesOverview.get_商品点击人数()
    商品点击率 = (商品点击人数 / 观看人数UV).apply(lambda x: format(x, ".11"))
    点击成交转化率 = (成交人数 / 商品点击人数).apply(lambda x: format(x, ".11"))
    观看转化率 = (成交人数 / 观看人数UV).apply(lambda x: format(x, ".11"))
    title = {
        "GMV_ACH_直播期间成交金额_": GMV_ACH_直播期间成交金额_,
        "GMV_ACH_直播间成交金额_": GMV_ACH_直播间成交金额_,
        "GPM_直播期间成交金额_": GPM_直播期间成交金额_,
        "GPM_直播间成交金额_": GPM_直播间成交金额_,
        "成交人数": 成交人数,
        "曝光人数": 曝光人数,
        "观看人数UV": 观看人数UV,
        "外层CTR": 外层CTR,
        "商品点击人数": 商品点击人数,
        "商品点击率": 商品点击率,
        "点击成交转化率": 点击成交转化率,
        "观看转化率": 观看转化率,
    }
    export = (
        pd.concat(
            title.values(),
            axis=1,
            keys=title.keys(),
        )
        .fillna(0)
        .replace("nan%", 0)
        .replace("nan", 0)
        .replace("inf%", 0)
        .replace("inf", 0)
        .sort_index()
    )
    try:
        export.to_csv(f"{SalesOverview.export_csv_floader}/【{SalesOverview.account_name}】_销售概览底表5.csv", index_label=["日期"])
        logger.success(f"{SalesOverview.export_csv_floader}/【{SalesOverview.account_name}】_销售概览底表5.csv")
    except Exception as e:
        logger.error(f"{SalesOverview.export_csv_floader}/【{SalesOverview.account_name}】_销售概览底表5.csv 导出失败/n:{e}")


# GMV1 = SalesOverview.get_sheet1(sum_filed=[[SalesOverview.整体视角_成交金额]])
# 自营视角_成交金额 = SalesOverview.get_sheet1(sum_filed=[[SalesOverview.自营视角_成交金额]])
# 退款金额_元_ = SalesOverview.get_sheet1(sum_filed=[[SalesOverview.整体视角_退款金额]])
# 成交订单数 = SalesOverview.get_sheet1(sum_filed=[[SalesOverview.整体视角_成交订单数]])
# 成交人数 = SalesOverview.get_sheet1(sum_filed=[[SalesOverview.整体视角_成交人数]])
# 商品曝光次数 = SalesOverview.get_sheet1(sum_filed=[[SalesOverview.商品曝光次数]])
# 商品点击次数 = SalesOverview.get_sheet1_商品点击次数()
# 商品点击率_次数_ = SalesOverview.get_sheet1_商品点击率_次数_()

# print(商品点击率_次数_)

saver_()