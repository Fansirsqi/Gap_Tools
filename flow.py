import os
import pandas as pd
from sys import stdout
from loguru import logger
from dotenv import load_dotenv


load_dotenv(override=True, verbose=True)

IS_DEBUG = os.getenv('IS_DEBUG', 'false').lower() == 'true'
print(type(IS_DEBUG), IS_DEBUG)
logger.remove()
logger.add(stdout, level="INFO", colorize=True, format="<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}")

logger.add("流量.log", encoding="utf-8", format="<g>{time:MM-DD HH:mm:ss}</g> <level><w>[</w>{level}<w>]</w></level> | {message}")
if IS_DEBUG:
    _account_name = os.getenv("ACCOUNT_NAME")
    """账号名称"""
    _s_date = os.getenv("START_DATE")
    """开始日期"""
    _e_date = os.getenv("END_DATE")
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


class Flow:
    class_name = "店播流量"
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
    baseFload = os.getenv("BASEFLOAD")
    """底表父文件夹"""
    csvPerfix = "scrm_dy_report_app_fxg_"
    """百库底表前缀"""
    export_csv_floader = f"./export_{class_name}"
    """导出文件夹"""
    df1 = pd.read_csv(f"{baseFload}/{csvPerfix}live_list_details_traffic_time_day.csv")
    df2 = pd.read_csv(f"{baseFload}/{csvPerfix}live_detail_day.csv")
    df3 = pd.read_csv("csv/scrm_ocean_daily.csv")
    dfs = {
        "live_list_details_traffic_time_day": df1,
        "live_detail_day": df2,
        "scrm_ocean_daily": df3,
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
            dfs[df_key] = df[(df["account_name"] == account_name) & (df["marketing_goal"] == "LIVE_PROM_GOODS") & (df[qianchuan_date_type].between(start_date, end_date))].sort_values(by=[qianchuan_date_type])
        elif df_key == "fly_live_detail_first_prchase_day":
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
        elif df_key == "live_list_details_traffic_time_day":  # 此处需要过滤部分条件
            df[baiku_date_type] = pd.to_datetime(df[baiku_date_type])
            dfs[df_key] = df[(df["account_name"] == account_name) & (df["flowChannel"] != "整体") & (df["flowChannel"] != "全域推广") & (df["channelName"] != "整体") & (df[baiku_date_type].between(start_date, end_date))].sort_values(by=baiku_date_type)
        else:
            df[baiku_date_type] = pd.to_datetime(df[baiku_date_type])
            dfs[df_key] = df[(df["account_name"] == account_name) & (df[baiku_date_type].between(start_date, end_date))].sort_values(by=baiku_date_type)

    logger.debug("data import success !")

    not_channelName = "整体"
    not_flowChannel = "全域推广"
    PV = "audience"
    订单数 = "transactionOrderNumber"
    GMV = "dealAmount"
    消耗 = "stat_cost"
    成交订单金额 = "pay_order_amount"
    直播间成金额 = "roomTransactionAmount"

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

    def get_PV(df=dfs["live_list_details_traffic_time_day"], filed=PV, date_type=baiku_date_type, not_channelName=not_channelName, not_flowChannel=not_flowChannel):
        ndf = df[(df["channelName"] != not_channelName) & (df["flowChannel"] != not_flowChannel)]
        return ndf.groupby([date_type])[filed].sum()

    def get_自然PV(df=dfs["live_list_details_traffic_time_day"], date_type=baiku_date_type, filed=PV, not_channelName=not_channelName):
        ndf = df[(df["flowChannel"] == "自然流量") & (df["channelName"] != not_channelName)]
        return ndf.groupby([date_type])[filed].sum()

    def get_付费PV(df=dfs["live_list_details_traffic_time_day"], date_type=baiku_date_type, filed=PV, not_channelName=not_channelName):
        ndf = df[(df["flowChannel"] == "付费流量") & (df["channelName"] != not_channelName)]
        return ndf.groupby([date_type])[filed].sum()

    def get_自然订单数(df=dfs["live_list_details_traffic_time_day"], date_type=baiku_date_type, filed=订单数, not_channelName=not_channelName):
        ndf = df[(df["flowChannel"] == "自然流量") & (df["channelName"] != not_channelName)]
        return ndf.groupby([date_type])[filed].sum()

    def get_付费订单数(df=dfs["live_list_details_traffic_time_day"], date_type=baiku_date_type, filed=订单数, not_channelName=not_channelName):
        ndf = df[(df["flowChannel"] == "付费流量") & (df["channelName"] != not_channelName)]
        return ndf.groupby([date_type])[filed].sum()

    def get_指定渠道PV(df=dfs["live_list_details_traffic_time_day"], date_type=baiku_date_type, filed=PV, not_flowChannel=not_flowChannel, filter_fields=None):
        ndf = df[(df["flowChannel"] != not_flowChannel) & (df["channelName"] == filter_fields)]
        return ndf.groupby([date_type])[filed].sum()

    def get_指定渠道GMV(df=dfs["live_list_details_traffic_time_day"], date_type=baiku_date_type, filed=GMV, not_flowChannel=not_flowChannel, filter_fields=None):
        ndf = df[(df["flowChannel"] != not_flowChannel) & (df["channelName"] == filter_fields)]
        return ndf.groupby([date_type])[filed].sum()

    def get_指定渠道订单数(df=dfs["live_list_details_traffic_time_day"], date_type=baiku_date_type, filed=订单数, not_flowChannel=not_flowChannel, filter_fields=None):
        """
        filter_fields:
        @千川PC版 : 千川PC版渠道
        @推荐feed : 推荐feed渠道
        @关注 : 关注渠道
        @其他 : 其他渠道
        @搜索 : 搜索渠道
        @品牌广告 : 品牌广告渠道
        @小店随心推 : 小店随心推渠道
        @个人主页&店铺&橱窗 : 个人主页_店铺_橱窗
        @抖音商城推荐 : 抖音商城推荐
        @直播广场 : 直播广场渠道
        @短视频引流 : 短视频引流渠道
        @其他推荐场景 : 其他推荐场景渠道
        @活动页 : 活动页渠道
        @同城 : 同城渠道
        @头条西瓜 : 头条西瓜渠道
        @其他广告 : 其他广告渠道
        @千川品牌广告 : 千川品牌广告渠道
        @return
        """
        ndf = df[(df["flowChannel"] != not_flowChannel) & (df["channelName"] == filter_fields)]
        return ndf.groupby([date_type])[filed].sum()

    ## 以下是表2内容
    def get_投放金额(df=dfs["scrm_ocean_daily"], date_type=qianchuan_date_type, filed=消耗):
        return df.groupby([date_type])[filed].sum()

    def get_投放转化金额(df=dfs["scrm_ocean_daily"], date_type=qianchuan_date_type, filed=成交订单金额):
        return df.groupby([date_type])[filed].sum()

    def get_直播间成交金额(df=dfs["live_detail_day"], date_type=baiku_date_type, filed=直播间成金额):
        return df.groupby([date_type])[filed].sum()


def save2():
    投放金额 = Flow.get_投放金额()
    投放转化金额 = Flow.get_投放转化金额()
    直播间成交金额 = Flow.get_直播间成交金额()
    Take_Rate_直播期间 = (投放金额 / 直播间成交金额).apply(lambda x: format(x, ".9%"))
    Take_Rate_直播间 = (投放金额 / 直播间成交金额).apply(lambda x: format(x, ".2%"))
    Media_Contribution_直播期间 = (投放转化金额 / 直播间成交金额).apply(lambda x: format(x, ".9%"))
    Media_Contribution_直播间 = Media_Contribution_直播期间
    TTLROI_直播期间 = (直播间成交金额 / 投放金额).apply(lambda x: format(x, ".10"))
    TTLROI_直播间 = TTLROI_直播期间
    投放ROI = (投放转化金额 / 投放金额).apply(lambda x: format(x, ".10"))
    title = {
        "投放金额_元": 投放金额,
        "投放转化金额": 投放转化金额,
        "Take_Rate_直播期间成交金额": Take_Rate_直播期间,
        "Take_Rate_直播间成交金额": Take_Rate_直播间,
        "Media_Contribution_直播期间成交金额": Media_Contribution_直播期间,
        "Media_Contribution_直播间成交金额": Media_Contribution_直播间,
        "TTL_ROI_直播期间成交金额": TTLROI_直播期间,
        "TTL_ROI_直播间成交金额": TTLROI_直播间,
        "投放ROI": 投放ROI,
    }
    try:
        export = (
            pd.concat(
                title.values(),
                axis=1,
                keys=title.keys(),
            )
            .fillna(0)
            .replace("nan%", 0)
            .replace("inf%", 0)
            .replace("nan", 0)
            .replace("inf", 0)
            .sort_index()
        )
        export.sort_index().to_csv(f"{Flow.export_csv_floader}/【{Flow.account_name}】_流量底表2.csv", index_label=["日期"])
        logger.success(f"{Flow.export_csv_floader}/【{Flow.account_name}】_流量底表2.csv 保存成功")
    except Exception as e:
        print(e)
        logger.error(f"{Flow.export_csv_floader}/【{Flow.account_name}】_流量底表2.csv 保存失败:\n{e}")


# inf


def save1():
    PV = Flow.get_PV()
    自然PV = Flow.get_自然PV()
    自然PV_率 = (自然PV / PV).apply(lambda x: format(x, ".9%"))
    付费PV = Flow.get_付费PV()
    付费PV_率 = (付费PV / PV).apply(lambda x: format(x, ".9%"))
    自然订单数 = Flow.get_自然订单数()
    付费订单数 = Flow.get_付费订单数()
    TTL订单数 = 自然订单数 + 付费订单数
    自然CVR = (自然订单数 / 自然PV).apply(lambda x: format(x, ".9%"))
    付费CVR = (付费订单数 / 付费PV).apply(lambda x: format(x, ".9%"))
    TTLCVR = (TTL订单数 / PV).apply(lambda x: format(x, ".9%"))
    推荐feed渠道PV = Flow.get_指定渠道PV(filter_fields="推荐feed")
    直播广场渠道PV = Flow.get_指定渠道PV(filter_fields="直播广场")
    同城渠道PV = Flow.get_指定渠道PV(filter_fields="同城")
    其他推荐场景渠道PV = Flow.get_指定渠道PV(filter_fields="其他推荐场景")
    关注渠道PV = Flow.get_指定渠道PV(filter_fields="关注")
    搜索渠道PV = Flow.get_指定渠道PV(filter_fields="搜索")
    短视频引流渠道PV = Flow.get_指定渠道PV(filter_fields="短视频引流")
    个人主页_店铺_橱窗PV = Flow.get_指定渠道PV(filter_fields="个人主页&店铺&橱窗")
    抖音商城推荐PV = Flow.get_指定渠道PV(filter_fields="抖音商城推荐")
    活动页渠道PV = Flow.get_指定渠道PV(filter_fields="活动页")
    头条西瓜渠道PV = Flow.get_指定渠道PV(filter_fields="头条西瓜")
    其他渠道PV = Flow.get_指定渠道PV(filter_fields="其他")
    千川PC版渠道PV = Flow.get_指定渠道PV(filter_fields="千川PC版")
    小店随心推渠道PV = Flow.get_指定渠道PV(filter_fields="小店随心推")
    品牌广告渠道PV = Flow.get_指定渠道PV(filter_fields="品牌广告")
    其他广告渠道PV = Flow.get_指定渠道PV(filter_fields="其他广告")
    千川品牌广告渠道PV = Flow.get_指定渠道PV(filter_fields="千川品牌广告")

    推荐feed渠道GMV = Flow.get_指定渠道GMV(filter_fields="推荐feed")
    直播广场渠道GMV = Flow.get_指定渠道GMV(filter_fields="直播广场")
    同城渠道GMV = Flow.get_指定渠道GMV(filter_fields="同城")
    其他推荐场景渠道GMV = Flow.get_指定渠道GMV(filter_fields="其他推荐场景")
    关注渠道GMV = Flow.get_指定渠道GMV(filter_fields="关注")
    搜索渠道GMV = Flow.get_指定渠道GMV(filter_fields="搜索")
    短视频引流渠道GMV = Flow.get_指定渠道GMV(filter_fields="短视频引流")
    个人主页_店铺_橱窗GMV = Flow.get_指定渠道GMV(filter_fields="个人主页&店铺&橱窗")
    抖音商城推荐GMV = Flow.get_指定渠道GMV(filter_fields="抖音商城推荐")
    活动页渠道GMV = Flow.get_指定渠道GMV(filter_fields="活动页")
    头条西瓜渠道GMV = Flow.get_指定渠道GMV(filter_fields="头条西瓜")
    其他渠道GMV = Flow.get_指定渠道GMV(filter_fields="其他")
    千川PC版渠道GMV = Flow.get_指定渠道GMV(filter_fields="千川PC版")
    小店随心推渠道GMV = Flow.get_指定渠道GMV(filter_fields="小店随心推")
    品牌广告渠道GMV = Flow.get_指定渠道GMV(filter_fields="品牌广告")
    其他广告渠道GMV = Flow.get_指定渠道GMV(filter_fields="其他广告")
    千川品牌广告渠道GMV = Flow.get_指定渠道GMV(filter_fields="千川品牌广告")

    推荐feed渠道GPM = 推荐feed渠道GMV / 推荐feed渠道PV * 1000
    直播广场渠道GPM = 直播广场渠道GMV / 直播广场渠道PV * 1000
    同城渠道GPM = 同城渠道GMV / 同城渠道PV * 1000
    其他推荐场景渠道GPM = 其他推荐场景渠道GMV / 其他推荐场景渠道PV * 1000
    关注渠道GPM = 关注渠道GMV / 关注渠道PV * 1000
    搜索渠道GPM = 搜索渠道GMV / 搜索渠道PV * 1000
    短视频引流渠道GPM = 短视频引流渠道GMV / 短视频引流渠道PV * 1000
    个人主页_店铺_橱窗GPM = 个人主页_店铺_橱窗GMV / 个人主页_店铺_橱窗PV * 1000
    抖音商城推荐GPM = 抖音商城推荐GMV / 抖音商城推荐PV * 1000
    活动页渠道GPM = 活动页渠道GMV / 活动页渠道PV * 1000
    头条西瓜渠道GPM = 头条西瓜渠道GMV / 头条西瓜渠道PV * 1000
    其他渠道GPM = 其他渠道GMV / 其他渠道PV * 1000
    千川PC版渠道GPM = 千川PC版渠道GMV / 千川PC版渠道PV * 1000
    小店随心推渠道GPM = 小店随心推渠道GMV / 小店随心推渠道PV * 1000
    品牌广告渠道GPM = 品牌广告渠道GMV / 品牌广告渠道PV * 1000
    其他广告渠道GPM = 其他广告渠道GMV / 其他广告渠道PV * 1000
    千川品牌广告渠道GPM = 千川品牌广告渠道GMV / 千川品牌广告渠道PV * 1000

    推荐feed渠道订单数 = Flow.get_指定渠道订单数(filter_fields="推荐feed")
    直播广场渠道订单数 = Flow.get_指定渠道订单数(filter_fields="直播广场")
    同城渠道订单数 = Flow.get_指定渠道订单数(filter_fields="同城")
    其他推荐场景渠道订单数 = Flow.get_指定渠道订单数(filter_fields="其他推荐场景")
    关注渠道订单数 = Flow.get_指定渠道订单数(filter_fields="关注")
    搜索渠道订单数 = Flow.get_指定渠道订单数(filter_fields="搜索")
    短视频引流渠道订单数 = Flow.get_指定渠道订单数(filter_fields="短视频引流")
    个人主页_店铺_橱窗订单数 = Flow.get_指定渠道订单数(filter_fields="个人主页&店铺&橱窗")
    抖音商城推荐订单数 = Flow.get_指定渠道订单数(filter_fields="抖音商城推荐")
    活动页渠道订单数 = Flow.get_指定渠道订单数(filter_fields="活动页")
    头条西瓜渠道订单数 = Flow.get_指定渠道订单数(filter_fields="头条西瓜")
    其他渠道订单数 = Flow.get_指定渠道订单数(filter_fields="其他")
    千川PC版渠道订单数 = Flow.get_指定渠道订单数(filter_fields="千川PC版")
    小店随心推渠道订单数 = Flow.get_指定渠道订单数(filter_fields="小店随心推")
    品牌广告渠道订单数 = Flow.get_指定渠道订单数(filter_fields="品牌广告")
    其他广告渠道订单数 = Flow.get_指定渠道订单数(filter_fields="其他广告")
    千川品牌广告渠道订单数 = Flow.get_指定渠道订单数(filter_fields="千川品牌广告")

    推荐feed渠道CVR = (推荐feed渠道订单数 / 推荐feed渠道PV).apply(lambda x: format(x, ".9%"))
    直播广场渠道CVR = (直播广场渠道订单数 / 直播广场渠道PV).apply(lambda x: format(x, ".9%"))
    同城渠道CVR = (同城渠道订单数 / 同城渠道PV).apply(lambda x: format(x, ".9%"))
    其他推荐场景渠道CVR = (其他推荐场景渠道订单数 / 其他推荐场景渠道PV).apply(lambda x: format(x, ".9%"))
    关注渠道CVR = (关注渠道订单数 / 关注渠道PV).apply(lambda x: format(x, ".9%"))
    搜索渠道CVR = (搜索渠道订单数 / 搜索渠道PV).apply(lambda x: format(x, ".9%"))
    短视频引流渠道CVR = (短视频引流渠道订单数 / 短视频引流渠道PV).apply(lambda x: format(x, ".9%"))
    个人主页_店铺_橱窗CVR = (个人主页_店铺_橱窗订单数 / 个人主页_店铺_橱窗PV).apply(lambda x: format(x, ".9%"))
    抖音商城推荐CVR = (抖音商城推荐订单数 / 抖音商城推荐PV).apply(lambda x: format(x, ".9%"))
    活动页渠道CVR = (活动页渠道订单数 / 活动页渠道PV).apply(lambda x: format(x, ".9%"))
    头条西瓜渠道CVR = (头条西瓜渠道订单数 / 头条西瓜渠道PV).apply(lambda x: format(x, ".9%"))
    其他渠道CVR = (其他渠道订单数 / 其他渠道PV).apply(lambda x: format(x, ".9%"))
    千川PC版渠道CVR = (千川PC版渠道订单数 / 千川PC版渠道PV).apply(lambda x: format(x, ".9%"))
    小店随心推渠道CVR = (小店随心推渠道订单数 / 小店随心推渠道PV).apply(lambda x: format(x, ".9%"))
    品牌广告渠道CVR = (品牌广告渠道订单数 / 品牌广告渠道PV).apply(lambda x: format(x, ".9%"))
    其他广告渠道CVR = (其他广告渠道订单数 / 其他广告渠道PV).apply(lambda x: format(x, ".9%"))
    千川品牌广告渠道CVR = (千川品牌广告渠道订单数 / 千川品牌广告渠道PV).apply(lambda x: format(x, ".9%"))
    title = {
        "PV": PV,
        "自然PV": 自然PV,
        "付费PV": 付费PV,
        "自然PV_率": 自然PV_率,
        "付费PV_率": 付费PV_率,
        "自然CVR": 自然CVR,
        "付费CVR": 付费CVR,
        "TTLCVR": TTLCVR,
        "自然订单数": 自然订单数,
        "付费订单数": 付费订单数,
        "TTL订单数": TTL订单数,
        "推荐feed渠道PV": 推荐feed渠道PV,
        "直播广场渠道PV": 直播广场渠道PV,
        "同城渠道PV": 同城渠道PV,
        "其他推荐场景渠道PV": 其他推荐场景渠道PV,
        "关注渠道PV": 关注渠道PV,
        "搜索渠道PV": 搜索渠道PV,
        "短视频引流渠道PV": 短视频引流渠道PV,
        "个人主页_店铺_橱窗PV": 个人主页_店铺_橱窗PV,
        "抖音商城推荐PV": 抖音商城推荐PV,
        "活动页渠道PV": 活动页渠道PV,
        "头条西瓜渠道PV": 头条西瓜渠道PV,
        "其他渠道PV": 其他渠道PV,
        "千川PC版渠道PV": 千川PC版渠道PV,
        "小店随心推渠道PV": 小店随心推渠道PV,
        "品牌广告渠道PV": 品牌广告渠道PV,
        "其他广告渠道PV": 其他广告渠道PV,
        "千川品牌广告渠道PV": 千川品牌广告渠道PV,
        "推荐feed渠道GMV": 推荐feed渠道GMV,
        "直播广场渠道GMV": 直播广场渠道GMV,
        "同城渠道GMV": 同城渠道GMV,
        "其他推荐场景渠道GMV": 其他推荐场景渠道GMV,
        "关注渠道GMV": 关注渠道GMV,
        "搜索渠道GMV": 搜索渠道GMV,
        "短视频引流渠道GMV": 短视频引流渠道GMV,
        "个人主页_店铺_橱窗GMV": 个人主页_店铺_橱窗GMV,
        "抖音商城推荐GMV": 抖音商城推荐GMV,
        "活动页渠道GMV": 活动页渠道GMV,
        "头条西瓜渠道GMV": 头条西瓜渠道GMV,
        "其他渠道GMV": 其他渠道GMV,
        "千川PC版渠道GMV": 千川PC版渠道GMV,
        "小店随心推渠道GMV": 小店随心推渠道GMV,
        "品牌广告渠道GMV": 品牌广告渠道GMV,
        "其他广告渠道GMV": 其他广告渠道GMV,
        "千川品牌广告渠道GMV": 千川品牌广告渠道GMV,
        "推荐feed渠道GPM": 推荐feed渠道GPM,
        "直播广场渠道GPM": 直播广场渠道GPM,
        "同城渠道GPM": 同城渠道GPM,
        "其他推荐场景渠道GPM": 其他推荐场景渠道GPM,
        "关注渠道GPM": 关注渠道GPM,
        "搜索渠道GPM": 搜索渠道GPM,
        "短视频引流渠道GPM": 短视频引流渠道GPM,
        "个人主页_店铺_橱窗GPM": 个人主页_店铺_橱窗GPM,
        "抖音商城推荐GPM": 抖音商城推荐GPM,
        "活动页渠道GPM": 活动页渠道GPM,
        "头条西瓜渠道GPM": 头条西瓜渠道GPM,
        "其他渠道GPM": 其他渠道GPM,
        "千川PC版渠道GPM": 千川PC版渠道GPM,
        "小店随心推渠道GPM": 小店随心推渠道GPM,
        "品牌广告渠道GPM": 品牌广告渠道GPM,
        "其他广告渠道GPM": 其他广告渠道GPM,
        "千川品牌广告渠道GPM": 千川品牌广告渠道GPM,
        "推荐feed渠道CVR": 推荐feed渠道CVR,
        "直播广场渠道CVR": 直播广场渠道CVR,
        "同城渠道CVR": 同城渠道CVR,
        "其他推荐场景渠道CVR": 其他推荐场景渠道CVR,
        "关注渠道CVR": 关注渠道CVR,
        "搜索渠道CVR": 搜索渠道CVR,
        "短视频引流渠道CVR": 短视频引流渠道CVR,
        "个人主页_店铺_橱窗CVR": 个人主页_店铺_橱窗CVR,
        "抖音商城推荐CVR": 抖音商城推荐CVR,
        "活动页渠道CVR": 活动页渠道CVR,
        "头条西瓜渠道CVR": 头条西瓜渠道CVR,
        "其他渠道CVR": 其他渠道CVR,
        "千川PC版渠道CVR": 千川PC版渠道CVR,
        "品牌广告渠道CVR": 品牌广告渠道CVR,
        "其他广告渠道CVR": 其他广告渠道CVR,
        "千川品牌广告渠道CVR": 千川品牌广告渠道CVR,
        "小店随心推渠道CVR": 小店随心推渠道CVR,
        "推荐feed渠道订单数": 推荐feed渠道订单数,
        "直播广场渠道订单数": 直播广场渠道订单数,
        "同城渠道订单数": 同城渠道订单数,
        "其他推荐场景渠道订单数": 其他推荐场景渠道订单数,
        "关注渠道订单数": 关注渠道订单数,
        "搜索渠道订单数": 搜索渠道订单数,
        "短视频引流渠道订单数": 短视频引流渠道订单数,
        "个人主页_店铺_橱窗订单数": 个人主页_店铺_橱窗订单数,
        "抖音商城推荐订单数": 抖音商城推荐订单数,
        "活动页渠道订单数": 活动页渠道订单数,
        "头条西瓜渠道订单数": 头条西瓜渠道订单数,
        "其他渠道订单数": 其他渠道订单数,
        "千川PC版渠道订单数": 千川PC版渠道订单数,
        "品牌广告渠道订单数": 品牌广告渠道订单数,
        "其他广告渠道订单数": 其他广告渠道订单数,
        "千川品牌广告渠道订单数": 千川品牌广告渠道订单数,
        "小店随心推渠道订单数": 小店随心推渠道订单数,
    }
    try:
        export = (
            pd.concat(
                title.values(),
                axis=1,
                keys=title.keys(),
            )
            .fillna(0)
            .replace("nan%", 0)
            .replace("inf%", 0)
            .replace("nan", 0)
            .replace("inf", 0)
            .sort_index()
        )
        export.sort_index().to_csv(f"{Flow.export_csv_floader}/【{Flow.account_name}】_流量底表1.csv", index_label=["日期"])
        logger.success(f"{Flow.export_csv_floader}/【{Flow.account_name}】_流量底表1.csv 保存成功")
    except Exception as e:
        print(e)
        logger.error(f"{Flow.export_csv_floader}/【{Flow.account_name}】_流量底表1.csv保存失败:\n{e}")


def merg_import(folder_path: str = Flow.export_csv_floader):
    """合并表格

    Args:
        folder_path (str, optional): _description_. Defaults to Daily.export_csv_floader.
    """
    logger.info("开始合并表格")
    output_file = f"{Flow.export_csv_floader}/{Flow.class_name}.xlsx"

    # 读取指定文件夹下的所有csv文件
    try:
        csv_files = [f for f in os.listdir(folder_path) if f.endswith(".csv")]

        # 如果文件不存在，创建一个新文件
        if not os.path.exists(output_file):
            pd.DataFrame().to_excel(output_file, index=False)

        # 将每个csv文件的数据作为sheet，并新增到整个文件中
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            for csv_file in csv_files:
                csv_path = os.path.join(folder_path, csv_file)
                temp_df = pd.read_csv(csv_path)
                temp_df.to_excel(writer, sheet_name=csv_file[:31], index=False)
        logger.success("合并表格完成")
    except Exception as e:
        logger.error(f"合并表格失败\n{e}")


if __name__ == "__main__":
    if IS_DEBUG:
        print("调试模式")
        Flow.export_import_csv()
        save1()
        save2()
        merg_import()
        print("需要留意日期是否完整,如果缺失需要手动补充日期填充0\n,全选表格,ctrl+g,选择空的,输入0,按ctrl+enter补充")
        input("按任意键退出")
    else:
        print("正常模式")
        Flow.export_import_csv()
        save1()
        save2()
        merg_import()
        print("需要留意日期是否完整,如果缺失需要手动补充日期填充0\n,全选表格,ctrl+g,选择空的,输入0,按ctrl+enter补充")
        input("按任意键退出")
