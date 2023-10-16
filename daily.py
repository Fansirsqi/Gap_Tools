import pandas as pd
import logging as log

log.basicConfig(level=log.DEBUG)


class Daily:
    # 读取文件夹下面的csv文件
    account_name = "YSL圣罗兰美妆送礼空间"
    start_date = "2023-08-01"
    end_date = "2023-09-30"
    baiku_date_type = "bizDate"
    csvPerfix = "scrm_dy_report_app_fxg_"
    df1 = pd.read_csv(f"csv/{csvPerfix}live_detail_day.csv")
    df2 = pd.read_csv(f"csv/{csvPerfix}live_list_details_traffic_time_day.csv")
    df3 = pd.read_csv(f"csv/{csvPerfix}live_growth_conversion_funnel_day.csv")
    df4 = pd.read_csv(f"csv/{csvPerfix}live_list_details_grouping_day.csv")
    df5 = pd.read_csv(f"csv/{csvPerfix}live_list_details_indicators_day.csv")
    df6 = pd.read_csv(f"csv/{csvPerfix}live_analysis_live_details_all_day.csv")
    dfs = {
        "live_detail_day": df1,
        "live_list_details_traffic_traffic_time_day": df2,
        "live_growth_conversion_funnel_day": df3,
        "live_list_details_grouping_day": df4,
        "live_list_details_indicators_day": df5,
        "live_analysis_live_details_all_day": df6,
    }
    for df_key, df in dfs.items():
        log.debug(f"import {df_key}")
        if df_key == "live_analysis_live_details_all_day":
            dfs[df_key] = df[(df["author_nick_name"] == account_name) & (df["biz_date"].between(start_date, end_date))].sort_values(by=["biz_date"])
        else:
            dfs[df_key] = df[(df["account_name"] == account_name) & (df[baiku_date_type].between(start_date, end_date))].sort_values(by=baiku_date_type)
    qianchuan_df = pd.read_csv("csv/scrm_ocean_daily.csv")
    直播间成交金额 = "roomTransactionAmount"
    观看人次 = "audience"
    成交人数 = "watchRate"
    观看人数 = "cumulativeAudience"
    直播间曝光人数 = "exposureCount"
    最高在线人数 = "highestOnlineUsers"
    直播时长_分钟 = "live_duration"
    平均在线人数 = "acu"
    商品点击人数 = "storeClickCount"
    消耗 = "stat_cost"

    def get_GMV(df=dfs["live_detail_day"], sum_filed=直播间成交金额, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].sum()

    def get_指定观看人次PV(df=dfs["live_list_details_traffic_traffic_time_day"], sum_filed=观看人次, date_type=baiku_date_type, flowChannel=None):
        if flowChannel is not None:
            df = df[df["flowChannel"] == flowChannel]
        return df.groupby(date_type)[sum_filed].sum()

    def get_成交人数(df=dfs["live_growth_conversion_funnel_day"], sum_filed=成交人数, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].sum()

    def get_观看人数(df=dfs["live_list_details_grouping_day"], sum_filed=观看人数, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].sum()

    def get_最高在线人数(df=dfs["live_list_details_indicators_day"], sum_filed=最高在线人数, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].sum()

    def get_平均在线人数(df=dfs["live_analysis_live_details_all_day"], sum_filed=[直播时长_分钟, 平均在线人数]):
        ndf1 = df.loc[:, ["biz_date", sum_filed[0], sum_filed[1]]]
        ndf1.loc[:, "TS"] = ndf1[sum_filed[0]] * ndf1[sum_filed[1]]
        ndf2 = ndf1.groupby("biz_date")[sum_filed[0]].sum()
        ndf1 = ndf1.groupby("biz_date")["TS"].sum()
        ndf3 = (ndf1 / ndf2).apply(lambda x: format(x, ".0f"))
        return ndf3

    def get_直播时长_分钟(df=dfs["live_analysis_live_details_all_day"], sum_filed=[直播时长_分钟]):
        return df.groupby("biz_date")[sum_filed[0]].sum().apply(lambda x: format(x, ".0f"))

    def get_外层CTR(df1=dfs["live_growth_conversion_funnel_day"], df2=dfs["live_list_details_grouping_day"], sum_filed=[观看人数, 直播间曝光人数], date_type=baiku_date_type):
        ndf1 = df2.groupby(date_type)[sum_filed[0]].sum()  #!!!
        ndf2 = df1.groupby(date_type)[sum_filed[1]].sum()  #!!!
        ndf3 = (ndf1 / ndf2).apply(lambda x: format(x, ".9%"))
        return ndf3

    def get_商品点击人数(df=dfs["live_growth_conversion_funnel_day"], sum_filed=[商品点击人数], date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed[0]].sum()

    def get_曝光人数(df=dfs["live_growth_conversion_funnel_day"], sum_filed=[直播间曝光人数], date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed[0]].sum()
    def get_投放金额(df=qianchuan_df,sum_filed=[消耗]):
        pass


GMV_ACH_直播期间成交金额_ = Daily.get_GMV()
GMV_ACH_直播间成交金额_ = GMV_ACH_直播期间成交金额_
Campaign = pd.Series(0, index=GMV_ACH_直播期间成交金额_.index)
GMV_TGT = Campaign
GMV_ACH_占比 = pd.Series(0, index=GMV_ACH_直播期间成交金额_.index).apply(lambda x: format(x, ".2%"))
观看人次PV = Daily.get_指定观看人次PV()

GPM_直播期间成交金额_ = (GMV_ACH_直播期间成交金额_ * 1000 / 观看人次PV).apply(lambda x: format(x, ".9f"))
GPM_直播间成交金额_ = (GMV_ACH_直播间成交金额_ * 1000 / 观看人次PV).apply(lambda x: format(x, ".2f"))
成交人数 = Daily.get_成交人数()
客单价_直播期间成交金额_ = (GMV_ACH_直播期间成交金额_ / 成交人数).apply(lambda x: format(x, ".9f"))
客单价_直播间成交金额_ = (GMV_ACH_直播间成交金额_ / 成交人数).apply(lambda x: format(x, ".2f"))
观看人数UV = Daily.get_观看人数()
UV价值_直播期间成交金额_ = (GMV_ACH_直播期间成交金额_ / 观看人数UV).apply(lambda x: format(x, ".9f"))
UV价值_直播间成交金额_ = (GMV_ACH_直播间成交金额_ / 观看人数UV).apply(lambda x: format(x, ".2f"))
自然PV = Daily.get_指定观看人次PV(flowChannel="自然流量")
自然PV占比 = (自然PV / 观看人次PV).apply(lambda x: format(x, ".9%"))
付费PV = Daily.get_指定观看人次PV(flowChannel="付费流量")
付费PV占比 = (付费PV / 观看人次PV).apply(lambda x: format(x, ".9%"))
最高在线人数 = Daily.get_最高在线人数()
平均在线人数 = Daily.get_平均在线人数()
直播时长 = Daily.get_直播时长_分钟()
外层CTR = Daily.get_外层CTR()
商品点击人数 = Daily.get_商品点击人数()
商品点击率 = (商品点击人数 / 观看人数UV).apply(lambda x: format(x, ".9%"))
点击成交转化率 = (成交人数 / 商品点击人数).apply(lambda x: format(x, ".9%"))
观看转化率 = (成交人数 / 观看人数UV).apply(lambda x: format(x, ".9%"))
曝光人数 = Daily.get_曝光人数()
观看人次 = 观看人次PV
观看总人数 = 观看人数UV
print(曝光人数)
