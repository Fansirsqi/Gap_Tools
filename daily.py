import os
import pandas as pd
import logging as log

log.basicConfig(level=log.DEBUG, filename="daily_log.txt", filemode="a", encoding="utf-8", format="%(asctime)s - %(levelname)s - %(message)s")


class Daily:
    # 读取文件夹下面的csv文件
    current_time = str(pd.Timestamp.now())

    log.info("-" * 30 + "| " + current_time + " |" + "-" * 30)
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
    df7 = pd.read_csv("csv/scrm_ocean_daily.csv")
    df8 = pd.read_csv(f"csv/{csvPerfix}live_detail_core_interaction_key_day.csv")
    df9 = pd.read_csv(f"csv/{csvPerfix}live_crowd_data_day.csv")
    df10 = pd.read_csv(f"csv/{csvPerfix}live_detail_person_day.csv")
    df11 = pd.read_csv(f"csv/{csvPerfix}live_list_details_download_day.csv")
    df12 = pd.read_csv(f"csv/{csvPerfix}fly_live_detail_first_prchase_day.csv")
    dfs = {
        "live_detail_day": df1,
        "live_list_details_traffic_traffic_time_day": df2,
        "live_growth_conversion_funnel_day": df3,
        "live_list_details_grouping_day": df4,
        "live_list_details_indicators_day": df5,
        "live_analysis_live_details_all_day": df6,
        "scrm_ocean_daily": df7,
        "live_detail_core_interaction_key_day": df8,
        "live_crowd_data_day": df9,
        "live_detail_person_day": df10,
        "live_list_details_download_day": df11,
        "fly_live_detail_first_prchase_day": df12,
    }
    log.debug("init data import ...")
    for df_key, df in dfs.items():
        log.debug(f"import {df_key}")
        if df_key == "live_analysis_live_details_all_day":
            log.debug("导入百库/ODP数据")
            dfs[df_key] = df[(df["author_nick_name"] == account_name) & (df["biz_date"].between(start_date, end_date))].sort_values(by=["biz_date"])
        elif df_key == "scrm_ocean_daily":
            log.debug("导入千川数据")
            dfs[df_key] = df[(df["account_name"] == account_name) & (df["marketing_goal"] == "LIVE_PROM_GOODS") & (df["date"].between(start_date, end_date))].sort_values(by=["date"])
        elif df_key == "fly_live_detail_first_prchase_day":
            if "圣罗兰" in account_name:
                account_name = "YSL圣罗兰美妆官方旗舰店"
            elif "兰蔻" in account_name:
                account_name = "兰蔻LANCOME官方旗舰店"
            dfs[df_key] = df[(df["store_name"] == account_name) & (df["date"].between(start_date, end_date)) & (df["account_type"] == "渠道账号")].sort_values(by="date")
        elif df_key == "live_list_details_traffic_traffic_time_day":  # 此处需要过滤部分条件
            dfs[df_key] = df[
                (df["account_name"] == account_name)
                & (df["flowChannel"] != "整体")
                & (df["flowChannel"] != "全域推广")
                & (df["channelName"] != "整体")
                & (df[baiku_date_type].between(start_date, end_date))
            ].sort_values(by=baiku_date_type)
        else:
            log.debug("导入百库/ODP数据")
            dfs[df_key] = df[(df["account_name"] == account_name) & (df[baiku_date_type].between(start_date, end_date))].sort_values(by=baiku_date_type)
    log.debug("data import success !")

    def export_import_csv(dfs=dfs):
        """导出csv文件"""
        export_folder = "./export_日报"
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
        for df_key, df in dfs.items():
            df.to_csv(f"{export_folder}/{df_key}.csv")

    直播间成交金额 = "roomTransactionAmount"
    粉丝成单占比 = "old_fans_pay_ucnt_ratio"
    观看人次 = "audience"
    成交人数 = "watchRate"
    观看人数 = "cumulativeAudience"
    直播间曝光人数 = "exposureCount"
    最高在线人数 = "highestOnlineUsers"
    直播时长_分钟 = "live_duration"
    平均在线人数 = "acu"
    商品点击人数 = "storeClickCount"
    消耗 = "stat_cost"
    成交订单金额 = "pay_order_amount"
    新增粉丝数 = "newFans"
    直播间观看人数 = "audienceStudio"
    粉丝占比 = "fans"
    成交人群分析_粉丝维度_粉丝人数 = "tranAnalysis_fans_fansCount"
    评论次数 = "commentNumber"
    人均观看时长_秒 = "averageViewingTime"

    def get_GMV(df=dfs["live_detail_day"], sum_filed=直播间成交金额, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].sum()

    def get_指定观看人次PV(df=dfs["live_list_details_traffic_traffic_time_day"], sum_filed=观看人次, date_type=baiku_date_type, flowChannel=None):
        # return df
        if flowChannel is not None:
            df = df[df["flowChannel"] == flowChannel]
        return df.groupby(date_type)[sum_filed].sum()

    def get_成交人数(df=dfs["live_growth_conversion_funnel_day"], sum_filed=成交人数, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].sum()

    def get_观看人数(df=dfs["live_list_details_grouping_day"], sum_filed=观看人数, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].sum()

    def get_最高在线人数(df=dfs["live_list_details_indicators_day"], sum_filed=最高在线人数, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].max()

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

    def get_投放金额(df=dfs["scrm_ocean_daily"], sum_filed=[消耗]):
        return df.groupby("date")[sum_filed[0]].sum()

    def get_投放转化金额(df=dfs["scrm_ocean_daily"], sum_filed=[成交订单金额]):
        return df.groupby("date")[sum_filed[0]].sum()

    def get_新增粉丝数(df=dfs["live_detail_core_interaction_key_day"], sum_filed=[新增粉丝数], date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed[0]].sum()

    def get_访问粉丝数(df=dfs["live_crowd_data_day"], sum_filed=[直播间观看人数, 粉丝占比], date_type=baiku_date_type):
        ndf1 = df.loc[:, [date_type, sum_filed[0], sum_filed[1]]]
        ndf1.loc[:, "访问粉丝数"] = ndf1[sum_filed[0]] * ndf1[sum_filed[1]] * 0.01
        ndf1 = ndf1.groupby(date_type)["访问粉丝数"].sum().apply(lambda x: format(x, ".2f")).astype("float64")
        return ndf1

    def get_购买粉丝数(df=dfs["live_detail_person_day"], sum_filed=[成交人群分析_粉丝维度_粉丝人数], date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed[0]].sum()

    def get_评论次数(df=dfs["live_list_details_download_day"], sum_filed=[评论次数], date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed[0]].sum()

    def get_人均观看时长(df1=dfs["live_detail_core_interaction_key_day"], df2=dfs["live_list_details_grouping_day"], sum_filed=[人均观看时长_秒, 观看人数], date_type=baiku_date_type):
        ndf1 = df1.set_index("studioId")[[sum_filed[0]]]
        ndf2 = df2.set_index("studioId")[[sum_filed[1], date_type]]
        ndfx = Daily.get_观看人数()
        ndf3 = pd.concat([ndf1, ndf2], axis=1)
        ndf3.loc[:, "TS"] = ndf3[sum_filed[0]] * ndf3[sum_filed[1]]
        ndf3 = (ndf3.groupby(date_type)["TS"].sum() / ndfx).apply(lambda x: format(x, ".0f")).astype("int64")
        return ndf3

    def get_粉丝成交GMV(df1=dfs["live_detail_day"], df2=dfs["fly_live_detail_first_prchase_day"], sum_filed1=直播间成交金额, sum_filed2=粉丝成单占比):
        ndf1 = df1.set_index("studioId")[[sum_filed1]]
        ndf2 = df2.set_index("live_room_id")[[sum_filed2, "date"]]
        result = pd.concat([ndf1, ndf2], axis=1)
        result["粉丝成交GMV"] = result[sum_filed1] * result[sum_filed2]
        result = result.groupby("date")["粉丝成交GMV"].sum().apply(lambda x: format(x, ".2f"))
        return result.astype("float64")

    def get_GPM(df1=dfs["live_detail_day"], df2=dfs["live_list_details_traffic_traffic_time_day"]):
        pass


def save_daily():
    GMV_ACH_直播期间成交金额_ = Daily.get_GMV()
    GMV_ACH_直播间成交金额_ = GMV_ACH_直播期间成交金额_
    Campaign = pd.Series(0, index=GMV_ACH_直播期间成交金额_.index)
    GMV_TGT = Campaign
    GMV_ACH占比 = pd.Series(0, index=GMV_ACH_直播期间成交金额_.index).apply(lambda x: format(x, ".2%"))
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
    投放金额_元_ = Daily.get_投放金额()
    投放转化金额 = Daily.get_投放转化金额()
    Take_Rate_直播期间成交金额_ = (投放金额_元_ / GMV_ACH_直播期间成交金额_).apply(lambda x: format(x, ".9%"))
    Take_Rate_直播间成交金额_ = (投放金额_元_ / GMV_ACH_直播间成交金额_).apply(lambda x: format(x, ".2%"))
    Media_Contribution_直播期间成交金额_ = (投放转化金额 / GMV_ACH_直播期间成交金额_).apply(lambda x: format(x, ".9%"))
    Media_Contribution_直播间成交金额_ = (投放转化金额 / GMV_ACH_直播间成交金额_).apply(lambda x: format(x, ".9%"))
    店播ROI_直播期间成交金额_ = (GMV_ACH_直播期间成交金额_ / 投放金额_元_).apply(lambda x: format(x, ".9f"))
    店播ROI_直播间成交金额_ = (GMV_ACH_直播间成交金额_ / 投放金额_元_).apply(lambda x: format(x, ".9f"))
    投放ROI = (投放转化金额 / 投放金额_元_).apply(lambda x: format(x, ".9f"))
    新增粉丝数 = Daily.get_新增粉丝数()
    转粉率 = (新增粉丝数 / 观看人数UV).apply(lambda x: format(x, ".9%"))
    访问粉丝数 = Daily.get_访问粉丝数()
    访问粉丝流量占比 = (访问粉丝数 / 观看人数UV).apply(lambda x: format(x, ".9%"))  # cute
    购买粉丝数 = Daily.get_购买粉丝数()
    购买粉丝数占比 = (购买粉丝数 / 成交人数).apply(lambda x: format(x, ".9%"))  # cute
    评论次数 = Daily.get_评论次数()
    互动率 = (评论次数 / (自然PV + 付费PV)).apply(lambda x: format(x, ".9%"))
    人均观看时长 = Daily.get_人均观看时长()
    粉丝成交GMV = Daily.get_粉丝成交GMV()
    粉丝成交金额占比 = (粉丝成交GMV / GMV_ACH_直播期间成交金额_).apply(lambda x: format(x, ".2%"))

    title = {
        "Campaign": Campaign,
        "GMV_TGT": GMV_TGT,
        "GMV_ACH_直播期间成交金额_": GMV_ACH_直播期间成交金额_,
        "GMV_ACH_直播间成交金额_": GMV_ACH_直播间成交金额_,
        "GMV_ACH占比": GMV_ACH占比,
        "GPM_直播期间成交金额_": GPM_直播期间成交金额_,
        "GPM_直播间成交金额_": GPM_直播间成交金额_,
        "客单价_直播期间成交金额_": 客单价_直播期间成交金额_,
        "客单价_直播间成交金额_": 客单价_直播间成交金额_,
        "UV价值_直播期间成交金额_": UV价值_直播期间成交金额_,
        "UV价值_直播间成交金额_": UV价值_直播间成交金额_,
        "成交人数": 成交人数,
        "观看人数UV": 观看人数UV,
        "观看人次PV": 观看人次PV,
        "自然PV占比": 自然PV占比,
        "自然PV": 自然PV,
        "付费PV占比": 付费PV占比,
        "付费PV": 付费PV,
        "最高在线人数": 最高在线人数,
        "平均在线人数": 平均在线人数,
        "直播时长": 直播时长,
        "外层CTR": 外层CTR,
        "商品点击率": 商品点击率,
        "点击成交转化率": 点击成交转化率,
        "观看转化率": 观看转化率,
        "曝光人数": 曝光人数,
        "观看人次": 观看人次,
        "观看总人数": 观看总人数,
        "商品点击人数": 商品点击人数,
        "投放金额_元_": 投放金额_元_,
        "投放转化金额": 投放转化金额,
        "Take_Rate_直播期间成交金额_": Take_Rate_直播期间成交金额_,
        "Take_Rate_直播间成交金额_": Take_Rate_直播间成交金额_,
        "Media_Contribution_直播期间成交金额_": Media_Contribution_直播期间成交金额_,
        "Media_Contribution_直播间成交金额_": Media_Contribution_直播间成交金额_,
        "店播ROI_直播期间成交金额_": 店播ROI_直播期间成交金额_,
        "店播ROI_直播间成交金额_": 店播ROI_直播间成交金额_,
        "投放ROI": 投放ROI,
        "新增粉丝数": 新增粉丝数,
        "转粉率": 转粉率,
        "访问粉丝数": 访问粉丝数,
        "访问粉丝流量占比": 访问粉丝流量占比,
        "购买粉丝数": 购买粉丝数,
        "购买粉丝数占比": 购买粉丝数占比,
        "评论次数": 评论次数,
        "互动率": 互动率,
        "人均观看时长": 人均观看时长,
        "粉丝成交金额占比": 粉丝成交金额占比,
        "粉丝成交GMV": 粉丝成交GMV,
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
        .sort_index()
    )
    try:
        export.to_csv("export_日报_by_底表.csv", index_label=["日期"])
        log.info("导出成功")
    except Exception as e:
        log.error(f"导出失败:{e}")
    Daily.export_import_csv()


def merg_import(folder_path: str = "export_日报", excel_file: str = None):
    """合并导入依赖底表

    Args:
        folder_path (str, optional): _description_. Defaults to "export_人群".
        excel_file (str, optional): _description_. Defaults to None.
    """
    # 读取指定文件夹下的所有csv文件
    csv_files = [f for f in os.listdir(folder_path) if f.endswith(".csv")]
    # 将每个csv文件的数据作为sheet，并新增到整个文件中
    with pd.ExcelWriter("new.xlsx", engine="openpyxl") as writer:
        with pd.ExcelFile(excel_file) as xls:
            sheet_names = xls.sheet_names
            for sheet_name in sheet_names:
                temp_df = xls.parse(sheet_name)
                temp_df.to_excel(writer, sheet_name=sheet_name, index=False)
        for csv_file in csv_files:
            csv_path = os.path.join(folder_path, csv_file)
            temp_df = pd.read_csv(csv_path)
            temp_df.to_excel(writer, sheet_name=csv_file[:31], index=False)


# save_daily()
# merg_import(excel_file='【抖音看板】日报_2023-08-01_2023-09-30.xlsx')
投放金额_元_ = Daily.get_投放金额()
print(投放金额_元_)
