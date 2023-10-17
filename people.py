import pandas as pd
import os
import logging as log

log.basicConfig(level=log.DEBUG)


class People:
    # 读取文件夹下面的csv文件
    account_name = "YSL圣罗兰美妆送礼空间"
    start_date = "2023-08-01"
    end_date = "2023-09-30"
    baiku_date_type = "bizDate"
    csvPerfix = "scrm_dy_report_app_fxg_"
    df1 = pd.read_csv(f"csv/{csvPerfix}live_number_people_covered_day.csv")
    df2 = pd.read_csv(f"csv/{csvPerfix}live_growth_conversion_funnel_day.csv")
    df3 = pd.read_csv(f"csv/{csvPerfix}live_detail_person_day.csv")
    df4 = pd.read_csv(f"csv/{csvPerfix}live_crowd_data_day.csv")
    df5 = pd.read_csv(f"csv/{csvPerfix}live_detail_day.csv")
    df6 = pd.read_csv(f"csv/{csvPerfix}fly_live_detail_first_prchase_day.csv")
    df7 = pd.read_csv(f"csv/{csvPerfix}live_list_details_grouping_day.csv")
    dfs = {
        "live_number_people_covered_day": df1,
        "live_growth_conversion_funnel_day": df2,
        "live_detail_person_day": df3,
        "live_crowd_data_day": df4,
        "live_detail_day": df5,
        "live_list_details_grouping_day": df7,
        "fly_live_detail_first_prchase_day": df6,
    }
    for df_key, df in dfs.items():
        log.debug(f"import {df_key}")
        if df_key == "fly_live_detail_first_prchase_day":
            if "圣罗兰" in account_name:
                account_name = "YSL圣罗兰美妆官方旗舰店"
            elif "兰蔻" in account_name:
                account_name = "兰蔻LANCOME官方旗舰店"
            dfs[df_key] = df[(df["store_name"] == account_name) & (df["date"].between(start_date, end_date)) & (df["account_type"] == "渠道账号")].sort_values(by="date")
        else:
            dfs[df_key] = df[(df["account_name"] == account_name) & (df[baiku_date_type].between(start_date, end_date))].sort_values(by=baiku_date_type)

    观看人数 = "imageCoverageCount"
    成交人数 = "watchRate"
    成交人群 = "gmv"
    直播间观看人数 = "audienceStudio"
    成交人群分析_新老客维度_首购人数 = "tranAnalysis_customers_firstBuyCount"
    成交人群分析_新老客维度_首购人数_粉丝占比 = "tranAnalysis_customers_firstBuyCount_fansRate"
    成交人群分析_新老客维度_复购人数 = "tranAnalysis_customers_secBuyCount"
    成交人群分析_新老客维度_复购人数_粉丝占比 = "tranAnalysis_customers_secBuyCount_fansRate"
    成交人群分析_粉丝维度_粉丝人数 = "tranAnalysis_fans_fansCount"
    成交人群分析_粉丝维度_粉丝人数_新增粉丝占比 = "tranAnalysis_fans_fansCount_newFansRate"
    成交人群分析_互动维度_有互动人数 = "tranAnalysis_interaction_intrCount"
    粉丝占比 = "fans"
    直播间成交金额 = "roomTransactionAmount"
    粉丝成单占比 = "old_fans_pay_ucnt_ratio"
    累计观看人数 = "cumulativeAudience"

    def export_import_csv(dfs=dfs):
        """导出csv文件"""
        export_folder = "./export_人群"
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
        for df_key, df in dfs.items():
            df.to_csv(f"{export_folder}/{df_key}.csv")

    def get_指定观看人数(df=dfs["live_number_people_covered_day"], sum_filed=观看人数, date_type=baiku_date_type, turnover=None, fans=None, interaction=None, customers=None, watch=None):
        """参数说明
        df: 数据框，默认为 People_df
        sum_filed: 需要汇总的列名，默认为 观看人数
        account_name: 账号名称 从类开头获取
        date_type: 日期类型，默认为 baiku_date_type
        turnover: 成交，默认为 None
        fans: 粉丝，默认为 None
        interaction: 互动，默认为 None
        customers: 新老客，默认为 None
        watch: 观看，默认为 None

        函数功能：根据传入的参数筛选数据，并对指定列进行汇总
        """
        ndf = df[
            (df["turnover"] == turnover)  # 成交
            & (df["fans"] == fans)  # 粉丝
            & (df["interaction"] == interaction)  # 互动
            & (df["customers"] == customers)  # 新老客
            & (df["watch"] == watch)  # 观看
        ]
        ndf1 = ndf.groupby([date_type])[sum_filed].sum()  # .sort_index().to_csv("demo.csv")
        return ndf1

    def get_成交人数(df=dfs["live_growth_conversion_funnel_day"], sum_filed=成交人数, date_type=baiku_date_type):
        ndf = df.groupby([date_type])[sum_filed].sum()
        return ndf

    def get_观看转化率(df=dfs["live_detail_person_day"], df2=dfs["live_crowd_data_day"], sum_filed1=成交人群, sum_filed2=直播间观看人数, date_type=baiku_date_type):
        ndf1 = df.groupby([date_type])[sum_filed1].sum()
        ndf2 = df2.groupby([date_type])[sum_filed2].sum()
        return (ndf1 / ndf2).apply(lambda x: format(x, ".9%"))

    def get_首购人数(df=dfs["live_detail_person_day"], sum_filed=成交人群分析_新老客维度_首购人数, date_type=baiku_date_type):
        return df.groupby([date_type])[sum_filed].sum()

    def get_首购粉丝数(df=dfs["live_detail_person_day"], sum_filed1=成交人群分析_新老客维度_首购人数, sum_filed2=成交人群分析_新老客维度_首购人数_粉丝占比, date_type=baiku_date_type):
        ndf1 = df[[sum_filed1, "id", date_type]]
        ndf2 = df[[sum_filed2, "id"]]
        result = pd.merge(ndf1, ndf2, on="id")
        result["首购粉丝"] = result[sum_filed1] * result[sum_filed2]
        result = result.groupby(date_type)["首购粉丝"].sum().apply(lambda x: format(x, ".0f"))
        return result.astype("int64")

    def get_复购粉丝数(df=dfs["live_detail_person_day"], sum_filed1=成交人群分析_新老客维度_复购人数, sum_filed2=成交人群分析_新老客维度_复购人数_粉丝占比, date_type=baiku_date_type):
        ndf1 = df.loc[:, [sum_filed1, date_type, sum_filed2, "id"]]
        ndf1.loc[:, "复购粉丝"] = ndf1[sum_filed1] * ndf1[sum_filed2]
        # result["复购粉丝"] = result[sum_filed1] * result[sum_filed2]
        return ndf1.groupby(date_type)["复购粉丝"].sum().apply(lambda x: format(x, ".0f")).astype("int64")

    def get_复购人数(df=dfs["live_detail_person_day"], sum_filed=成交人群分析_新老客维度_复购人数, date_type=baiku_date_type):
        return df.groupby([date_type])[sum_filed].sum()

    def get_粉丝成交人数(df=dfs["live_detail_person_day"], sum_filed=成交人群分析_粉丝维度_粉丝人数, date_type=baiku_date_type):
        return df.groupby([date_type])[sum_filed].sum()

    def get_粉丝观看人数(df=dfs["live_crowd_data_day"], sum_filed1=直播间观看人数, sum_filed2=粉丝占比, date_type=baiku_date_type):
        ndf1 = df[[sum_filed1, "id", date_type]]
        ndf2 = df[[sum_filed2, "id"]]
        result = pd.merge(ndf1, ndf2, on="id")
        result["粉丝观看人数"] = result[sum_filed1] * result[sum_filed2] * 0.01
        result = result.groupby(date_type)["粉丝观看人数"].sum().apply(lambda x: format(x, ".2f"))
        return result.astype("float64")

    def get_粉丝成交GMV(df1=dfs["live_detail_day"], df2=dfs["fly_live_detail_first_prchase_day"], sum_filed1=直播间成交金额, sum_filed2=粉丝成单占比):
        ndf1 = df1.set_index("studioId")[[sum_filed1]]
        ndf2 = df2.set_index("live_room_id")[[sum_filed2, "date"]]
        result = pd.concat([ndf1, ndf2], axis=1)
        result["粉丝成交GMV"] = result[sum_filed1] * result[sum_filed2]
        result = result.groupby("date")["粉丝成交GMV"].sum().apply(lambda x: format(x, ".2f"))
        return result.astype("float64")

    def get_累计观看人数(df=dfs["live_list_details_grouping_day"], sum_filed=累计观看人数, date_type=baiku_date_type):
        return df.groupby([date_type])[sum_filed].sum()

    def get_新粉成交人数(df=dfs["live_detail_person_day"], sum_filed1=成交人群分析_粉丝维度_粉丝人数, sum_filed2=成交人群分析_粉丝维度_粉丝人数_新增粉丝占比, date_type=baiku_date_type):
        ndf1 = df.loc[:, [sum_filed1, date_type, sum_filed2, "id"]]
        ndf1.loc[:, "新粉成交人数"] = ndf1[sum_filed1] * ndf1[sum_filed2]
        return ndf1.groupby(date_type)["新粉成交人数"].sum()

    def get_互动成交人数(df=dfs["live_detail_person_day"], sum_filed=成交人群分析_互动维度_有互动人数, date_type=baiku_date_type):
        return df.groupby(date_type)[sum_filed].sum()


def merg_import(folder_path: str = "export_人群", excel_file: str = None):
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


def save_people():
    观看人数 = People.get_指定观看人数(turnover="不限", fans="不限", interaction="不限", customers="不限", watch="不限")
    成交人数 = People.get_成交人数()
    观看转化率 = People.get_观看转化率()
    首购粉丝 = People.get_首购粉丝数()

    首购人数 = People.get_首购人数()
    首购粉丝占比 = (首购粉丝 / 首购人数).apply(lambda x: format(x, ".2%"))

    复购粉丝 = People.get_复购粉丝数()

    复购人数 = People.get_复购人数()

    复购粉丝占比 = (复购粉丝 / 复购人数).apply(lambda x: format(x, ".2%"))
    首购占比 = (首购人数 / 成交人数).apply(lambda x: format(x, ".9%"))
    复购占比 = (复购人数 / 成交人数).apply(lambda x: format(x, ".9%"))
    粉丝成交人数 = People.get_粉丝成交人数()

    粉丝观看人数 = People.get_粉丝观看人数()
    粉丝成交GMV = People.get_粉丝成交GMV()
    粉丝GPM = (粉丝成交GMV * 1000 / 粉丝观看人数).apply(lambda x: format(x, ".9f"))
    观看未购粉丝 = (粉丝观看人数 - 粉丝成交人数).apply(lambda x: format(x, ".2f"))
    累计观看人数 = People.get_累计观看人数()  # maybe do not need
    观看人数粉丝占比 = (粉丝观看人数 / 累计观看人数).apply(lambda x: format(x, ".2%"))
    成交人数粉丝占比 = (粉丝成交人数 / 成交人数).apply(lambda x: format(x, ".9%"))
    粉丝观看转化率 = (粉丝成交人数 / 粉丝观看人数).apply(lambda x: format(x, ".9%"))
    新粉观看人数 = People.get_指定观看人数(turnover="不限", fans="新粉丝", interaction="不限", customers="不限", watch="不限")
    观看人数新粉占比 = (新粉观看人数 / 观看人数).apply(lambda x: format(x, ".9%"))
    新粉成交人数 = People.get_新粉成交人数()
    粉丝成交人数新粉占比 = (新粉成交人数 / 粉丝成交人数).apply(lambda x: format(x, ".2%"))
    成交人数新粉占比 = (新粉成交人数 / 成交人数).apply(lambda x: format(x, ".9%"))
    新粉观看转化率 = (新粉成交人数 / 新粉观看人数).apply(lambda x: format(x, ".9%"))
    老粉观看人数 = (粉丝观看人数 - 新粉观看人数).apply(lambda x: format(x, ".2f")).astype("float64")
    观看人数老粉占比 = (老粉观看人数 / 观看人数).apply(lambda x: format(x, ".9%"))
    老粉成交人数 = (粉丝成交人数 - 新粉成交人数).apply(lambda x: format(x, ".0f")).astype("float64")
    成交人数老粉占比 = (老粉成交人数 / 成交人数).apply(lambda x: format(x, ".9%"))
    老粉观看转化率 = (老粉成交人数 / 老粉观看人数).apply(lambda x: format(x, ".9%"))
    互动成交人数 = People.get_互动成交人数()
    互动成交人数占比 = (互动成交人数 / 成交人数).apply(lambda x: format(x, ".9%"))

    # print(互动成交人数占比)

    title = {
        "观看人数": 观看人数,
        "成交人数": 成交人数,
        "观看转化率": 观看转化率,
        "首购粉丝": 首购粉丝,
        "首购粉丝占比": 首购粉丝占比,
        "复购粉丝": 复购粉丝,
        "复购粉丝占比": 复购粉丝占比,
        "首购人数": 首购人数,
        "首购占比": 首购占比,
        "复购人数": 复购人数,
        "复购占比": 复购占比,
        "观看未购粉丝": 观看未购粉丝,
        "粉丝成交GMV": 粉丝成交GMV,
        "粉丝GPM": 粉丝GPM,
        "粉丝观看人数": 粉丝观看人数,
        "观看人数粉丝占比": 观看人数粉丝占比,
        "粉丝成交人数": 粉丝成交人数,
        "成交人数粉丝占比": 成交人数粉丝占比,
        "粉丝观看转化率": 粉丝观看转化率,
        "新粉观看人数": 新粉观看人数,
        "观看人数新粉占比": 观看人数新粉占比,
        "新粉成交人数": 新粉成交人数,
        "粉丝成交人数新粉占比": 粉丝成交人数新粉占比,
        "成交人数新粉占比": 成交人数新粉占比,
        "新粉观看转化率": 新粉观看转化率,
        "老粉观看人数": 老粉观看人数,
        "观看人数老粉占比": 观看人数老粉占比,
        "老粉成交人数": 老粉成交人数,
        "成交人数老粉占比": 成交人数老粉占比,
        "老粉观看转化率": 老粉观看转化率,
        "互动成交人数": 互动成交人数,
        "互动成交人数占比": 互动成交人数占比,
    }
    export = (
        pd.concat(
            title.values(),
            axis=1,
            keys=title.keys(),
        )
        .fillna(0)
        .replace("nan%", 0)
        .replace("inf%", 0)
        .sort_index()
    )
    try:
        export.to_csv("export_人群_by_底表.csv", index_label=["日期"])
        log.info("导出成功")
    except Exception as e:
        log.error(f"导出失败:{e}")
    People.export_import_csv()


# save_people()
观看人数 = People.get_指定观看人数(turnover="不限", fans="不限", interaction="不限", customers="不限", watch="不限")
# GMV = People.get_粉丝成交GMV()
print(观看人数)
