import pandas as pd

input("欢迎使[店播流量]用取数工具")
print("需要用到的底表如下")
print("scrm_dy_report_app_fxg_live_list_details_traffic_time_day.csv")
print("scrm_ocean_daily.csv")
print("scrm_dy_report_app_fxg_live_detail_day.csv")
input("确认文件夹下面有以上文件按任意按键开始执行")


class 店播流量_sheet1:
    baiku_csv_prefix = "scrm_dy_report_"
    account_name = input("请输入account_name：")
    s_date = "2023-08-01"
    e_date = "2023-09-30"
    baiku_csv_file = f"{baiku_csv_prefix}app_fxg_live_list_details_traffic_time_day.csv"
    baiku_df = pd.read_csv(baiku_csv_file)
    not_channelName = "整体"
    not_flowChannel = "全域推广"
    PV = "audience"
    订单数 = "transactionOrderNumber"
    GMV = "dealAmount"

    def get_PV(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, filed=PV, not_channelName=not_channelName, not_flowChannel=not_flowChannel):
        df_new = df[(df["account_name"] == account_name) & (df["channelName"] != not_channelName) & (df["flowChannel"] != not_flowChannel)]
        df_new1 = df_new[df_new["bizDate"].between(start_date, end_date)]
        return df_new1.groupby(["bizDate"])[filed].sum().fillna(0)

    def get_自然PV(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, filed=PV, not_channelName=not_channelName):
        df_new = df[(df["account_name"] == account_name) & (df["flowChannel"] == "自然流量") & (df["channelName"] != not_channelName)]
        df_new1 = df_new[df_new["bizDate"].between(start_date, end_date)]
        return df_new1.groupby(["bizDate"])[filed].sum().fillna(0)

    def get_付费PV(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, filed=PV, not_channelName=not_channelName):
        df_new = df[(df["account_name"] == account_name) & (df["flowChannel"] == "付费流量") & (df["channelName"] != not_channelName)]
        df_new1 = df_new[df_new["bizDate"].between(start_date, end_date)]
        return df_new1.groupby(["bizDate"])[filed].sum().fillna(0)

    def get_自然订单数(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, filed=订单数, not_channelName=not_channelName):
        df_new = df[(df["account_name"] == account_name) & (df["flowChannel"] == "自然流量") & (df["channelName"] != not_channelName)]
        df_new1 = df_new[df_new["bizDate"].between(start_date, end_date)]
        return df_new1.groupby(["bizDate"])[filed].sum().fillna(0)

    def get_付费订单数(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, filed=订单数, not_channelName=not_channelName):
        df_new = df[(df["account_name"] == account_name) & (df["flowChannel"] == "付费流量") & (df["channelName"] != not_channelName)]
        df_new1 = df_new[df_new["bizDate"].between(start_date, end_date)]
        return df_new1.groupby(["bizDate"])[filed].sum().fillna(0)

    def get_指定渠道PV(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, filed=PV, not_flowChannel=not_flowChannel, filter_fields=None):
        """
        filter_fields = 千川PC版
        推荐feed
        关注
        其他
        搜索
        品牌广告
        小店随心推
        个人主页&店铺&橱窗
        抖音商城推荐
        直播广场
        短视频引流
        其他推荐场景
        活动页
        同城
        头条西瓜
        其他广告
        千川品牌广告
        """
        df_new = df[(df["account_name"] == account_name) & (df["flowChannel"] != not_flowChannel) & (df["channelName"] == filter_fields)]
        df_new1 = df_new[df_new["bizDate"].between(start_date, end_date)]
        return df_new1.groupby(["bizDate"])[filed].sum().fillna(0)

    def get_指定渠道GMV(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, filed=GMV, not_flowChannel=not_flowChannel, filter_fields=None):
        """
        filter_fields = 千川PC版
        推荐feed
        关注
        其他
        搜索
        品牌广告
        小店随心推
        个人主页&店铺&橱窗
        抖音商城推荐
        直播广场
        短视频引流
        其他推荐场景
        活动页
        同城
        头条西瓜
        其他广告
        千川品牌广告
        """

        df_new = df[(df["account_name"] == account_name) & (df["flowChannel"] != not_flowChannel) & (df["channelName"] == filter_fields)]
        df_new1 = df_new[df_new["bizDate"].between(start_date, end_date)]
        return df_new1.groupby(["bizDate"])[filed].sum().fillna(0)

    def get_指定渠道订单数(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, filed=订单数, not_flowChannel=not_flowChannel, filter_fields=None):
        """
        filter_fields = 千川PC版
        推荐feed
        关注
        其他
        搜索
        品牌广告
        小店随心推
        个人主页&店铺&橱窗
        抖音商城推荐
        直播广场
        短视频引流
        其他推荐场景
        活动页
        同城
        头条西瓜
        其他广告
        千川品牌广告
        """
        df_new = df[(df["account_name"] == account_name) & (df["flowChannel"] != not_flowChannel) & (df["channelName"] == filter_fields)]
        df_new1 = df_new[df_new["bizDate"].between(start_date, end_date)]
        return df_new1.groupby(["bizDate"])[filed].sum().fillna(0)


class 店播流量_sheet2:
    baiku_csv_prefix = "scrm_dy_report_"
    account_name = input("请输入account_name：")
    s_date = "2023-08-01"
    e_date = "2023-09-30"
    baiku_csv_file = f"{baiku_csv_prefix}app_fxg_live_detail_day.csv"
    baiku_df = pd.read_csv(baiku_csv_file)
    qianchuan_csv_file = "scrm_ocean_daily.csv"
    qianchuan_df = pd.read_csv(qianchuan_csv_file)
    not_channelName = "整体"
    not_flowChannel = "全域推广"
    baiku_date_type = "bizDate"
    qianchuan_date_type = "date"
    消耗 = "stat_cost"
    成交订单金额 = "pay_order_amount"
    直播间成金额 = "roomTransactionAmount"

    def get_投放金额(df=qianchuan_df, account_name=account_name, start_date=s_date, end_date=e_date, date_filed=qianchuan_date_type, filed=消耗):
        df_new = df[(df["account_name"] == account_name) & (df["marketing_goal"] == "LIVE_PROM_GOODS")]
        df_new1 = df_new[df_new[date_filed].between(start_date, end_date)]
        return df_new1.groupby([date_filed])[filed].sum().fillna(0)

    def get_投放转化金额(df=qianchuan_df, account_name=account_name, start_date=s_date, end_date=e_date, date_filed=qianchuan_date_type, filed=成交订单金额):
        df_new = df[(df["account_name"] == account_name) & (df["marketing_goal"] == "LIVE_PROM_GOODS")]
        df_new1 = df_new[df_new[date_filed].between(start_date, end_date)]
        return df_new1.groupby([date_filed])[filed].sum().fillna(0)

    def get_直播间成交金额(df=baiku_df, account_name=account_name, start_date=s_date, end_date=e_date, date_filed=baiku_date_type, filed=直播间成金额):
        df_new = df[(df["account_name"] == account_name)]
        df_new1 = df_new[df_new[date_filed].between(start_date, end_date)]
        return df_new1.groupby([date_filed])[filed].sum().fillna(0)


def save_sheet2():
    投放金额 = 店播流量_sheet2.get_投放金额()
    投放转化金额 = 店播流量_sheet2.get_投放转化金额()
    直播间成交金额 = 店播流量_sheet2.get_直播间成交金额()
    Take_Rate_直播期间 = (投放金额 / 直播间成交金额).apply(lambda x: format(x, ".9%"))
    Take_Rate_直播间 = (投放金额 / 直播间成交金额).apply(lambda x: format(x, ".2%"))
    Media_Contribution_直播期间 = (投放转化金额 / 直播间成交金额).apply(lambda x: format(x, ".9%"))
    Media_Contribution_直播间 = Media_Contribution_直播期间
    TTLROI_直播期间 = (直播间成交金额 / 投放金额).apply(lambda x: format(x, ".10"))
    TTLROI_直播间 = TTLROI_直播期间
    投放ROI = (投放转化金额 / 投放金额).apply(lambda x: format(x, ".10"))
    try:
        export = (
            pd.concat(
                [投放金额, 投放转化金额, Take_Rate_直播期间, Take_Rate_直播间, Media_Contribution_直播期间, Media_Contribution_直播间, TTLROI_直播期间, TTLROI_直播间, 投放ROI],
                axis=1,
                keys=[
                    "投放金额_元",
                    "投放转化金额",
                    "Take_Rate_直播期间成交金额",
                    "Take_Rate_直播间成交金额",
                    "Media_Contribution_直播期间成交金额",
                    "Media_Contribution_直播间成交金额",
                    "TTL_ROI_直播期间成交金额",
                    "TTL_ROI_直播间成交金额",
                    "投放ROI",
                ],
            )
            .fillna(0)
            .replace("nan%", 0)
            .replace("nan", 0)
        )
        export.sort_index().to_csv(f"{店播流量_sheet2.account_name}_[流量表2].csv", encoding="utf-8", index=True)
        print(f"{店播流量_sheet2.account_name}_[流量表2].csv ==> 保存成功")
    except Exception as e:
        print(e)


def save_sheet1():
    PV = 店播流量_sheet1.get_PV()
    自然PV = 店播流量_sheet1.get_自然PV()
    自然PV_率 = (自然PV / PV).apply(lambda x: format(x, ".9%")).dropna().replace(0, "0.00%")
    付费PV = 店播流量_sheet1.get_付费PV()
    付费PV_率 = (付费PV / PV).apply(lambda x: format(x, ".9%"))
    自然订单数 = 店播流量_sheet1.get_自然订单数()
    付费订单数 = 店播流量_sheet1.get_付费订单数()
    TTL订单数 = 自然订单数 + 付费订单数
    自然CVR = (自然订单数 / 自然PV).apply(lambda x: format(x, ".9%"))
    付费CVR = (付费订单数 / 付费PV).apply(lambda x: format(x, ".9%"))
    TTLCVR = (TTL订单数 / PV).apply(lambda x: format(x, ".9%"))
    推荐feed渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="推荐feed")
    直播广场渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="直播广场")
    同城渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="同城")
    其他推荐场景渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="其他推荐场景")
    关注渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="关注")
    搜索渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="搜索")
    短视频引流渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="短视频引流")
    个人主页_店铺_橱窗PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="个人主页&店铺&橱窗")
    抖音商城推荐PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="抖音商城推荐")
    活动页渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="活动页")
    头条西瓜渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="头条西瓜")
    其他渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="其他")
    千川PC版渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="千川PC版")
    小店随心推渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="小店随心推")
    品牌广告渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="品牌广告")
    其他广告渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="其他广告")
    千川品牌广告渠道PV = 店播流量_sheet1.get_指定渠道PV(filter_fields="千川品牌广告")

    推荐feed渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="推荐feed")
    直播广场渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="直播广场")
    同城渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="同城")
    其他推荐场景渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="其他推荐场景")
    关注渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="关注")
    搜索渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="搜索")
    短视频引流渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="短视频引流")
    个人主页_店铺_橱窗GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="个人主页&店铺&橱窗")
    抖音商城推荐GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="抖音商城推荐")
    活动页渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="活动页")
    头条西瓜渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="头条西瓜")
    其他渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="其他")
    千川PC版渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="千川PC版")
    小店随心推渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="小店随心推")
    品牌广告渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="品牌广告")
    其他广告渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="其他广告")
    千川品牌广告渠道GMV = 店播流量_sheet1.get_指定渠道GMV(filter_fields="千川品牌广告")

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

    推荐feed渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="推荐feed")
    直播广场渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="直播广场")
    同城渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="同城")
    其他推荐场景渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="其他推荐场景")
    关注渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="关注")
    搜索渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="搜索")
    短视频引流渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="短视频引流")
    个人主页_店铺_橱窗订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="个人主页&店铺&橱窗")
    抖音商城推荐订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="抖音商城推荐")
    活动页渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="活动页")
    头条西瓜渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="头条西瓜")
    其他渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="其他")
    千川PC版渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="千川PC版")
    小店随心推渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="小店随心推")
    品牌广告渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="品牌广告")
    其他广告渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="其他广告")
    千川品牌广告渠道订单数 = 店播流量_sheet1.get_指定渠道订单数(filter_fields="千川品牌广告")

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
    try:
        export = (
            pd.concat(
                [
                    PV,
                    自然PV,
                    付费PV,
                    自然PV_率,
                    付费PV_率,
                    自然CVR,
                    付费CVR,
                    TTLCVR,
                    自然订单数,
                    付费订单数,
                    TTL订单数,
                    推荐feed渠道PV,
                    直播广场渠道PV,
                    同城渠道PV,
                    其他推荐场景渠道PV,
                    关注渠道PV,
                    搜索渠道PV,
                    短视频引流渠道PV,
                    个人主页_店铺_橱窗PV,
                    抖音商城推荐PV,
                    活动页渠道PV,
                    头条西瓜渠道PV,
                    其他渠道PV,
                    千川PC版渠道PV,
                    小店随心推渠道PV,
                    品牌广告渠道PV,
                    其他广告渠道PV,
                    千川品牌广告渠道PV,
                    推荐feed渠道GMV,
                    直播广场渠道GMV,
                    同城渠道GMV,
                    其他推荐场景渠道GMV,
                    关注渠道GMV,
                    搜索渠道GMV,
                    短视频引流渠道GMV,
                    个人主页_店铺_橱窗GMV,
                    抖音商城推荐GMV,
                    活动页渠道GMV,
                    头条西瓜渠道GMV,
                    其他渠道GMV,
                    千川PC版渠道GMV,
                    小店随心推渠道GMV,
                    品牌广告渠道GMV,
                    其他广告渠道GMV,
                    千川品牌广告渠道GMV,
                    推荐feed渠道GPM,
                    直播广场渠道GPM,
                    同城渠道GPM,
                    其他推荐场景渠道GPM,
                    关注渠道GPM,
                    搜索渠道GPM,
                    短视频引流渠道GPM,
                    个人主页_店铺_橱窗GPM,
                    抖音商城推荐GPM,
                    活动页渠道GPM,
                    头条西瓜渠道GPM,
                    其他渠道GPM,
                    千川PC版渠道GPM,
                    小店随心推渠道GPM,
                    品牌广告渠道GPM,
                    其他广告渠道GPM,
                    千川品牌广告渠道GPM,
                    推荐feed渠道CVR,
                    直播广场渠道CVR,
                    同城渠道CVR,
                    其他推荐场景渠道CVR,
                    关注渠道CVR,
                    搜索渠道CVR,
                    短视频引流渠道CVR,
                    个人主页_店铺_橱窗CVR,
                    抖音商城推荐CVR,
                    活动页渠道CVR,
                    头条西瓜渠道CVR,
                    其他渠道CVR,
                    千川PC版渠道CVR,
                    品牌广告渠道CVR,
                    其他广告渠道CVR,
                    千川品牌广告渠道CVR,
                    小店随心推渠道CVR,
                    推荐feed渠道订单数,
                    直播广场渠道订单数,
                    同城渠道订单数,
                    其他推荐场景渠道订单数,
                    关注渠道订单数,
                    搜索渠道订单数,
                    短视频引流渠道订单数,
                    个人主页_店铺_橱窗订单数,
                    抖音商城推荐订单数,
                    活动页渠道订单数,
                    头条西瓜渠道订单数,
                    其他渠道订单数,
                    千川PC版渠道订单数,
                    品牌广告渠道订单数,
                    其他广告渠道订单数,
                    千川品牌广告渠道订单数,
                    小店随心推渠道订单数,
                ],
                axis=1,
                keys=[
                    "PV",
                    "自然PV",
                    "付费PV",
                    "自然PV_率",
                    "付费PV_率",
                    "自然CVR",
                    "付费CVR",
                    "TTLCVR",
                    "自然订单数",
                    "付费订单数",
                    "TTL订单数",
                    "推荐feed渠道PV",
                    "直播广场渠道PV",
                    "同城渠道PV",
                    "其他推荐场景渠道PV",
                    "关注渠道PV",
                    "搜索渠道PV",
                    "短视频引流渠道PV",
                    "个人主页_店铺_橱窗PV",
                    "抖音商城推荐PV",
                    "活动页渠道PV",
                    "头条西瓜渠道PV",
                    "其他渠道PV",
                    "千川PC版渠道PV",
                    "小店随心推渠道PV",
                    "品牌广告渠道PV",
                    "其他广告渠道PV",
                    "千川品牌广告渠道PV",
                    "推荐feed渠道GMV",
                    "直播广场渠道GMV",
                    "同城渠道GMV",
                    "其他推荐场景渠道GMV",
                    "关注渠道GMV",
                    "搜索渠道GMV",
                    "短视频引流渠道GMV",
                    "个人主页_店铺_橱窗GMV",
                    "抖音商城推荐GMV",
                    "活动页渠道GMV",
                    "头条西瓜渠道GMV",
                    "其他渠道GMV",
                    "千川PC版渠道GMV",
                    "小店随心推渠道GMV",
                    "品牌广告渠道GMV",
                    "其他广告渠道GMV",
                    "千川品牌广告渠道GMV",
                    "推荐feed渠道GPM",
                    "直播广场渠道GPM",
                    "同城渠道GPM",
                    "其他推荐场景渠道GPM",
                    "关注渠道GPM",
                    "搜索渠道GPM",
                    "短视频引流渠道GPM",
                    "个人主页_店铺_橱窗GPM",
                    "抖音商城推荐GPM",
                    "活动页渠道GPM",
                    "头条西瓜渠道GPM",
                    "其他渠道GPM",
                    "千川PC版渠道GPM",
                    "小店随心推渠道GPM",
                    "品牌广告渠道GPM",
                    "其他广告渠道GPM",
                    "千川品牌广告渠道GPM",
                    "推荐feed渠道CVR",
                    "直播广场渠道CVR",
                    "同城渠道CVR",
                    "其他推荐场景渠道CVR",
                    "关注渠道CVR",
                    "搜索渠道CVR",
                    "短视频引流渠道CVR",
                    "个人主页_店铺_橱窗CVR",
                    "抖音商城推荐CVR",
                    "活动页渠道CVR",
                    "头条西瓜渠道CVR",
                    "其他渠道CVR",
                    "千川PC版渠道CVR",
                    "品牌广告渠道CVR",
                    "其他广告渠道CVR",
                    "千川品牌广告渠道CVR",
                    "小店随心推渠道CVR",
                    "推荐feed渠道订单数",
                    "直播广场渠道订单数",
                    "同城渠道订单数",
                    "其他推荐场景渠道订单数",
                    "关注渠道订单数",
                    "搜索渠道订单数",
                    "短视频引流渠道订单数",
                    "个人主页_店铺_橱窗订单数",
                    "抖音商城推荐订单数",
                    "活动页渠道订单数",
                    "头条西瓜渠道订单数",
                    "其他渠道订单数",
                    "千川PC版渠道订单数",
                    "品牌广告渠道订单数",
                    "其他广告渠道订单数",
                    "千川品牌广告渠道订单数",
                    "小店随心推渠道订单数",
                ],
            )
            .fillna(0)
            .replace("nan%", 0)
            .sort_index()
        )
        export.to_csv(f"{店播流量_sheet1.account_name}_[流量表1].csv")
        print(f"{店播流量_sheet1.account_name}_[流量表1].csv ==> 保存成功")
    except Exception as e:
        print(e)


if __name__ == "__main__":
    save_sheet1()
    save_sheet2()
    print("需要留意日期是否完整,如果缺失需要手动补充日期填充0\n,全选表格,ctrl+g,选择空的,输入0,按ctrl+enter补充")
    input("按任意键退出")
