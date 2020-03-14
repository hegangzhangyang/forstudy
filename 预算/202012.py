# -*-coding:utf-8-*-
import sqlite3

import pandas as pd
from lib.common import excelFormat


def consoleFormat():
    # 这两段代码设置输出时对齐
    pd.set_option('display.unicode.ambiguous_as_wide', True)
    pd.set_option('display.unicode.east_asian_width', True)
    # 设置最大显示列数为10
    pd.set_option('display.max_columns', 100)
    # 设置控制台现实不允许换行
    pd.set_option('expand_frame_repr', False)


if __name__ == '__main__':
    consoleFormat()
    df_budget = pd.read_excel("2020年工资明细工资明细.xlsx").fillna(0)

    df_budget.loc[df_budget["职务"] == "部长", "奖金"] = 11333
    df_budget.loc[df_budget["职务"] == "副部长", "奖金"] = 9066
    df_budget.loc[df_budget["职务"] == "部长助理", "奖金"] = 7253
    df_budget.loc[~df_budget["职务"].isin(["部长", "副部长", "部长助理"]), "奖金"] = df_budget["奖金系数"] * 1700

    df_budget["预支年薪"] = df_budget["预支年薪标准"]
    df_budget["岗位工资"] = df_budget["岗位工资系数（唐钢）"]*360
    df_budget["年功保留工资"] = df_budget["年功保留工资标准"]
    df_budget["竞业津贴"] = df_budget["竞业津贴标准"]
    df_budget["技术津贴"] = df_budget["技术津贴标准"]
    df_budget["回民补贴"] = df_budget["回民补贴标准"]
    df_budget["女工洗理费"] = df_budget["女工洗理费标准"]
    df_budget["住房补贴"] = df_budget["住房补贴标准"]
    df_budget["工作补贴"] = df_budget["工作补贴标准"]
    df_budget["交通补贴"] = df_budget["交通补贴标准"]
    df_budget["班中餐"] = df_budget["班中餐标准"]
    df_budget["误餐"] = df_budget["误餐标准"]
    df_budget["中班津贴"] = df_budget["中班津贴标准"]
    df_budget["夜班津贴"] = df_budget["夜班津贴标准"]


    # 发现pandas的一个特点，如果有na值，就不会计算这一行的求和
    df_budget = df_budget.fillna(0)

    df_budget["应发工资"] = (df_budget["预支年薪"] + df_budget["岗位工资"] + df_budget["奖金"] + df_budget["加班工资"] + df_budget[
        "年功保留工资"] \
                         + df_budget["竞业津贴"] + df_budget["技术津贴"] + df_budget["回民补贴"] + df_budget["女工洗理费"] + df_budget[
                             "住房补贴"] \
                         + df_budget["工作补贴"] + df_budget["交通补贴"] + df_budget["班中餐"] + df_budget["误餐"]  + df_budget["中班津贴"] + df_budget["夜班津贴"] \
                         + df_budget["其他补发"] + df_budget["其他补扣"]+ df_budget["取暖费"] ) * df_budget["合计"]+ df_budget["异地津贴"] \
                         + df_budget["异地差旅费"] + df_budget["外租房津贴"]

    df_budget["养老保险单位实缴"] = df_budget["应发工资"] * 0.16
    df_budget["失业保险单位实缴"] = df_budget["应发工资"] * 0.007
    df_budget["工伤保险单位实缴"] = df_budget["应发工资"] * 0.017
    df_budget["医疗保险单位实缴"] = df_budget["应发工资"] * 0.07
    df_budget["生育保险单位实缴"] = df_budget["应发工资"] * 0.008
    df_budget["公积金单位实缴"] = df_budget["应发工资"] * 0.12
    df_budget["职工教育经费"] = df_budget["应发工资"] * 0.015
    df_budget["补充医疗保险"] = 11 * df_budget["合计"]
    df_budget["工资总额（含保险）"] = df_budget["应发工资"] + df_budget["养老保险单位实缴"] + df_budget["失业保险单位实缴"] + df_budget["工伤保险单位实缴"] \
                             + df_budget["医疗保险单位实缴"] + df_budget["生育保险单位实缴"] + df_budget["公积金单位实缴"] + df_budget[
                                 "补充医疗保险"] + df_budget["职工教育经费"]

    float_workincome_summary = df_budget["应发工资"].sum()
    float_workincome_summary2 = df_budget["工资总额（含保险）"].sum()

    print(df_budget.columns)
    print("工资总额（不含保险）为{}".format(float_workincome_summary))
    print("工资总额（含保险）为{}".format(float_workincome_summary2))

    conn = sqlite3.connect("budget")
    df_budget.to_sql("bdget202012", conn, if_exists="replace")
    excelFormat("2020年12月份岗位预算表.xlsx", df_budget, 9, 15, "A:BU")
