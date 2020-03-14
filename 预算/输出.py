# -*-coding:utf-8-*-
import sqlite3
import pandas as pd

from lib.common import excelFormat

if __name__ == '__main__':

    conn = sqlite3.connect("budget")
    list_need_out = ['部室', '作业区', "职务",
                     '合计', '奖金系数', '岗位工资系数',
                     '预支年薪', '岗位工资', '奖金', '加班工资', '年功保留工资', '竞业津贴',
                     '技术津贴', '回民补贴', '女工洗理费', '住房补贴', '工作补贴',
                     '交通补贴', '班中餐', '误餐', '通讯补贴', '异地津贴', '异地差旅费',
                     '外租房津贴', '中班津贴', '夜班津贴', '其他补发', '其他补扣', '取暖费',
                     '应发工资', '养老保险单位实缴', '失业保险单位实缴', '工伤保险单位实缴',
                     '医疗保险单位实缴', '生育保险单位实缴', '公积金单位实缴',
                     '职工教育经费', '补充医疗保险', '工资总额（含保险）']
    list_extra = ["残疾人保障金", "挖潜增效奖"]

    book0 = pd.read_sql_query("select * from bdget20201", conn)[list_need_out]
    # excelFormat("./发冯文樵/2020年1月工资总额预算.xlsx", book0, 9, 15, "A:BU")

    for i in range(2, 13):
        book = pd.read_sql_query("select * from bdget2020{}".format(i), conn)[list_need_out]
        # excelFormat("./发冯文樵/2020年{}月工资总额预算.xlsx".format(i), book, 9, 15, "A:BU")
        book0 = book0.append(book, ignore_index=True)
    for i in book0.columns:
        if not i in ["部室","作业区","职务","合计","异地津贴","异地差旅费","外租房津贴"]:
            book0["{}总额".format(i)] = book0[i] * book0["合计"]
    # book0.to_excel("123.xlsx")


    list_need_out.extend(list_extra)
    book5 = pd.read_sql_query("select * from bdget20205", conn)[list_need_out]
    # excelFormat("./发冯文樵/2020年5月工资总额预算.xlsx", book5, 9, 15, "A:BU")



    print("工资总额为{}".format(book0["应发工资"].sum()+4500000+3200000-7000000))
    print("平均职工人数为{}".format(book0["合计"].sum() / 12))
    print("人均工资为{}".format(book0["应发工资"].sum() / (book0["合计"].sum() / 12)))
