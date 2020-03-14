# -*-coding:utf-8-*-
import sqlite3

import pandas as pd

book = pd.read_excel("123.xlsx")

# 冷轧工资总额
# 冷轧岗位工资
float_lengzha_pos = 1771228

# 冷轧奖金
float_lengzha_JiangJin = 2513412

# 冷轧挖潜增效奖
float_lengzha_WaQaian=643360

# 岗位工资
float_pos = (book["岗位工资总额"] + book["预支年薪总额"]).sum()+float_lengzha_pos
# 技术津贴
float_tehc = book["技术津贴总额"].sum()
# 加班工资
float_extra_work = book["加班工资总额"].sum()
# 津补贴
float_JinBuTie = (book["年功保留工资总额"] + book["竞业津贴总额"] + book["回民补贴总额"] + book["女工洗理费总额"] + book["住房补贴总额"] \
                  + book["夜班津贴总额"] + book["中班津贴总额"]).sum()

# 绩效奖金（普通职工）
float_common_bouns = (book.loc[~book["职务"].isin(["部长", "副部长", "部长助理"]), "奖金总额"]).sum()+float_lengzha_JiangJin

# 厂部级绩效薪
float_CHangBuJi_NianXin = (book.loc[book["职务"].isin(["部长", "副部长", "部长助理"]), "奖金总额"]).sum()

# 高层
float_gaoceng = 4500000

# 挖潜增效奖
conn = sqlite3.connect("budget")
book5 = pd.read_sql_query("select * from bdget20205", conn)
float_WaQaian = (book5["挖潜增效奖"] * book5["合计"]).sum()+float_lengzha_WaQaian

# 福利性津补贴
float_FuLi = (book["工作补贴总额"] + book["交通补贴总额"] + book["班中餐总额"] + book["误餐总额"] + book["异地津贴"] + book["异地差旅费"] + book[
    "外租房津贴"]).sum()
# 专家年薪结算
float_ZHuanJia = 3200000

# 取暖费
float_QuNuanFei = book["取暖费总额"].sum()

# 工资总额
float_all = float_pos + float_tehc + float_extra_work + float_JinBuTie + float_common_bouns + float_WaQaian \
            + float_CHangBuJi_NianXin + float_FuLi + float_gaoceng + float_ZHuanJia

# 平均职工人数

int_people_num = book["合计"].sum() / 12

print("岗位工资总额为{}".format(float_pos))
print("技术津贴总额为{}".format(float_tehc))
print("加班工资总额为{}".format(float_extra_work))
print("岗位津补贴总额为{}".format(float_JinBuTie))
print("奖金总额为{}".format(float_common_bouns+float_WaQaian+float_CHangBuJi_NianXin+float_ZHuanJia+float_gaoceng))
print("经济责任制奖金总额为{}".format(float_common_bouns))
print("挖潜增效奖为{}".format(float_WaQaian))
print("年薪结算总额为{}".format(float_CHangBuJi_NianXin+float_ZHuanJia))
print("高层工资总额为{}".format(float_gaoceng))
print("福利性津补贴总额为{}".format(float_FuLi))
print("取暖费总额为{}".format(float_QuNuanFei))
print("工资总额为{}".format(float_all))
print("平均工资为{}".format(float_all / int_people_num))
