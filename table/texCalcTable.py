# -*-coding:utf-8-*-
import os

import pandas as pd
import xlsxwriter
from PyQt5.QtWidgets import QMessageBox

from lib.common import strReturnFileName
from lib.common import consoleFormat, excelFormat
from lib.erorClass import DataFormatError


class TexCalc(object):

    def __init__(self, window, conn, date):
        consoleFormat()
        self.window = window
        self.conn = conn
        self.date = date

    def dfTexCalc(self):
        list_need_calc = ["员工编号", "姓名", "身份证号码", "系统工资税基", "养老保险员工实缴", "失业保险员工实缴", "医疗保险员工实缴", "住房公积金员工实缴",
                          "养老保险-个人补缴", "失业保险-个人补缴", "医疗保险-个人补缴", "住房公积金-个人补缴"]
        df_wage_detail = pd.read_sql_query("select  * from wage{}sql".format(self.date), self.conn)[list_need_calc]

        df_wage_detail["基本养老保险费"] = df_wage_detail["养老保险员工实缴"] + df_wage_detail["养老保险-个人补缴"]
        df_wage_detail["基本医疗保险费"] = df_wage_detail["医疗保险员工实缴"] + df_wage_detail["医疗保险-个人补缴"]
        df_wage_detail["失业保险费"] = df_wage_detail["失业保险员工实缴"] + df_wage_detail["失业保险-个人补缴"]
        df_wage_detail["住房公积金"] = df_wage_detail["住房公积金员工实缴"] + df_wage_detail["住房公积金-个人补缴"]

        df_wage_detail.rename(columns={"员工编号": "工号", "身份证号码": "*证照号码", "系统工资税基": "*本期收入"}, inplace=True)
        list_need_drop = ["养老保险员工实缴", "失业保险员工实缴", "医疗保险员工实缴", "住房公积金员工实缴", "养老保险-个人补缴", "失业保险-个人补缴", "医疗保险-个人补缴",
                          "住房公积金-个人补缴"]
        df_wage_detail.drop(columns=list_need_drop, inplace=True)
        df_wage_detail.insert(2, "*证照类型", "居民身份证")
        df_wage_detail.insert(5, "本期免税收入", "")
        df_wage_detail.insert(10, "累计子女教育支出", "")
        df_wage_detail.insert(11, "累计继续教育支出", "")
        df_wage_detail.insert(12, "累计住房贷款利息支出", "")
        df_wage_detail.insert(13, "累计住房租金支出", "")
        df_wage_detail.insert(14, "累计赡养老人支出", "")
        df_wage_detail.insert(15, "企业(职业)年金", "")
        df_wage_detail.insert(16, "商业健康保险", "")
        df_wage_detail.insert(17, "税延养老保险", "")
        df_wage_detail.insert(18, "其他", "")
        df_wage_detail.insert(19, "准予扣除的捐赠额", "")
        df_wage_detail.insert(20, "减免税额", "")
        df_wage_detail.insert(21, "备注", "")

        return df_wage_detail

    def start(self):
        str_file_name = strReturnFileName(self.window, self.date, "报税文件")
        print(str_file_name)
        df_wage_detail = self.dfTexCalc()
        try:
            if str_file_name:
                df_wage_detail.to_excel(str_file_name, index=False)
                excelFormat(str_file_name, df_wage_detail, 9, 12, "A:V")

        except xlsxwriter.exceptions.FileCreateError:
            QMessageBox.warning(self.window, "河钢乐亭薪酬管理系统", "同名文件正在打开，请关闭后重新生成。", QMessageBox.Yes)

        except DataFormatError as e:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", str(e), QMessageBox.Ok)

        except Exception as e:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "发生未知错误，内容如下{}".format(str(e)), QMessageBox.Ok)
        else:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "报税文件生成完毕！", QMessageBox.Ok)
