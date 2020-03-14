# -*-coding:utf-8-*-
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from lib.common import consoleFormat
from lib.erorClass import DataFormatError


class ExcelSplit(object):
    def __init__(self, window, conn):
        consoleFormat()
        self.window = window
        self.conn = conn

    def boolCheckTitle(self, df):
        "对列表中的每一项进行检查，如果不在excel中，那么意味着无法拆分"
        list_need_check = ["部室"]
        for item in list_need_check:
            if not item in df.columns:
                return False
            else:
                return True

    def excelSplit(self):
        str_file_name = QFileDialog.getOpenFileName(self.window, "请选择拆分文件", './')[0]
        str_file_path = QFileDialog.getExistingDirectory(self.window, "请选择拆分到那个文件夹", './')
        print(str_file_path)
        if str_file_name != "" and str_file_path != "":
            df_all_detail = pd.read_excel(str_file_name)
            res = self.boolCheckTitle(df_all_detail)
            if res:
                list_dept_neme = df_all_detail["部室"].unique()
                for item in list_dept_neme:
                    df_dept = df_all_detail[df_all_detail["部室"] == item]
                    df_dept.to_excel(str_file_path + "/" + item.replace("\n", "") + ".xlsx")

        else:
            raise DataFormatError("拆分文件和拆分后文件夹必须指定，否则无法拆分！")

    def start(self):
        try:
            self.excelSplit()
        except DataFormatError as e:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", str(e), QMessageBox.Ok)

        except Exception as e:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "发生未知错误，内容如下{}".format(str(e)), QMessageBox.Ok)

        else:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "文件拆分完成，请查看！", QMessageBox.Ok)
