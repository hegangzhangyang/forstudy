# -*-coding:utf-8-*-

import pandas as pd

from PyQt5.QtWidgets import QMessageBox


class ExcelToSql(object):
    """
    将excel导入时，进行各类处理
    """

    def __init__(self):
        pass

    @staticmethod
    def importToSql(window, conn, dataframe_name, sql_table_name):
        """
        试图将文件导入数据库
        :param file_name: 需要导入的文件名
        :param sql_table_name: 导入数据库后的表名
        :return:无
        """

        try:
            dataframe_name.to_sql(sql_table_name, conn,  if_exists="replace",index=False)
        except Exception as e:
            QMessageBox.information(window, "河钢乐亭薪酬管理系统", "发生错误，内容如下：{}".format(str(e)), QMessageBox.Ok,
                                    QMessageBox.Ok)
            print(e)
        else:

            QMessageBox.information(window, "河钢乐亭薪酬管理系统", "成功将  文件导入数据库",
                                    QMessageBox.Ok, QMessageBox.Ok)

    @staticmethod
    def data_clean(window, file_name):
        """
        对excel文件内容进行清洗，去字符串内空格，逗号，空值填充为0
        :param file_name: 需要导入的文件名
        :return: 清洗过后的dataframe，全部为字符串格式
        """
        try:
            # 尝试读取excel表到内存中
            dataframe_name = pd.read_excel(file_name, dtype=str)

        except FileNotFoundError:
            QMessageBox.information(window, "河钢乐亭薪酬管理系统", "文件没有找到，请重试。", QMessageBox.Ok, QMessageBox.Ok)

        except Exception as e:
            QMessageBox.information(window, "河钢乐亭薪酬管理系统", "发生未知错误，内容如下：{}".format(str(e)), QMessageBox.Ok,
                                    QMessageBox.Ok)
        else:

            # 将excel中的空值、na项全部填充为0
            dataframe_name = dataframe_name.fillna("0")

            for col in dataframe_name.columns:
                # 尝试将excel中的空白字符替换掉
                dataframe_name[col] = dataframe_name[col].apply(lambda s: "".join(s.split()))
                print("{}列的空格已经去除".format(col))

                # 如果内容里面有逗号，就替换掉
                dataframe_name[col] = dataframe_name[col].apply(lambda s: s.replace(",", ""))
                print("{}列的逗号已经去除".format(col))

            return dataframe_name

    @classmethod
    def excel_to_sql(cls, window, conn, file_name, sql_table_name):
        """把其他方法综合起来，这样就可以利用这一个方法，简化操作"""
        dataframe = cls.data_clean(window, file_name)
        cls.importToSql(window, conn, dataframe, sql_table_name)
