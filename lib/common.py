# -*-coding:utf-8-*-
import os

import pandas as pd
import datetime
from PyQt5.QtWidgets import QFileDialog


def consoleFormat():
    # 这两段代码设置输出时对齐
    pd.set_option('display.unicode.ambiguous_as_wide', True)
    pd.set_option('display.unicode.east_asian_width', True)
    # 设置最大显示列数为10
    pd.set_option('display.max_columns', 100)
    # 设置控制台现实不允许换行
    pd.set_option('expand_frame_repr', False)


def getCurrentYear():
    """返回字符串格式的当前年"""
    return int(datetime.datetime.now().year)


def getCurrentMonth():
    """返回字符串格式的当前月"""
    return int(datetime.datetime.now().month)


def getCurrentDay():
    """返回字符串格式的当前日"""
    return int(datetime.datetime.now().day)


def floatReturnPercent(year):
    t = 0
    if year < 5:
        t = 0.4
    elif year < 10:
        t = 0.3
    elif year < 15:
        t = 0.2
    elif year < 20:
        t = 0.1
    elif year > 20:
        t = 0
    return t


def max0(num):
    """
    如果数值小鱼0，那么返回0
    """
    if num < 0:
        return 0
    else:
        return num


def tex(income):
    """
    用来计税的函数
    :param income: 计税额
    :return: 返回个税
    """
    if 0 <= income <= 36000:
        temp = income * 0.03
    elif 3600 < income <= 144000:
        temp = income * 0.1 - 2520
    elif 144000 < income <= 300000:
        temp = income * 0.2 - 16920
    elif 300000 < income <= 420000:
        temp = income * 0.25 - 31920
    elif 420000 < income <= 660000:
        temp = income * 0.3 - 52920
    elif 660000 < income <= 960000:
        temp = income * 0.35 - 85920
    elif income > 960000:
        temp = income * 0.45 - 181920
    else:
        temp = 0
    return temp


def strReturnFileName(window, date, str_file_name):
    file_path = QFileDialog.getExistingDirectory(window, "保存至此文件夹",
                                                 r"C:\Users\Administrator\Desktop")
    if file_path != "":
        file_name = os.path.join(file_path, date + "{}.xlsx".format(str_file_name))
        return file_name
    else:
        return False


def excelFormat(str_file_path, df, int_font_size, float_column_width, str_cloumn_area="A:CZ"):
    writer = pd.ExcelWriter(str_file_path, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    excel_format = workbook.add_format({'align': 'center', 'font_size': int_font_size})
    worksheet.set_column(str_cloumn_area, float_column_width, excel_format)
    writer.save()


def listGetExcelFullName(path):
    """
    获取路径，输出路径下所有的excel文件的全路径名
    :param path: 文件夹名称
    :return: 列表，包含文件的全路径名
    """
    file_list = os.listdir(path)
    file_list = list(filter(lambda x: x.endswith(".xlsx") or x.endswith(".xls"), file_list))
    file_list = list(map(lambda x: os.path.join(path, x), file_list))

    return file_list


def dfChangeFormat(filename):
    """
    获取excel，转成pandas的dateframe输出
    :param filename: 文件的全路径名
    :return: dataframe
    """
    book = pd.read_excel(filename, skiprows=5, skipfooter=3, usecols=["员工编号", "应发奖金"],
                         dtype=str)
    book = book.dropna()
    return book


def dfSumExcel(excel_name_list):
    """
    传入含有excel全路径名的列表，自动统计成一个dataframe
    :param excel_name_list: 含有excel全路径名的列表
    :return: datafame
    """
    book = dfChangeFormat(excel_name_list[0])

    for index, item in enumerate(excel_name_list):
        if index > 0:
            res = dfChangeFormat(item)
            book = book.append(res)
            book.reset_index(drop=True, inplace=True)
    return book


def sumExcelFromPathAndToSql(input_path, str_input_to_sql_name, conn):
    """
    集成以上的操作，给定路径，直接汇总并输出到给定路径下生成excel
    :param input_path: 待汇总的文件所在文件夹
    :param output_path: 汇总后Excel输出的路径
    :param out_file_name: 输出的文件名
    :return:
    """
    final = listGetExcelFullName(input_path)
    book = dfSumExcel(final)
    book.to_excel("hahaha.xlsx")
    book.to_sql(str_input_to_sql_name, conn, if_exists="replace")
