# -*-coding:utf-8-*-
import pandas as pd
from lib.erorClass import DataFormatError
from PyQt5.QtWidgets import QMessageBox


class DataCheck(object):
    """
    用来检查文件是否符合规范
    """

    def __init__(self):
        pass

    # 文件格式检查
    @staticmethod
    def fileNameCheck(file_name):
        file_name = ".".join([s.lower() for s in file_name.split(".")])
        if not (file_name.endswith("xlsx") or file_name.endswith("xls")):
            raise DataFormatError("文件扩展名不正确，应为'xlsx'或者'xls'!")

    @staticmethod
    def fileTitleCheck(file_name, standrad_list):
        book = pd.read_excel(file_name)
        need_check_list = list(book.columns)

        # 判断标准列表中的元素是否是导入文件第一列的子集，如果是，说明含有必须项
        res = set(standrad_list).issubset(need_check_list)
        if not res:
            raise DataFormatError("文件标题列含有非法值或者你选择导入的文件不正确，请检查。")

    @staticmethod
    def fileCloumnDuplicateCheck(file_name):

        book = pd.read_excel(file_name)
        columns_list = list(book.columns)
        for name in columns_list:
            if ".1" in name or ".2" in name or ".3" in name:
                raise DataFormatError("文件的标题列有重复项，请检查！")

    # 导入文件内容检查
    @staticmethod
    def fileCLoumnsCheck(file_name):
        book = pd.read_excel(file_name)
        list_need_check_na = ["员工编号", "人员编号", "身份证号码", "银行账号", "工资项", "变更后岗位系数", "证件号码"
            , "个人月缴额", "身份证号", "个人编号", "缴费基数", "证照号码", "工号", "身份信息", "部室", "民族", "员工子组名称","SAP部门代码"]

        list_need_check_dumplicate = ["员工编号", "人员编号", "身份证号码", "银行账号", "证件号码", "身份证号", "个人编号"
            , "证照号码", "工号"]

        list_need_check_dumplicate2 = ["人员编号", "银行账号", "身份证号码","证件号码", "身份证号", "个人编号" , "证照号码", "工号"]
        # 如果表中含有工资项这一列，那么允许员工编号重复，因为一个员工可能会导入多个工资项
        if "工资项" in book.columns:
            for str_col_name in list_need_check_dumplicate2:
                if str_col_name in book.columns:
                    # 检查是否有重复
                    if book[str_col_name].duplicated().any():
                        raise DataFormatError("文件内{}这一列含有重复项，请检查。".format(str_col_name))
        else:
            for str_col_name in list_need_check_dumplicate:
                if str_col_name in book.columns:
                    # 检查是否有重复
                    if book[str_col_name].duplicated().any():
                        raise DataFormatError("文件内{}这一列含有重复项，请检查。".format(str_col_name))

        for str_col_name in list_need_check_na:
            if str_col_name in book.columns:
                # 检查是否有空值
                if book[str_col_name].isna().any():
                    raise DataFormatError("文件内{}这一列含有空白值，请检查。".format(str_col_name))

    @classmethod
    def fileCheckWhole(cls, window, file_name, standrad_list):
        """
        整合检查文件名，文件第一列和员工编号列
        :param file_name: 文件名
        :param standrad_list: 需要检查的第一列名字
        :return: 无，如果检查出错会抛出异常
        """
        try:
            cls.fileNameCheck(file_name)
            print("文件名检查通过")
            cls.fileTitleCheck(file_name, standrad_list)
            print("文件标题列合法性检查通过")
            cls.fileCLoumnsCheck(file_name)
            print("文件各列内容检查通过")
            cls.fileCloumnDuplicateCheck(file_name)
            print("标题列重复性检查通过")
        except Exception as e:
            QMessageBox.information(window, "河钢乐亭薪酬管理系统", "发生错误，内容如下：{}".format(str(e)), QMessageBox.Ok,
                                    QMessageBox.Ok)
            return False
        else:
            return True
