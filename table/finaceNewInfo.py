import pandas as pd
import xlsxwriter
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from pandas import to_datetime
import arrow
from lib.common import consoleFormat


class NewPeopleInfo(object):

    def __init__(self, window, conn, year, month):
        consoleFormat()
        self.window = window
        self.year = year
        self.month = month
        self.conn = conn
    def strForamtDate(self, str_date):
        str_date1 = arrow.get(str_date)
        str_date2 = str_date1.format("YYYY-MM-01")
        if not (str_date1.year == self.year and str_date1.month == self.month):
            str_date2 = ""
        return str_date2

    def excelCreatTable(self):
        df_wage_detail = pd.read_sql_query(
            "select 员工编号,姓名,身份证号码 from wage{}sql".format(str(self.year) + str(self.month)), self.conn)
        df_other_info = pd.read_sql_query(
            "select 员工编号,联系方式,合同起始时间 from t{}PeopleInfo".format(str(self.year) + str(self.month)), self.conn)

        df_all_info = pd.merge(df_wage_detail, df_other_info, on="员工编号", how="left")
        df_all_info = df_all_info.dropna()
        # 把单元格的格式更改为时间格式，以便按照年月日取出
        # df_all_info["合同起始时间"] = to_datetime(df_all_info["合同起始时间"], format=("%Y-%m-%d"))
        # 只选出年月日都符合条件的人，即参加工作时间等同于计算工资年和月的人
        df_all_info["合同起始时间"] = df_all_info["合同起始时间"].apply(self.strForamtDate)

        df_all_info = df_all_info.loc[df_all_info["合同起始时间"] != ""]

        df_all_info.rename(columns={"员工编号": "工号", "姓名": "*姓名", "身份证号码": "*证照号码", "联系方式": "手机号码", "合同起始时间": "任职受雇日期"},
                           inplace=True)

        df_all_info.insert(2, column="*证照类型", value="居民身份证")
        df_all_info.insert(4, column="*国籍(地区)", value="")
        df_all_info.insert(5, column="*性别", value="")
        df_all_info.insert(6, column="*出生日期", value="")
        df_all_info.insert(7, column="*人员状态", value="正常")
        df_all_info.insert(8, column="*是否雇员", value="是")
        df_all_info.insert(11, column="离职日期", value="")
        df_all_info.insert(12, column="是否残疾", value="")
        df_all_info.insert(13, column="是否烈属", value="")
        df_all_info.insert(14, column="是否孤老", value="")
        df_all_info.insert(15, column="残疾证号", value="")
        df_all_info.insert(16, column="烈属证号", value="")
        df_all_info.insert(17, column="个人投资额", value="")
        df_all_info.insert(18, column="个人投资比例(%)", value="")
        df_all_info.insert(19, column="备注", value="")
        df_all_info.insert(20, column="是否境外人员", value="")
        df_all_info.insert(21, column="中文名", value="")
        df_all_info.insert(22, column="出生国家(地区)", value="")
        df_all_info.insert(23, column="首次入境时间", value="")
        df_all_info.insert(24, column="预计离境时间", value="")
        df_all_info.insert(25, column="其他证照类型", value="")
        df_all_info.insert(26, column="其他证照号码", value="")
        df_all_info.insert(27, column="户籍所在地（省）", value="")
        df_all_info.insert(28, column="户籍所在地（市）", value="")
        df_all_info.insert(29, column="户籍所在地（区县）", value="")
        df_all_info.insert(30, column="户籍所在地（详细地址", value="")
        df_all_info.insert(31, column="居住地址（省）", value="")
        df_all_info.insert(32, column="居住地址（市）", value="")
        df_all_info.insert(33, column="居住地址（区县）", value="")
        df_all_info.insert(34, column="居住地址（详细地址）", value="")
        df_all_info.insert(35, column="联系地址（省））", value="")
        df_all_info.insert(36, column="联系地址（市）", value="")
        df_all_info.insert(37, column="联系地址（区县）", value="")
        df_all_info.insert(38, column="联系地址（详细地址）", value="")
        df_all_info.insert(39, column="电子邮箱", value="")
        df_all_info.insert(40, column="学历", value="")
        df_all_info.insert(41, column="开户银行", value="")
        df_all_info.insert(42, column="银行账号", value="")
        df_all_info.insert(43, column="职务", value="")

        str_file_path, res = QFileDialog.getSaveFileName(self.window, "保存文件至", r"d:\{}.xlsx"
                                                         .format("{}{}新增人员信息".format(self.year, self.month)),
                                                         "excel 文件(*.xlsx)")
        if res:
            writer = pd.ExcelWriter(str_file_path, engine='xlsxwriter')
            df_all_info.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            excel_format = workbook.add_format({'align': 'center', 'font_size': 9})
            worksheet.set_column("A:BI", 15.5, excel_format)
            writer.save()
        else:
            return

    def start(self):
        try:
            self.excelCreatTable()

        except xlsxwriter.exceptions.FileCreateError:

            QMessageBox.warning(self.window, "河钢乐亭薪酬管理系统", "同名文件正在打开，请关闭后重新生成。", QMessageBox.Yes)

        except Exception as e:
            QMessageBox.warning(self.window, "河钢乐亭薪酬管理系统", "发生错误{}".format(str(e)), QMessageBox.Yes)

        else:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "新增人员信息表生成完毕！", QMessageBox.Ok)
