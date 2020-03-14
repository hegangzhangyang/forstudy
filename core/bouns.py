# -*-coding:utf-8-*-
import pandas as pd
from PyQt5.QtWidgets import QMessageBox

from lib.common import consoleFormat, max0, tex
from lib.erorClass import DataFormatError


class Bouns(object):

    def __init__(self, window, date, conn):
        # 设定控制台输出格式
        consoleFormat()
        self.window = window
        self.date = date
        self.conn = conn

    def dfCreatBasicTable(self):
        """
        根据上月的工资创建基础表
        :return: 包含姓名、员工编号、身份证和银行账号的dataframe
        """
        # basic_book = pd.read_sql_query("select * from wage{}sql".format(int(self.date) - 1), self.conn)
        basic_book = pd.read_sql_query("select * from main.wage201912sql", self.conn)
        basic_book = basic_book[["员工编号", "姓名", "身份证号码", "银行账号", "员工子组名称"]]

        return basic_book

    def dfGetInsurance(self):

        if not self.date is None:
            insurance_book = pd.read_sql_query(("select * from wage{}sql".format(self.date)), self.conn)
            insurance_book = insurance_book[
                ["员工编号", "部室", "累计收入额", "累计减除费用", "累计子女教育", "累计赡养老人", "累计继续教育",
                 "累计住房贷款利息", "累计住房租金", "累计专项扣除", "累计已预缴税额", "累计应补(退)税额"]]

            insurance_book.rename(columns={"累计应补(退)税额": "工资已缴税额"}, inplace=True)

            return insurance_book

    def dfGetBounsSummaryAndExtra(self):
        """把奖金汇总表和补发补扣表到进来"""

        df_bouns_summary = pd.read_sql_query("select * from t{}SummaryBouns".format(self.date), self.conn)[
            ["员工编号", "应发奖金"]]
        df_bouns_summary.to_excel("10.xlsx")
        df_bouns_extra = pd.read_sql_query("select * from t201911BounsExtra".format(self.date), self.conn)[
            ["员工编号", "其他补发", "其他补扣"]]
        df_return = pd.merge(df_bouns_summary, df_bouns_extra, on="员工编号", how="left")

        return df_return

    def dfGetDataFromSql(self):
        """把各种表整合起来，填充"""
        df_bouns_baisc = self.dfCreatBasicTable()
        df_bouns_insurance = self.dfGetInsurance()

        df_bouns_summary = self.dfGetBounsSummaryAndExtra()
        df_return = pd.merge(df_bouns_baisc, df_bouns_summary, on="员工编号", how="left")
        df_return = pd.merge(df_return, df_bouns_insurance, on="员工编号", how="left")

        df_return.fillna("0", inplace=True)

        return df_return

    def dfChangeDataFormat(self, df_bouns):

        list_need_change_format = ['应发奖金', '其他补发', '其他补扣', '累计收入额',
                                   '累计减除费用', '累计子女教育', '累计赡养老人', '累计继续教育',
                                   '累计住房贷款利息', '累计住房租金', '累计专项扣除', '累计已预缴税额', '工资已缴税额']

        try:
            for str_col_name in list_need_change_format:
                if str_col_name in df_bouns.columns:
                    df_bouns[str_col_name] = df_bouns[str_col_name].astype(float)
        except ValueError as e:
            raise DataFormatError(r"你所导入的文件当中，应该为全部为数值的某一列含有字母、字符等无法进行计算的内容，请核实。具体错误信息如下：{}"
                                  .format(str(e)))
        else:
            return df_bouns

    def dfBounsDataCalc(self, df_bouns):
        # todo

        # 对补发补扣处理
        df_bouns["应发奖金"] = df_bouns["应发奖金"] + df_bouns["其他补发"] - df_bouns["其他补扣"]

        # 下面开始算税
        # 1.累计收入额
        df_bouns["累计收入额"] = df_bouns["累计收入额"] + df_bouns["应发奖金"]

        # 累计应纳税所得额
        df_bouns["累计应纳税所得额"] = df_bouns["累计收入额"] - df_bouns["累计减除费用"] - df_bouns["累计专项扣除"] - df_bouns["累计子女教育"] - \
                               df_bouns[
                                   "累计赡养老人"] - df_bouns["累计继续教育"] - df_bouns["累计住房贷款利息"] - df_bouns["累计住房租金"]

        df_bouns["累计应纳税所得额"] = df_bouns["累计应纳税所得额"].apply(max0)

        # 累计应纳税额
        df_bouns["累计应纳税额"] = df_bouns["累计应纳税所得额"].apply(tex)

        # 累计减免税额，这三个残疾人减免80%的税
        df_bouns["累计减免税额"] = 0.0
        for i in df_bouns["员工编号"].index:
            if df_bouns["员工编号"].at[i] in ["9000102", "9000161", "9000367"]:
                df_bouns["累计减免税额"].at[i] = (df_bouns["累计应纳税额"].at[i] * 0.8).round(2)

        # 累计应扣缴税额
        df_bouns["累计应扣缴税额"] = df_bouns["累计应纳税额"] - df_bouns["累计减免税额"]

        df_bouns["累计应补(退)税额"] = (df_bouns["累计应扣缴税额"] - df_bouns["累计已预缴税额"] - df_bouns["工资已缴税额"]).apply(max0)
        df_bouns["实发奖金"] = (df_bouns["应发奖金"] - df_bouns["累计应补(退)税额"]).apply(max0)
        df_bouns["实发奖金"]=df_bouns["实发奖金"].round(2)
        # print(df_bouns.head(20))
        return df_bouns

    def bounsDataCheck(self, df_bouns):
        pass

    def bouns(self):
        """ 奖金计算主函数 """
        df_bouns = self.dfGetDataFromSql()

        # 第二步 将所需要的数据全部取出,将数据格式统一调整至float64
        df_bouns = self.dfChangeDataFormat(df_bouns)

        # 第三步 将数据进行加工
        df_bouns = self.dfBounsDataCalc(df_bouns)

        # 第四部 将生成的表格输入到数据库
        df_bouns.to_sql("bouns{}sql".format(self.date), self.conn, if_exists="replace", index=False)

        # 第五步，数据检查
        self.bounsDataCheck(df_bouns)

    def start(self):
        try:
            self.bouns()
        except pd.io.sql.DatabaseError as e:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "奖金计算所需的表格不存在与数据库中，奖金计算无法继续进行，详细内容如下{}".format(str(e)),
                                    QMessageBox.Ok)
        except Exception as e:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "发生错误，内容如下：{}".format(str(e)), QMessageBox.Ok,
                                    QMessageBox.Ok)
        else:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "奖金数据已经导入数据库，请进行下一步操作！", QMessageBox.Ok)

        # self.bouns()
