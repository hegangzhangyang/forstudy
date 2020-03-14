# -*-coding:utf-8-*-
import xlsxwriter
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from lib.wrapper import printHello
from lib.common import consoleFormat, excelFormat
import pandas as pd


class BounsDetail(object):
    def __init__(self, window, date, conn):
        # 设定控制台输出格式
        consoleFormat()
        self.window = window
        self.date = date
        self.conn = conn


    def start(self):
        try:
            str_sql = "select * from bouns{}sql".format(self.date)
            list_need = ['员工编号', '姓名', '身份证号码', '银行账号', '员工子组名称', '部室', '应发奖金', '其他补发', '其他补扣', '实发奖金', '累计收入额',
                         '累计减除费用', '累计子女教育', '累计赡养老人', '累计继续教育', '累计住房贷款利息', '累计住房租金', '累计专项扣除', '累计已预缴税额',
                         '工资已缴税额', '累计应纳税所得额', '累计应纳税额', '累计减免税额', '累计应扣缴税额', '累计应补(退)税额']
            df_bouns = pd.read_sql_query(str_sql, self.conn)[list_need]
            str_file_path, res = QFileDialog.getSaveFileName(self.window, "保存文件至", r"d:\{}.xlsx"
                                                             .format("河钢乐亭{}奖金明细".format(self.date)),
                                                             "excel 文件(*.xlsx)")
            excelFormat(str_file_path, df_bouns, 9, 20)


        except xlsxwriter.exceptions.FileCreateError:

            QMessageBox.warning(self.window, "河钢乐亭薪酬管理系统", "同名文件正在打开，请关闭后重新生成。", QMessageBox.Yes)

        except Exception as e:

            QMessageBox.warning(self.window, "河钢乐亭薪酬管理系统", "发生错误{}".format(str(e)), QMessageBox.Yes)

        else:

            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "奖金明细生成完毕！", QMessageBox.Ok)
