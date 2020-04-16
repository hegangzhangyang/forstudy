# -*-coding:utf-8-*-
import sqlite3

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QInputDialog, QFileDialog, QMessageBox
from PyQt5.QtWidgets import QMainWindow

from UI.UI_mianWindow import Ui_mainWindow
from core.bouns import Bouns
from core.wage import Wage
from lib.common import getCurrentMonth, getCurrentYear, sumExcelFromPathAndToSql, strReturnFileName
from lib.dataCheck import DataCheck
from lib.exceImport import ExcelToSql
from table.bounsDetail import BounsDetail
from table.finaceNewInfo import NewPeopleInfo
from table.texCalcTable import TexCalc
from table.wageDetail import WageDetail
from table.wageSplit import ExcelSplit
from table.wageSummary import WageSummary


class MainWindow(QMainWindow, Ui_mainWindow):
    """
    Class documentation goes here.
    """

    def __init__(self, parent=None):
        """
        构造函数，进行必要初始化。如时间，数据库连接
        """
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.conn = sqlite3.connect(r"..\db\workincome1.db")
        # 自动设定为当前年月，方便修改
        self.spinBox_year.setValue(getCurrentYear())
        self.spinBox_month.setValue(getCurrentMonth())
        self.spinBox_bouns_year.setValue(getCurrentYear())
        self.spinBox_bouns_month.setValue(getCurrentMonth())

    @pyqtSlot()
    def on_btn_table_standard_clicked(self):
        """工资标准表"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["员工编号", "姓名", "身份证号码"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}StandradTable".format(current_time))

    @pyqtSlot()
    def on_btn_table_position_change_clicked(self):
        """岗位变更表"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["员工编号", "姓名", "变更后岗位系数"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}PositionChange".format(current_time))

    @pyqtSlot()
    def on_btn_table_sap_workday_clicked(self):
        """sap考勤表"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["人员编号", "姓名", "应出勤天数", "年假天数", "事假天数", "病假天数", "工伤假天数", "探亲假天数", "婚假天数", "丧假天数", "产假天数"
            , "育儿假天数", "节育假天数", "旷工天数", "高温出勤", "有毒有害出勤", "工作补贴出勤", "节假日加班", "加班天数", "中班天数", "夜班天数"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}SapWorkday".format(current_time))

    @pyqtSlot()
    def on_btn_table_people_info_clicked(self):
        """人员信息表"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["部室", "姓名", "员工编号", "身份信息", "员工子组名称",  "合同起始时间","身份证号码"]
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}PeopleInfo".format(current_time))

    @pyqtSlot()
    def on_btn_table_endowment_insurance_roster_clicked(self):
        """养老保险缴费花名册"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["个人编号", "姓名", "证件号码", "出生日期", "参加工作日期",
                         "企业养老保险", "失业保险"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}EndowmentInsuranceRoster".format(current_time))

    @pyqtSlot()
    def on_btn_table_endowment_insurance_detail_clicked(self):
        """养老保险缴费明细"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["个人编号", "单位编号", "姓名", "费款所属期", "险种类型", "缴费类型",
                         "缴费标志", "缴费基数", "个人缴"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}EndowmentInsuranceDetail".format(current_time))

    @pyqtSlot()
    def on_btn_table_unemployment_insurance_detail_clicked(self):
        """失业保险缴费明细"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["个人编号", "单位编号", "姓名", "费款所属期", "险种类型", "缴费类型",
                         "缴费标志", "缴费基数", "个人缴"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}UnemployementInsuranceDetail".format(current_time))

    @pyqtSlot()
    def on_btn_table_house_insurance_clicked(self):
        """公积金缴费明细"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["个人账号", "个人账户状态", "开户日期",  "个人缴存基数", "单位月缴额",
                         "个人月缴额", "月缴额", "缴至日期"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}HouseInsuranceDetail".format(current_time))

    @pyqtSlot()
    def on_btn_table_medical_insurance_roster_clicked(self):
        """医疗保险缴费花名册"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["证件号码", "个人编号", "姓名", "实际缴费月数", "计算后视同缴费",
                         "险种类型", "医疗离退休状态", "参保状态"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}MedicalInsuranceRoster".format(current_time))

    @pyqtSlot()
    def on_btn_table_medical_insurance_detail_clicked(self):
        """医疗保险缴费明细"""

        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["个人编号", "单位编号", "姓名", "费款所属期", "险种类型", "缴费类型",
                         "缴费标志", "缴费基数", "个人缴"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}MedicalInsuranceDetail".format(current_time))

    @pyqtSlot()
    def on_btn_table_civil_servants_insurance_detail_clicked(self):
        """公务员缴费明细"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["个人编号", "单位编号", "姓名", "费款所属期", "险种类型", "缴费类型",
                         "缴费标志", "缴费基数", "个人缴"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}CivilServantsInsuranceDetail".format(current_time))

    @pyqtSlot()
    def on_btn_table_insurance_extra_payment_clicked(self):
        """保险补缴信息"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["员工编号", "工资项", "金额"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}InsuranceExtraPayment".format(current_time))

    @pyqtSlot()
    def on_btn_table_tax_extra_clicked(self):
        """财务额外计税项"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["员工编号", "工资项","金额"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}TaxExtra".format(current_time))

    @pyqtSlot()
    def on_btn_table_tax_remove_clicked(self):
        """财务附加扣除"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["证照类型", "姓名", "本期收入", "累计子女教育", "累计住房贷款利息",
                         "累计住房租金", "累计赡养老人"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}TaxRemove".format(current_time))

    @pyqtSlot()
    def on_btn_table_last_month_tax_detail_clicked(self):
        """上月缴税明细"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["工号", "累计收入额", "累计减除费用", "累计子女教育支出扣除", "累计赡养老人支出扣除",
                         "累计继续教育支出扣除", "累计住房贷款利息支出扣除"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}LastMonthTaxDetail".format(current_time))

    @pyqtSlot()
    def on_btn_table_extra_workincome_clicked(self):
        """工资补发补扣"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["员工编号", "姓名", "其他补扣", "其他补发"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}ExtraWorkincome".format(current_time))

    @pyqtSlot()
    def on_btn_last_month_wage_detail_clicked(self):
        """上月工资明细"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["员工编号", "姓名", "岗位工资系数"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}LastMonthWageDetail".format(current_time))

    @pyqtSlot()
    def on_btn_workincome_calc_clicked(self):
        """工资生成"""
        workday, ok = QInputDialog.getInt(self, '河钢乐亭薪酬管理系统', '请输入当前月份应出勤天数：', min=18)
        if ok:
            current_time = self.spinBox_year.text() + self.spinBox_month.text()
            wage = Wage(self, current_time, int(workday), self.conn)
            print(type(workday))
            wage.start()

    @pyqtSlot()
    def on_btn_bouns_extra_clicked(self):
        """奖金补发补扣"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        file_name = QFileDialog.getOpenFileName(self, '打开文件', './')[0]
        standrad_list = ["员工编号", "姓名", "其他补扣", "其他补发"]
        # 对文件进行内容格式进行检查
        if DataCheck.fileCheckWhole(self, file_name, standrad_list):
            ExcelToSql.excel_to_sql(self, self.conn, file_name, "t{}BounsExtra".format(current_time))

    @pyqtSlot()
    def on_btn_workincome_detail_clicked(self):
        """工资明细表生成"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        xlsx_wage_detail = WageDetail(self, current_time, self.conn)
        xlsx_wage_detail.start()

    @pyqtSlot()
    def on_btn_workincome_summary_clicked(self):
        """工资汇总表生成"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        excel_wage_summary = WageSummary(self, self.conn, current_time)
        excel_wage_summary.start()

    @pyqtSlot()
    def on_btn_finace_new_info_clicked(self):
        """新增人员信息"""
        year = self.spinBox_year.value()
        month = self.spinBox_month.value()

        excel_new_people_info = NewPeopleInfo(self, self.conn, year, month)
        excel_new_people_info.start()

    @pyqtSlot()
    def on_btn_finace_tax_calc_template_clicked(self):
        """正常薪金所得（算税模板）"""
        current_time = self.spinBox_year.text() + self.spinBox_month.text()
        print(current_time)
        excel_tex_calc = TexCalc(self, self.conn, current_time)
        excel_tex_calc.start()

    @pyqtSlot()
    def on_btn_wage_split_clicked(self):
        """工资拆分表"""
        wage_split = ExcelSplit(self, self.conn)
        wage_split.start()

    @pyqtSlot()
    def on_btn_sum_excel_to_sql_clicked(self):
        """各部门绩效汇总"""
        current_time = self.spinBox_bouns_year.text() + self.spinBox_bouns_month.text()
        str_file_path = QFileDialog.getExistingDirectory(self, "请指定各部门奖金明细表所在文件夹", "./")

        try:
            if str_file_path != "":
                sumExcelFromPathAndToSql(str_file_path, "t{}SummaryBouns".format(current_time), self.conn)
        except Exception as e:
            QMessageBox.information(self, "河钢乐亭薪酬管理系统", "绩效导入时发生错误，请检查各部门奖金明细格式是否正确,详细内容为{}".format(str(e)),
                                    QMessageBox.Ok)
        else:
            QMessageBox.information(self, "河钢乐亭薪酬管理系统", "各部门明细导入完成，请继续进行下一步操作", QMessageBox.Ok)

    @pyqtSlot()
    def on_btn_bouns_calc_clicked(self):
        """奖金计算表"""
        current_time = self.spinBox_bouns_year.text() + self.spinBox_bouns_month.text()
        bouns = Bouns(self, current_time, self.conn)
        bouns.start()

    @pyqtSlot()
    def on_btn_bouns_detail_clicked(self):
        """奖金明细表"""
        current_time = self.spinBox_bouns_year.text() + self.spinBox_bouns_month.text()
        bouns_detail = BounsDetail(self, current_time, self.conn)
        bouns_detail.start()
