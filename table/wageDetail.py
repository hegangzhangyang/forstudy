# -*-coding:utf-8-*-

import pandas as pd
import xlsxwriter
from PyQt5.QtWidgets import QMessageBox, QFileDialog


class WageDetail(object):
    def __init__(self, window, date, conn):
        self.window = window
        self.date = date
        self.conn = conn

    def wageDetail(self):
        df_wage = pd.read_sql_query(r"select * from wage{}sql".format(self.date), self.conn)

        # list_wage_detail = ["员工编号", "姓名", "身份证号码", "银行账号", "部室", "员工子组名称", '岗序', '薪等', '岗位工资系数', '预支年薪',
        #                     '岗位工资', "加班工资", '技能保留工资', '年功保留工资', '竞业津贴', '技术津贴',
        #                     '回民补贴', '女工洗理费', '住房补贴', '工作补贴', '交通补贴',
        #                     '班中餐补贴', '误餐补贴', '通讯补贴', '异地津贴', '异地差旅费',
        #                     '外租房津贴', "中班津贴", "夜班津贴", '其他补发', '其他补扣', "应发工资",
        #                     '养老保险员工实缴', '失业保险员工实缴', '医疗保险员工实缴', '住房公积金员工实缴',
        #                     '养老保险-个人补缴', '失业保险-个人补缴', '医疗保险-个人补缴', '住房公积金-个人补缴', "大额保险-个人",
        #                     "大额保险-企业", "个人所得税", "实发工资",
        #                     "系统工资税基", '财务计税附加', '累计收入额', '累计减除费用',
        #                     '累计子女教育', '累计赡养老人', '累计继续教育', '累计住房贷款利息', '累计住房租金', '累计专项扣除',
        #                     "累计应纳税所得额", "累计应纳税额", "累计减免税额", "累计应扣缴税额", "累计已预缴税额", "累计应补(退)税额"
        #                     ]
        list_wage_detail = ["员工编号", "姓名", "身份证号码", "银行账号", "部室", "员工子组名称", '岗序', '薪等', '岗位工资系数', '预支年薪',
                            '岗位工资', "加班工资", '技能保留工资', '年功保留工资', '竞业津贴', '技术津贴',
                            '回民补贴', '女工洗理费', '住房补贴', '工作补贴', '交通补贴',
                            '班中餐补贴', '误餐补贴', '通讯补贴', '异地津贴', '异地差旅费',
                            '外租房津贴', "中班津贴", "夜班津贴", '其他补发', '其他补扣', "应发工资",
                            '养老保险员工实缴', '失业保险员工实缴', '医疗保险员工实缴', '住房公积金员工实缴',
                            '养老保险-个人补缴', '失业保险-个人补缴', '医疗保险-个人补缴', '住房公积金-个人补缴', "大额保险-个人",
                            "大额保险-企业", "个人所得税", "实发工资",
                            "系统工资税基",  '累计收入额', '累计减除费用',
                            '累计子女教育', '累计赡养老人', '累计继续教育', '累计住房贷款利息', '累计住房租金', '累计专项扣除',
                            "累计应纳税所得额", "累计应纳税额", "累计减免税额", "累计应扣缴税额", "累计已预缴税额", "累计应补(退)税额"
                            ]

        df_wage_detail = df_wage[list_wage_detail]

        str_file_path, res = QFileDialog.getSaveFileName(self.window, "保存文件至", r"d:\{}.xlsx"
                                                         .format("河钢乐亭{}工资明细".format(self.date)), "excel 文件(*.xlsx)")
        if res:
            writer = pd.ExcelWriter(str_file_path, engine='xlsxwriter')
            df_wage_detail.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            excel_format = workbook.add_format({'align': 'center', 'font_size': 9})
            worksheet.set_column("A:BI", 15.5, excel_format)
            writer.save()
        else:
            return

    def start(self):
        try:
            self.wageDetail()
        except xlsxwriter.exceptions.FileCreateError:
            QMessageBox.warning(self.window, "河钢乐亭薪酬管理系统", "同名文件正在打开，请关闭后重新生成。", QMessageBox.Yes)
        except Exception as e:
            QMessageBox.warning(self.window, "河钢乐亭薪酬管理系统", "发生错误{}".format(str(e)), QMessageBox.Yes)
        else:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "工资明细生成完毕！", QMessageBox.Ok)
