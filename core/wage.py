# -*-coding:utf-8-*-

import pandas as pd
from PyQt5.QtWidgets import QMessageBox, QInputDialog
import decimal
from lib.common import consoleFormat, floatReturnPercent, max0, tex
from lib.erorClass import DataFormatError


class Wage(object):
    """计算工资的主函数"""

    def __init__(self, window, date, workday, conn):
        # 设定控制台输出格式
        consoleFormat()
        self.window = window
        self.date = date
        self.conn = conn
        self.workday = workday

    def dfGetDataFromSql(self):

        # 第一步 检查数据库中各项表的数据是否合乎要求
        # 标准表导入，检查
        try:
            # 导入标准表
            df_standrad = pd.read_sql_query(r"select * from t{}StandradTable".format(self.date), self.conn)
            # 导入工资补发补扣
            df_wage_extra_workincome = \
                pd.read_sql_query(r"select * from t{}ExtraWorkincome".format(self.date), self.conn) \
                    [["员工编号", "其他补发", "其他补扣"]]

            # 导入各类保险,并把名称统一
            df_endowment_insurance_roster = \
                pd.read_sql_query(r"select * from t{}EndowmentInsuranceRoster".format(self.date), self.conn) \
                    [["个人编号", "证件号码"]]
            df_endowment_insurance_roster.rename(columns={"证件号码": "身份证号码", "个人编号": "社保个人编号"}, inplace=True)

            df_endowment_insurance_detail = \
                pd.read_sql_query(r"select * from t{}EndowmentInsuranceDetail".format(self.date), self.conn) \
                    [["个人编号", "个人缴"]]
            df_endowment_insurance_detail.rename(columns={"个人缴": "养老保险员工实缴", "个人编号": "社保个人编号"}, inplace=True)

            df_unemployement_detail = \
                pd.read_sql_query(r"select * from t{}UnemployementInsuranceDetail".format(self.date), self.conn) \
                    [["个人编号", "个人缴"]]
            df_unemployement_detail.rename(columns={"个人缴": "失业保险员工实缴", "个人编号": "社保个人编号"}, inplace=True)

            df_medical_insurance_roster = \
                pd.read_sql_query(r"select * from t{}MedicalInsuranceRoster".format(self.date), self.conn) \
                    [["个人编号", "证件号码"]]
            df_medical_insurance_roster.rename(columns={"证件号码": "身份证号码", "个人编号": "医保个人编号"}, inplace=True)

            df_medical_insurance_detail = \
                pd.read_sql_query(r"select * from t{}MedicalInsuranceDetail".format(self.date), self.conn) \
                    [["个人编号", "个人缴"]]
            df_medical_insurance_detail.rename(columns={"个人缴": "医疗保险员工实缴", "个人编号": "医保个人编号"}, inplace=True)

            df_house_insurance_Detail = \
                pd.read_sql_query(r"select * from t{}HouseInsuranceDetail".format(self.date), self.conn) \
                    [["证件号码", "个人月缴额"]]
            df_house_insurance_Detail.rename(columns={"证件号码": "身份证号码", "个人月缴额": "住房公积金员工实缴"}, inplace=True)

            df_civil_servants_insurance_detail = \
                pd.read_sql_query(r"select * from t{}CivilServantsInsuranceDetail".format(self.date), self.conn) \
                    [["个人编号", "险种类型"]]

            df_civil_servants_insurance_detail.rename(columns={"个人编号": "医保个人编号", "险种类型": "是否享受公务员医疗待遇"}, \
                                                      inplace=True)
            # 保险补缴信息
            df_insurance_extra_payment = pd.read_sql_query(r"select * from t{}InsuranceExtraPayment".format(self.date),
                                                           self.conn)
            df_endowment_insurance_extra_payment = \
                df_insurance_extra_payment[df_insurance_extra_payment["工资项"] == "4102"][["员工编号", "金额"]]
            df_endowment_insurance_extra_payment.rename(columns={"金额": "养老保险-个人补缴"}, inplace=True)

            df_unemployement_insurance_extra_payment = \
                df_insurance_extra_payment[df_insurance_extra_payment["工资项"] == "4122"][["员工编号", "金额"]]
            df_unemployement_insurance_extra_payment.rename(columns={"金额": "失业保险-个人补缴"}, inplace=True)

            df_medical_insurance_extra_payment = \
                df_insurance_extra_payment[df_insurance_extra_payment["工资项"] == "4112"][["员工编号", "金额"]]
            df_medical_insurance_extra_payment.rename(columns={"金额": "医疗保险-个人补缴"}, inplace=True)

            df_house_insurance_extra_payment = \
                df_insurance_extra_payment[df_insurance_extra_payment["工资项"] == "4092"][["员工编号", "金额"]]
            df_house_insurance_extra_payment.rename(columns={"金额": "住房公积金-个人补缴"}, inplace=True)

            # 财务计税附加
            """因为计税附加可能含有某一个职工多个工资项均有计数，所以需要把金额累计起来，因为数据库内各列格式均为str
            所以先把金额转换格式，然后按员工编号分组求和"""
            df_tex_extra = pd.read_sql_query(r"select * from t{}TaxExtra".format(self.date), self.conn)[
                ["员工编号", "工资项", "金额"]]
            df_tex_extra.rename(columns={"金额": "财务计税附加"}, inplace=True)
            df_tex_extra["财务计税附加"] = df_tex_extra["财务计税附加"].astype(float)
            print(df_tex_extra["财务计税附加"].dtypes)
            df_tex_extra = df_tex_extra.groupby("员工编号").sum()

            # 上月缴税明细信息导入
            df_last_month_tax_detail = \
                pd.read_sql_query(r"select * from t{}LastMonthTaxDetail".format(self.date), self.conn) \
                    [["证照号码", "累计收入额", "累计减除费用", "累计专项扣除", "累计应补(退)税额", "累计已预缴税额"]]
            df_last_month_tax_detail.rename(columns={"证照号码": "身份证号码"}, inplace=True)

            # 本月的个人附加扣除信息
            df_tex_remove = pd.read_sql_query(r"select * from t{}TaxRemove".format(self.date), self.conn) \
                [["证照号码", "累计子女教育", "累计赡养老人", "累计继续教育", "累计住房贷款利息", "累计住房租金"]]
            df_tex_remove.rename(columns={"证照号码": "身份证号码"}, inplace=True)

            # 出勤信息导入
            df_sap_workday = pd.read_sql_query(r"select * from t{}SapWorkday".format(self.date), self.conn) \
                [["人员编号", "年假天数", "事假天数", "病假天数", "工伤假天数", "探亲假天数", "婚假天数", "丧假天数", "产假天数",
                  "育儿假天数", "节育假天数", "旷工天数", "高温出勤", "有毒有害出勤", "工作补贴出勤", "节假日加班","加班天数", "中班天数", "夜班天数"]]
            df_sap_workday.rename(columns={"人员编号": "员工编号"}, inplace=True)

            df_people_info = pd.read_sql_query(r"select * from t{}PeopleInfo".format(self.date), self.conn) \
                [["员工编号", "部室", "员工子组名称", "性别", "合同起始时间"]]

            # 导入上月工资明细
            df_last_month_wage_detail = \
                pd.read_sql_query(r"select * from t{}LastMonthWageDetail".format(self.date), self.conn) \
                    [["员工编号", "岗位工资系数"]]
            df_last_month_wage_detail.rename(columns={"岗位工资系数": "上月岗位工资系数"}, inplace=True)

        except pd.io.sql.DatabaseError as e:
            raise DataFormatError(r"没有在数据库中找到所需的数据表，请确认是否已经导入。具体错误信息如下：{}"
                                  .format(str(e)))
        except ValueError as e:
            raise DataFormatError(r"你所导入的文件当中，应该为数字的某一列含有字母、字符等无法进行计算的内容，请核实。具体错误信息如下：{}"
                                  .format(str(e)))
        else:
            # 各种表依次整合，首先是标准表和补发补扣
            df_return = pd.merge(df_standrad, df_wage_extra_workincome, on="员工编号", how="left")
            # 其次是养老保险花名册
            df_return = pd.merge(df_return, df_endowment_insurance_roster, on="身份证号码", how="left")
            # 养老保险明细
            df_return = pd.merge(df_return, df_endowment_insurance_detail, on="社保个人编号", how="left")
            # 失业保险明细
            df_return = pd.merge(df_return, df_unemployement_detail, on="社保个人编号", how="left")
            # 医疗保险花名册
            df_return = pd.merge(df_return, df_medical_insurance_roster, on="身份证号码", how="left")
            # 医疗保险明细
            df_return = pd.merge(df_return, df_medical_insurance_detail, on="医保个人编号", how="left")
            # 公积金明细
            df_return = pd.merge(df_return, df_house_insurance_Detail, on="身份证号码", how="left")
            # 公务员明细
            df_return = pd.merge(df_return, df_civil_servants_insurance_detail, on="医保个人编号", how="left")
            # 养老保险补缴
            df_return = pd.merge(df_return, df_endowment_insurance_extra_payment, on="员工编号", how="left")
            # 失业保险补缴
            df_return = pd.merge(df_return, df_unemployement_insurance_extra_payment, on="员工编号", how="left")
            # 医疗保险补缴
            df_return = pd.merge(df_return, df_medical_insurance_extra_payment, on="员工编号", how="left")
            # 住房公积金补缴
            df_return = pd.merge(df_return, df_house_insurance_extra_payment, on="员工编号", how="left")
            # 财务计税附加
            df_return = pd.merge(df_return, df_tex_extra, on="员工编号", how="left")
            # 上月计税信息（累计收入额、累计减除、累计专项扣除）
            df_return = pd.merge(df_return, df_last_month_tax_detail, on="身份证号码", how="left")
            # 本月个人附加信息
            df_return = pd.merge(df_return, df_tex_remove, on="身份证号码", how="left")
            # sap出勤
            df_return = pd.merge(df_return, df_sap_workday, on="员工编号", how="left")
            # 人员信息
            df_return = pd.merge(df_return, df_people_info, on="员工编号", how="left")
            # 导入上月的工资系数
            df_return = pd.merge(df_return, df_last_month_wage_detail, on="员工编号", how="left")
            return df_return

    def dfChangeDataFormat(self, df_wage):

        df_wage.fillna("0", inplace=True)
        list_need_change_format = ['岗序', '薪等', '岗位工资系数', '奖金系数', '预支年薪',
                                   '岗位工资', '技能保留工资', '年功保留工资', '竞业津贴', '技术津贴',
                                   '回民补贴', '女工洗理费', '住房补贴', '工作补贴', '交通补贴',
                                   '班中餐补贴', '误餐补贴', '通讯补贴', '异地津贴', '异地差旅费',
                                   '外租房津贴', '其他补发', '其他补扣',
                                   '养老保险员工实缴', '失业保险员工实缴',
                                   '医疗保险员工实缴', '住房公积金员工实缴',
                                   '养老保险-个人补缴', '失业保险-个人补缴', '医疗保险-个人补缴',
                                   '住房公积金-个人补缴', '财务计税附加', '累计收入额', '累计减除费用',
                                   '累计专项扣除', '累计子女教育', '累计赡养老人', '累计继续教育',
                                   '累计住房贷款利息', '累计住房租金', '年假天数', '事假天数', '病假天数',
                                   '工伤假天数', '探亲假天数', '婚假天数', '丧假天数', '丧假天数',
                                   '产假天数', '育儿假天数', '节育假天数', '旷工天数', '高温出勤',
                                   '有毒有害出勤', '工作补贴出勤', '节假日加班', "加班天数",'中班天数', '夜班天数', "上月岗位工资系数"
            , "累计应补(退)税额", "累计已预缴税额"]
        # list_need_change_format = ['岗序', '薪等', '岗位工资系数', '奖金系数', '预支年薪',
        #                            '岗位工资', '技能保留工资', '年功保留工资', '竞业津贴', '技术津贴',
        #                            '回民补贴', '女工洗理费', '住房补贴', '工作补贴', '交通补贴',
        #                            '班中餐补贴', '误餐补贴', '通讯补贴', '异地津贴', '异地差旅费',
        #                            '外租房津贴', '其他补发', '其他补扣',
        #                            '养老保险员工实缴', '失业保险员工实缴',
        #                            '医疗保险员工实缴', '住房公积金员工实缴',
        #                            '养老保险-个人补缴', '失业保险-个人补缴', '医疗保险-个人补缴',
        #                            '住房公积金-个人补缴', '累计收入额', '累计减除费用',
        #                            '累计专项扣除', '累计子女教育', '累计赡养老人', '累计继续教育',
        #                            '累计住房贷款利息', '累计住房租金', '年假天数', '事假天数', '病假天数',
        #                            '工伤假天数', '探亲假天数', '婚假天数', '丧假天数', '丧假天数',
        #                            '产假天数', '育儿假天数', '节育假天数', '旷工天数', '高温出勤',
        #                            '有毒有害出勤', '工作补贴出勤', '节假日加班', '中班天数', '夜班天数', "上月岗位工资系数"
        #     , "累计应补(退)税额", "累计已预缴税额"]

        try:
            for str_col_name in list_need_change_format:
                if str_col_name in df_wage.columns:
                    df_wage[str_col_name] = df_wage[str_col_name].astype(float)
        except ValueError as e:
            raise DataFormatError(r"你所导入的文件当中，应该为全部为数值的某一列含有字母、字符等无法进行计算的内容，请核实。具体错误信息如下：{}"
                                  .format(str(e)))
        else:
            return df_wage

    def dfWageDataCalc(self, df_wage):
        # 把需要插入的列再开始就一次性插入
        df_wage.insert(13, "加班工资", 0.0)
        df_wage.insert(31, "应发工资", 0.0)
        df_wage.insert(29, "夜班班津贴", 0.0)
        df_wage.insert(29, "中班津贴", 0.0)
        df_wage.insert(41, "大额保险-个人", 0.0)
        df_wage.insert(42, "大额保险-企业", 0.0)
        df_wage.insert(43, "个人所得税", 0.0)
        df_wage.insert(48, "实发工资", 0.0)
        df_wage.insert(55, "系统工资税基", 0.0)
        df_wage.insert(56, "累计应纳税所得额", 0.0)
        df_wage.insert(57, "累计应纳税额", 0.0)
        df_wage.insert(58, "累计减免税额", 0.0)
        df_wage.insert(59, "累计应扣缴税额", 0.0)

        # 计算岗位工资
        df_wage["岗位工资"] = df_wage["岗位工资系数"] * 360

        # 计算加班工资，等于上月工资系数/21.75*360*3*加班天数再加上日常加班
        df_wage["加班工资"] = (df_wage["上月岗位工资系数"] / 21.75 * 3 * 360 * df_wage['节假日加班']).round(2)+(df_wage["上月岗位工资系数"] / 21.75 * 2 * 360 * df_wage['加班天数']).round(2)

        # 中夜班津补贴
        df_wage["中班津贴"] = df_wage["中班天数"] * 10
        df_wage["夜班津贴"] = df_wage["夜班天数"] * 12

        """下面计算各种假的扣减"""
        # 1 年休假扣减，不扣工资
        df_wage["扣竞业津贴（年休假）"] = (df_wage["竞业津贴"] / self.workday * df_wage["年假天数"]).round(2)
        df_wage["扣工作补贴（年休假）"] = (df_wage["工作补贴"] / self.workday * df_wage["年假天数"]).round(2)
        df_wage["扣交通补贴（年休假）"] = (df_wage["交通补贴"] / self.workday * df_wage["年假天数"]).round(2)
        df_wage["扣班中餐补贴（年休假）"] = (df_wage["班中餐补贴"] / self.workday * df_wage["年假天数"]).round(2)
        df_wage["扣误餐补贴（年休假）"] = (df_wage["误餐补贴"] / self.workday * df_wage["年假天数"]).round(2)

        # 产假，不扣工资
        df_wage["扣竞业津贴（产假）"] = (df_wage["竞业津贴"] / self.workday * df_wage["产假天数"]).round(2)
        df_wage["扣工作补贴（产假）"] = (df_wage["工作补贴"] / self.workday * df_wage["产假天数"]).round(2)
        df_wage["扣交通补贴（产假）"] = (df_wage["交通补贴"] / self.workday * df_wage["产假天数"]).round(2)
        df_wage["扣班中餐补贴（产假）"] = (df_wage["班中餐补贴"] / self.workday * df_wage["产假天数"]).round(2)
        df_wage["扣误餐补贴（产假）"] = (df_wage["误餐补贴"] / self.workday * df_wage["产假天数"]).round(2)
        df_wage["扣技术津贴（产假）"] = (df_wage["技术津贴"] / self.workday * df_wage["产假天数"]).round(2)
        df_wage["扣异地津贴（产假）"] = (df_wage["异地津贴"] / self.workday * df_wage["产假天数"]).round(2)
        df_wage["扣异地差旅费（产假）"] = (df_wage["异地差旅费"] / self.workday * df_wage["产假天数"]).round(2)
        df_wage["扣外租房津贴（产假）"] = (df_wage["外租房津贴"] / self.workday * df_wage["产假天数"]).round(2)

        # 探亲假，探亲假不扣工资
        df_wage["扣竞业津贴（探亲假）"] = (df_wage["竞业津贴"] / self.workday * df_wage["探亲假天数"]).round(2)
        df_wage["扣工作补贴（探亲假）"] = (df_wage["工作补贴"] / self.workday * df_wage["探亲假天数"]).round(2)
        df_wage["扣交通补贴（探亲假）"] = (df_wage["交通补贴"] / self.workday * df_wage["探亲假天数"]).round(2)
        df_wage["扣班中餐补贴（探亲假）"] = (df_wage["班中餐补贴"] / self.workday * df_wage["探亲假天数"]).round(2)
        df_wage["扣误餐补贴（探亲假）"] = (df_wage["误餐补贴"] / self.workday * df_wage["探亲假天数"]).round(2)
        df_wage["扣技术津贴（探亲假）"] = (df_wage["技术津贴"] / self.workday * df_wage["探亲假天数"]).round(2)

        # 婚假，不扣工资
        df_wage["扣竞业津贴（婚假）"] = (df_wage["竞业津贴"] / self.workday * df_wage["婚假天数"]).round(2)
        df_wage["扣工作补贴（婚假）"] = (df_wage["工作补贴"] / self.workday * df_wage["婚假天数"]).round(2)
        df_wage["扣交通补贴（婚假）"] = (df_wage["交通补贴"] / self.workday * df_wage["婚假天数"]).round(2)
        df_wage["扣班中餐补贴（婚假）"] = (df_wage["班中餐补贴"] / self.workday * df_wage["婚假天数"]).round(2)
        df_wage["扣误餐补贴（婚假）"] = (df_wage["误餐补贴"] / self.workday * df_wage["婚假天数"]).round(2)
        df_wage["扣技术津贴（婚假）"] = (df_wage["技术津贴"] / self.workday * df_wage["婚假天数"]).round(2)

        # 丧假，不扣工资
        df_wage["扣竞业津贴（丧假）"] = (df_wage["竞业津贴"] / self.workday * df_wage["丧假天数"]).round(2)
        df_wage["扣工作补贴（丧假）"] = (df_wage["工作补贴"] / self.workday * df_wage["丧假天数"]).round(2)
        df_wage["扣交通补贴（丧假）"] = (df_wage["交通补贴"] / self.workday * df_wage["丧假天数"]).round(2)
        df_wage["扣班中餐补贴（丧假）"] = (df_wage["班中餐补贴"] / self.workday * df_wage["丧假天数"]).round(2)
        df_wage["扣误餐补贴（丧假）"] = (df_wage["误餐补贴"] / self.workday * df_wage["丧假天数"]).round(2)
        df_wage["扣技术津贴（丧假）"] = (df_wage["技术津贴"] / self.workday * df_wage["丧假天数"]).round(2)

        # 事假
        df_wage["扣工资（事假）"] = (
                (df_wage["预支年薪"] + df_wage["岗位工资"] + df_wage["技能保留工资"] + df_wage["年功保留工资"]) / self.workday *
                df_wage["事假天数"]).round(2)
        df_wage["扣竞业津贴（事假）"] = (df_wage["竞业津贴"] / self.workday * df_wage["事假天数"]).round(2)
        df_wage["扣技术津贴（事假）"] = (df_wage["技术津贴"] / self.workday * df_wage["事假天数"]).round(2)
        df_wage["扣工作补贴（事假）"] = (df_wage["工作补贴"] / self.workday * df_wage["事假天数"]).round(2)
        df_wage["扣交通补贴（事假）"] = (df_wage["交通补贴"] / self.workday * df_wage["事假天数"]).round(2)
        df_wage["扣班中餐补贴（事假）"] = (df_wage["班中餐补贴"] / self.workday * df_wage["事假天数"]).round(2)
        df_wage["扣误餐补贴（事假）"] = (df_wage["误餐补贴"] / self.workday * df_wage["事假天数"]).round(2)

        # 旷工
        df_wage["扣工资（旷工）"] = (
                (df_wage["预支年薪"] + df_wage["岗位工资"] + df_wage["技能保留工资"] + df_wage["年功保留工资"]) / self.workday *
                df_wage["旷工天数"]).round(2)
        df_wage["扣竞业津贴（旷工）"] = (df_wage["竞业津贴"] / self.workday * df_wage["旷工天数"]).round(2)
        df_wage["扣技术津贴（旷工）"] = (df_wage["技术津贴"] / self.workday * df_wage["旷工天数"]).round(2)
        df_wage["扣工作补贴（旷工）"] = (df_wage["工作补贴"] / self.workday * df_wage["旷工天数"]).round(2)
        df_wage["扣交通补贴（旷工）"] = (df_wage["交通补贴"] / self.workday * df_wage["旷工天数"]).round(2)
        df_wage["扣班中餐补贴（旷工）"] = (df_wage["班中餐补贴"] / self.workday * df_wage["旷工天数"]).round(2)
        df_wage["扣误餐补贴（旷工）"] = (df_wage["误餐补贴"] / self.workday * df_wage["旷工天数"]).round(2)

        # 病假
        df_wage["扣竞业津贴（病假）"] = (df_wage["竞业津贴"] / self.workday * df_wage["病假天数"]).round(2)
        df_wage["扣工作补贴（病假）"] = (df_wage["工作补贴"] / self.workday * df_wage["病假天数"]).round(2)
        df_wage["扣交通补贴（病假）"] = (df_wage["交通补贴"] / self.workday * df_wage["病假天数"]).round(2)
        df_wage["扣班中餐补贴（病假）"] = (df_wage["班中餐补贴"] / self.workday * df_wage["病假天数"]).round(2)
        df_wage["扣误餐补贴（病假）"] = (df_wage["误餐补贴"] / self.workday * df_wage["病假天数"]).round(2)
        df_wage["扣技术津贴（病假）"] = (df_wage["技术津贴"] / self.workday * df_wage["病假天数"]).round(2)

        # 病假扣工资比较复杂，需要先判断职工工作年限，确定比例然后再计算
        df_wage["扣工资（病假）"] = 0.0
        for i in df_wage["病假天数"].index:
            if df_wage["病假天数"].at[i]:
                while True:
                    year, ok = QInputDialog.getInt(self.window, '河钢乐亭薪酬管理系统',
                                                   "员工编号为{}的员工{}有病假，\n\r\t请输入其工龄".format(df_wage["员工编号"].at[i],
                                                                                         df_wage["姓名"].at[i]), min=0)
                    percent = floatReturnPercent(year)
                    if ok:
                        df_wage["扣工资（病假）"].at[i] = ((df_wage["岗位工资"].at[i] + df_wage["技能保留工资"].at[i]
                                                     + df_wage["年功保留工资"].at[i]) / self.workday * percent *
                                                    df_wage["病假天数"].at[i]).round(2)
                        break
                    else:
                        QMessageBox.warning(self.window, "河钢乐亭薪酬管理系统", "取消后无法进行病假计算，请重新输入！", QMessageBox.Ok)

        """工作补贴也比较复杂
        一 工作补贴出勤天数大于0：
            1 工作补贴出勤天数大于应出勤天数 =工作补贴标准
            2 不大于应出勤天数 =工作补贴标准/应出勤天数*工作补贴标准
        二 工作补贴等于0
            工作补贴等于标准减去工作补贴扣减
        三 如果有异地津补贴，那么说明是宣钢职工，工作补贴为0
        """
        for i in df_wage["工作补贴"].index:
            if df_wage["工作补贴出勤"].at[i] > 0:
                if df_wage["工作补贴出勤"].at[i] > self.workday:
                    pass
                else:
                    df_wage["工作补贴"].at[i] = (df_wage["工作补贴"].at[i] / self.workday * df_wage["工作补贴出勤"].at[i]).round(2)

            else:
                df_wage["工作补贴"].at[i] = df_wage["工作补贴"].at[i] \
                                        - df_wage["扣工作补贴（产假）"].at[i] - df_wage["扣工作补贴（病假）"].at[i] \
                                        - df_wage["扣工作补贴（年休假）"].at[i] - df_wage["扣工作补贴（探亲假）"].at[i] \
                                        - df_wage["扣工作补贴（婚假）"].at[i] - df_wage["扣工作补贴（丧假）"].at[i] \
                                        - df_wage["扣工作补贴（事假）"].at[i] - df_wage["扣工作补贴（旷工）"].at[i]
            # # 如果有异地差旅费，说明是宣钢职工
            # if df_wage["异地津贴"].at[i]:
            #     df_wage["工作补贴"].at[i] = 0

        # 处理其他津贴项
        # 根据以上的各种假扣减来最终计算各项补贴金额
        df_wage["竞业津贴"] = df_wage["竞业津贴"] - df_wage["扣竞业津贴（年休假）"] - df_wage["扣竞业津贴（产假）"] \
                          - df_wage["扣竞业津贴（病假）"] - df_wage["扣竞业津贴（丧假）"] - df_wage["扣竞业津贴（婚假）"] \
                          - df_wage["扣竞业津贴（探亲假）"] - df_wage["扣竞业津贴（事假）"] - df_wage["扣竞业津贴（旷工）"]
        df_wage["交通补贴"] = df_wage["交通补贴"] - df_wage["扣交通补贴（年休假）"] - df_wage["扣交通补贴（产假）"] \
                          - df_wage["扣交通补贴（病假）"] - df_wage["扣交通补贴（丧假）"] - df_wage["扣交通补贴（婚假）"] \
                          - df_wage["扣交通补贴（探亲假）"] - df_wage["扣交通补贴（事假）"] - df_wage["扣交通补贴（旷工）"]
        df_wage["班中餐补贴"] = df_wage["班中餐补贴"] - df_wage["扣班中餐补贴（年休假）"] - df_wage["扣班中餐补贴（产假）"] \
                           - df_wage["扣班中餐补贴（病假）"] - df_wage["扣班中餐补贴（丧假）"] - df_wage["扣班中餐补贴（婚假）"] \
                           - df_wage["扣班中餐补贴（探亲假）"] - df_wage["扣班中餐补贴（事假）"] - df_wage["扣班中餐补贴（旷工）"]
        df_wage["误餐补贴"] = df_wage["误餐补贴"] - df_wage["扣误餐补贴（年休假）"] - df_wage["扣误餐补贴（产假）"] \
                          - df_wage["扣误餐补贴（病假）"] - df_wage["扣误餐补贴（丧假）"] - df_wage["扣误餐补贴（婚假）"] \
                          - df_wage["扣误餐补贴（探亲假）"] - df_wage["扣误餐补贴（事假）"] - df_wage["扣误餐补贴（旷工）"]
        df_wage["技术津贴"] = df_wage["技术津贴"] - df_wage["扣技术津贴（产假）"] \
                          - df_wage["扣技术津贴（病假）"] - df_wage["扣技术津贴（丧假）"] - df_wage["扣技术津贴（婚假）"] \
                          - df_wage["扣技术津贴（探亲假）"] - df_wage["扣技术津贴（事假）"] - df_wage["扣技术津贴（旷工）"]
        # df_wage["其他补扣"] = (df_wage["其他补扣"] + df_wage["扣工资（病假）"]) + df_wage["扣工资（事假）"] + df_wage["扣工资（旷工）"]
        df_wage["异地津贴"] = df_wage["异地津贴"] - df_wage["扣异地津贴（产假）"]
        df_wage["异地差旅费"] = df_wage["异地差旅费"] - df_wage["扣异地差旅费（产假）"]
        df_wage["外租房津贴"] = df_wage["外租房津贴"] - df_wage["扣外租房津贴（产假）"]

        # 增加一步处理过程，各项津贴最多扣至0，不能扣成负数
        df_wage["工作补贴"] = df_wage["工作补贴"].apply(max0)
        df_wage["竞业津贴"] = df_wage["竞业津贴"].apply(max0)
        df_wage["交通补贴"] = df_wage["交通补贴"].apply(max0)
        df_wage["班中餐补贴"] = df_wage["班中餐补贴"].apply(max0)
        df_wage["误餐补贴"] = df_wage["误餐补贴"].apply(max0)
        df_wage["技术津贴"] = df_wage["技术津贴"].apply(max0)
        df_wage["异地津贴"] = df_wage["异地津贴"].apply(max0)
        df_wage["异地差旅费"] = df_wage["异地差旅费"].apply(max0)
        df_wage["外租房津贴"] = df_wage["外租房津贴"].apply(max0)

        # 计算应发
        df_wage["应发工资"] = df_wage["预支年薪"] + df_wage["岗位工资"] + df_wage["加班工资"] + df_wage["技能保留工资"] + df_wage["年功保留工资"] \
                          + df_wage["竞业津贴"] + df_wage["技术津贴"] + df_wage["回民补贴"] + df_wage["女工洗理费"] + df_wage["住房补贴"] + \
                          df_wage["工作补贴"] + df_wage["班中餐补贴"] \
                          + df_wage["交通补贴"] + df_wage["误餐补贴"] + df_wage["通讯补贴"] + df_wage["异地津贴"] + df_wage["异地差旅费"] + \
                          df_wage["外租房津贴"] \
                          + df_wage["中班津贴"] + df_wage["夜班津贴"] - df_wage["扣工资（病假）"]
        # todo 这里有个问题，在计算产假和育儿假工资的时候计算应发还有问题，需要进一步查找问题所在
        """下面考虑一下最低工资标准的80%如何处理
        想法为：除了事假和旷工，其他的假都不能低于1900的80%，即1520.
        那么，，如果应发工资减去df_wage["扣工资（病假）"]（除了矿工事假，其他假都不扣工资）低于1520，那么应发工资就等于1520
        然后处理后的应发工资减去事假、旷工扣款
        """

        # 对应发工资进行判断，是否小于1520，同时要考虑只在这里交保险的人
        # 下面这个列表的人都是只在这边缴费的人，对于他们，不适用最低工资标准
        list_need_check_all_pay = ["9000160", "90001302"]  # 这个列表是代交保险人员，以后如果需要增加还得手动从这里增加
        for i in df_wage.index:
            if not df_wage["员工编号"].at[i] in list_need_check_all_pay:
                df_wage["应发工资"].at[i] = 1520 if df_wage["应发工资"].at[i] <= 1520 else df_wage["应发工资"].at[i]

        df_wage["其他补发"] = df_wage["其他补发"].round(2)
        df_wage["其他补扣"] = df_wage["其他补扣"].round(2)
        # 再次对应发工资进行处理，减去事假、旷工扣款、其他补扣，加上其他补发
        df_wage["应发工资"] = df_wage["应发工资"] - df_wage["扣工资（事假）"] - df_wage["扣工资（旷工）"] - df_wage["其他补扣"] + df_wage["其他补发"]

        # 再次处理，如果应发小于0，那么使其等于0
        # df_wage["应发工资"] = df_wage["应发工资"].apply(max0)

        # 计算大额保险
        df_wage.loc[df_wage["是否享受公务员医疗待遇"] != "参照公务员补助医疗保险", "大额保险-个人"] = 2
        df_wage.loc[df_wage["是否享受公务员医疗待遇"] != "参照公务员补助医疗保险", "大额保险-企业"] = 9

        """下面开始计算个人所得税，完全按照财务系统的逻辑来进行计算"""
        print(df_wage["财务计税附加"].dtypes)
        df_wage["系统工资税基"] = df_wage["应发工资"] - df_wage["大额保险-个人"] - df_wage["交通补贴"] * 0.7 + df_wage["财务计税附加"]
        # df_wage["系统工资税基"] = df_wage["应发工资"] - df_wage["大额保险-个人"] - df_wage["交通补贴"] * 0.7

        df_wage["累计收入额"] = df_wage["累计收入额"] + df_wage["系统工资税基"]

        df_wage["累计减除费用"] = df_wage["累计减除费用"] + 5000

        df_wage["累计专项扣除"] = df_wage["累计专项扣除"] + df_wage["养老保险员工实缴"] + df_wage["失业保险员工实缴"] + df_wage["医疗保险员工实缴"] + \
                            df_wage["住房公积金员工实缴"] \
                            + df_wage["养老保险-个人补缴"] + df_wage["医疗保险-个人补缴"] + df_wage["失业保险-个人补缴"] + df_wage["住房公积金-个人补缴"]

        df_wage["累计应纳税所得额"] = df_wage["累计收入额"] - df_wage["累计减除费用"] - df_wage["累计专项扣除"] - df_wage["累计子女教育"] - df_wage[
            "累计赡养老人"] - df_wage[
                                  "累计继续教育"] - df_wage["累计住房贷款利息"] - df_wage["累计住房租金"]

        df_wage["累计应纳税所得额"] = df_wage["累计应纳税所得额"].apply(max0)

        df_wage["累计应纳税额"] = df_wage["累计应纳税所得额"].apply(tex).round(2)

        for i in df_wage["员工编号"].index:
            if df_wage["员工编号"].at[i] in ["9000102", "9000161", "9000367"]:
                df_wage["累计减免税额"].at[i] = (df_wage["累计应纳税额"].at[i] * 0.8).round(2)

        df_wage["累计应扣缴税额"] = df_wage["累计应纳税额"] - df_wage["累计减免税额"]

        df_wage["累计已预缴税额"] = df_wage["累计已预缴税额"] + df_wage["累计应补(退)税额"]
        df_wage["累计应补(退)税额"] = (df_wage["累计应扣缴税额"] - df_wage["累计已预缴税额"]).apply(max0)

        # 重新计算实发工资，减去个税
        df_wage["实发工资"] = df_wage["应发工资"] - df_wage["累计应补(退)税额"] - df_wage["养老保险员工实缴"] - df_wage["失业保险员工实缴"] - df_wage[
            "医疗保险员工实缴"] - df_wage["住房公积金员工实缴"] - df_wage["养老保险-个人补缴"] - df_wage["医疗保险-个人补缴"] - df_wage["失业保险-个人补缴"] - \
                          df_wage["住房公积金-个人补缴"] - df_wage["大额保险-个人"]
        df_wage["实发工资"] = df_wage["实发工资"].round(2)

        # df_wage["实发工资"] = df_wage["实发工资"].apply(max0)
        df_wage["个人所得税"] = df_wage["累计应补(退)税额"]

        df_wage.to_excel("工资明细表.xlsx")
        return df_wage

    def wageDataCheck(self, df_wage):
        # 一、检查各项保险数值与社保系统中的数据是否相符
        # todo

        float_endowment_insurance = df_wage["养老保险员工实缴"].sum().round(2)
        float_unemployement_insurance = df_wage["失业保险员工实缴"].sum().round(2)
        float_medical_insurance = df_wage["医疗保险员工实缴"].sum().round(2)
        float_hourse_insurance = df_wage["住房公积金员工实缴"].sum().round(2)
        float_big_ill_insurance = df_wage["大额保险-个人"].sum().round(2)
        float_tex = df_wage["个人所得税"].sum()
        float_endowment_insurance_extra = df_wage["养老保险-个人补缴"].sum().round(2)
        float_unemployement_insurance_extra = df_wage["失业保险-个人补缴"].sum().round(2)
        float_medical_insurance_extra = df_wage["医疗保险-个人补缴"].sum().round(2)
        float_hourse_insurance_extra = df_wage["住房公积金-个人补缴"].sum().round(2)

        # 下面将数据库中各类保险数值导入计算
        # 养老
        df_endowment_insurance = pd.read_sql_query("select 个人缴 from t{}EndowmentInsuranceDetail".format(self.date),
                                                   self.conn)
        float_endowment_insurance_2 = df_endowment_insurance["个人缴"].astype(float).sum()
        # 失业
        df_unenployement_insurance = pd.read_sql_query(
            "select 个人缴 from t{}UnemployementInsuranceDetail".format(self.date), self.conn)
        float_unemployement_insurance_2 = df_unenployement_insurance["个人缴"].astype(float).sum()
        # 医疗
        df_medical_insurance = pd.read_sql_query(
            "select 个人缴 from t{}MedicalInsuranceDetail".format(self.date), self.conn)
        float_medical_insurance_2 = df_medical_insurance["个人缴"].astype(float).sum()

        # 公积金
        df_hourse_insurance = pd.read_sql_query(
            "select 个人月缴额 from t{}HouseInsuranceDetail".format(self.date), self.conn)
        float_hourse_insurance_2 = df_hourse_insurance["个人月缴额"].astype(float).sum()

        # 各类补缴
        df_extra_insurance = pd.read_sql_query(
            "select 金额 from t{}InsuranceExtraPayment".format(self.date), self.conn)
        float_extra_insurance = df_extra_insurance["金额"].astype(float).sum()

        list_all_wage_lower_then_0 = df_wage.loc[df_wage["应发工资"] < 0, "员工编号"].values
        list_in_hand_wage_lower_then_0 = df_wage.loc[df_wage["实发工资"] < 0, "员工编号"].values
        print(list_in_hand_wage_lower_then_0)

        # todo 应发工资小于0会造成实发工资小于0，进而造成工资汇总表上应发-各项扣缴！=实发工资，需要进行检查处理
        if len(list_all_wage_lower_then_0) > 0:
            raise DataFormatError("你可以继续生成明细表，汇总表。但请注意，以下员工{}的应发工资小于0，请检查后重新生成工资明细。".format(list_all_wage_lower_then_0))

        # todo  实发工资不能小于0，应该为错误，进行检查
        if len(list_in_hand_wage_lower_then_0) > 0:
            raise DataFormatError("你可以继续生成明细表，汇总表。但请注意，以下员工{}的实发工资小于0，请检查后重新生成工资明细以避免汇总表数据与明细表不一致。".format(
                list_in_hand_wage_lower_then_0))

        df_wage_extra_insurance = (
                float_endowment_insurance_extra + float_unemployement_insurance_extra + float_medical_insurance_extra + float_hourse_insurance_extra).round(
            2)
        if df_wage_extra_insurance != float_extra_insurance:
            raise DataFormatError(
                "你可以继续生成明细表，汇总表。但请注意，工资明细表中的保险补缴额{}与保险补缴导入文件中的金额{}不一致，请查明原因。".format(df_wage_extra_insurance,
                                                                                     float_extra_insurance))
        if float_endowment_insurance.round() != float_endowment_insurance_2.round():
            print("养老", float_endowment_insurance.round(), float_endowment_insurance_2.round())
            raise DataFormatError(
                "你可以继续生成明细表，汇总表。但请注意，工资明细表中的养老保险缴纳数为{}与养老保险缴纳明细表中的金额{}不一致，请查明原因。".format(float_endowment_insurance,
                                                                                         float_endowment_insurance_2))
        elif float_unemployement_insurance.round() != float_unemployement_insurance_2.round():
            raise DataFormatError(
                "你可以继续生成明细表，汇总表。但请注意，工资明细表中的失业保险缴纳数{}与失业保险缴纳明细表中的金额{}不一致，请查明原因。".format(float_unemployement_insurance,
                                                                                        float_unemployement_insurance_2))

        elif float_medical_insurance.round() != float_medical_insurance_2.round():
            raise DataFormatError(
                "你可以继续生成明细表，汇总表。但请注意，工资明细表中的医疗保险缴纳数{}与医疗保险缴纳明细表中的金额{}不一致，请查明原因。".format(float_medical_insurance,
                                                                                        float_medical_insurance_2))

        elif float_hourse_insurance.round() != float_hourse_insurance_2.round():
            raise DataFormatError(
                "你可以继续生成明细表，汇总表。但请注意，工资明细表中的医疗保险缴纳数{}与医疗保险缴纳明细表中的金额{}不一致，请查明原因。".format(float_hourse_insurance,
                                                                                        float_hourse_insurance_2))

    def wage(self):
        """ 工资计算主函数 """
        df_wage = self.dfGetDataFromSql()

        # 第二步 将所需要的数据全部取出,将数据格式统一调整至float64
        df_wage = self.dfChangeDataFormat(df_wage)

        # 第三步 将数据进行加工
        df_wage = self.dfWageDataCalc(df_wage)

        # 第四部 将生成的表格输入到数据库
        df_wage.to_sql("wage{}sql".format(self.date), self.conn, if_exists="replace")

        # 第五步，数据检查
        self.wageDataCheck(df_wage)

    def start(self):
        try:
            self.wage()
        except DataFormatError as e:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", str(e), QMessageBox.Ok)
        except Exception as e:
            print(e)
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "发生错误，内容如下：{}".format(str(e)), QMessageBox.Ok,
                                    QMessageBox.Ok)


        else:
            QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "工资数据已经导入数据库，请进行下一步操作！", QMessageBox.Ok)
