# -*-coding:utf-8-*-
import os
import shutil

import openpyxl
import pandas as pd
from PyQt5.QtWidgets import QMessageBox, QFileDialog

from lib.common import consoleFormat
from lib.erorClass import DataFormatError
from lib.moneyCHangeToChinese import strCHangeMoneyFormat


class WageSummary(object):
    def __init__(self, window, conn, date):
        consoleFormat()
        self.window = window
        self.conn = conn
        self.date = date

    def strReturnFileName(self):
        file_path = QFileDialog.getExistingDirectory(self.window, "工资明细保存至此文件夹", r"C:\Users\Administrator\Desktop")
        if file_path != "":
            file_name = os.path.join(file_path, self.date + "工资汇总表.xlsx")
            return file_name
        else:
            return False

    def excelWageSummary(self, str_file_name):

        df_wage_detail = pd.read_sql_query(r"select * from wage{}sql".format(self.date), self.conn)
        dict_people_count = df_wage_detail["员工子组名称"].value_counts().to_dict()
        int_high_position_count = dict_people_count["经营管理-高层"]
        int_middle_position_count = dict_people_count["经营管理-中层"]
        int_other_position_count = dict_people_count["经营管理-其他"]
        int_technology_people_count = dict_people_count["专业技术人员"]
        int_produce_people_count = dict_people_count["生产操作人员"]
        int_extra_people_count = dict_people_count["0"]

        # 岗位工资
        float_JingYingGuanLi_GWGZ = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), ["预支年薪", "岗位工资", "技能保留工资",
                                                                               "年功保留工资"]].sum().sum()
        float_ZHuanYeJiSHu_GWGZ = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "专业技术人员", ["预支年薪", "岗位工资",
                                                                                            "技能保留工资",
                                                                                            "年功保留工资"]].sum().sum()
        float_ShengChanCaoZuo_GWGZ = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "生产操作人员", ["预支年薪", "岗位工资",
                                                                                               "技能保留工资",
                                                                                               "年功保留工资"]].sum().sum()
        float_QiTaRenYuan_GWGZ = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "0", ["预支年薪", "岗位工资",
                                                                                      "技能保留工资", "年功保留工资"]].sum().sum()

        # 加班工资
        float_JingYingGuanLi_JBGZ = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), "加班工资"].sum().sum()

        float_ZHuanYeJiSHu_JBGZ = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "专业技术人员", "加班工资"].sum().sum()
        float_ShengChanCaoZuo_JBGZ = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "生产操作人员", "加班工资"].sum().sum()
        float_QiTaRenYuan_JBGZ = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "0", "加班工资"].sum().sum()

        # 岗位补贴
        float_JingYingGuanLi_GWBT = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), ["竞业津贴", "技术津贴", "回民补贴", "女工洗理费",
                                                                               "住房补贴", "中班津贴",
                                                                               "夜班津贴"]].sum().sum()

        float_ZHuanYeJiSHu_GWBT = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "专业技术人员", ["竞业津贴", "技术津贴", "回民补贴", "女工洗理费", "住房补贴", "中班津贴",
                                                   "夜班津贴"]].sum().sum()
        float_ShengChanCaoZuo_GWBT = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "生产操作人员", ["竞业津贴", "技术津贴", "回民补贴", "女工洗理费", "住房补贴", "中班津贴",
                                                   "夜班津贴"]].sum().sum()
        float_QiTaRenYuan_GWBT = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "0", ["竞业津贴", "技术津贴", "回民补贴", "女工洗理费", "住房补贴", "中班津贴",
                                              "夜班津贴"]].sum().sum()

        # 补发
        float_JingYingGuanLi_BuFa = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), "其他补发"].sum().sum()

        float_ZHuanYeJiSHu_BuFa = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "专业技术人员", "其他补发"].sum().sum()
        float_ShengChanCaoZuo_BuFa = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "生产操作人员", "其他补发"].sum().sum().round(
            2)
        float_QiTaRenYuan_BuFa = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "0", "其他补发"].sum().sum()

        # 补扣
        float_JingYingGuanLi_BuKou = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), "其他补扣"].sum().sum()

        float_ZHuanYeJiSHu_BuKou = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "专业技术人员", "其他补扣"].sum().sum()
        float_ShengChanCaoZuo_BuKou = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "生产操作人员", "其他补扣"].sum().sum()
        float_QiTaRenYuan_BuKou = df_wage_detail.loc[df_wage_detail["员工子组名称"] == "0", "其他补扣"].sum().sum()

        # 班中餐补贴
        float_JingYingGuanLi_BanZHongCan = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), ["班中餐补贴", "误餐补贴"]].sum().sum()

        float_ZHuanYeJiSHu_BanZHongCan = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "专业技术人员", ["班中餐补贴", "误餐补贴"]].sum().sum()
        float_ShengChanCaoZuo_BanZHongCan = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "生产操作人员", ["班中餐补贴", "误餐补贴"]].sum().sum()
        float_QiTaRenYuan_BanZHongCan = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "0", ["班中餐补贴", "误餐补贴"]].sum().sum()

        # 交通补贴
        float_JingYingGuanLi_JiaoTong = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), ["交通补贴"]].sum().sum()

        float_ZHuanYeJiSHu_JiaoTong = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "专业技术人员", ["交通补贴"]].sum().sum()
        float_ShengChanCaoZuo_JiaoTong = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "生产操作人员", ["交通补贴"]].sum().sum()
        float_QiTaRenYuan_JiaoTong = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "0", ["交通补贴"]].sum().sum()

        # 异地津贴
        float_JingYingGuanLi_YiDiJinTie = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), ["异地津贴", "工作补贴"]].sum().sum()

        float_ZHuanYeJiSHu_YiDiJinTie = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "专业技术人员", ["异地津贴", "工作补贴"]].sum().sum()
        float_ShengChanCaoZuo_YiDiJinTie = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "生产操作人员", ["异地津贴", "工作补贴"]].sum().sum()
        float_QiTaRenYuan_YiDiJinTie = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "0", ["异地津贴", "工作补贴"]].sum().sum()

        # 异地差旅费
        float_JingYingGuanLi_YiDiCHaiLueFei = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), ["异地差旅费"]].sum().sum()

        float_ZHuanYeJiSHu_YiDiCHaiLueFei = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "专业技术人员", ["异地差旅费"]].sum().sum()
        float_ShengChanCaoZuo_YiDiCHaiLueFei = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "生产操作人员", ["异地差旅费"]].sum().sum()
        float_QiTaRenYuan_YiDiCHaiLueFei = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "0", ["异地差旅费"]].sum().sum()

        # 外租房
        float_JingYingGuanLi_WaiZuFang = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), ["外租房津贴"]].sum().sum()

        float_ZHuanYeJiSHu_WaiZuFang = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "专业技术人员", ["外租房津贴"]].sum().sum()
        float_ShengChanCaoZuo_WaiZuFang = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "生产操作人员", ["外租房津贴"]].sum().sum()
        float_QiTaRenYuan_WaiZuFang = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "0", ["外租房津贴"]].sum().sum()

        # 应发工资
        float_JingYingGuanLi_YingFa = df_wage_detail.loc[
            df_wage_detail["员工子组名称"].isin(["经营管理-高层", "经营管理-中层", "经营管理-其他"]), ["应发工资"]].sum().sum()

        float_ZHuanYeJiSHu_YingFa = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "专业技术人员", ["应发工资"]].sum().sum()
        float_ShengChanCaoZuo_YingFa = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "生产操作人员", ["应发工资"]].sum().sum()
        float_QiTaRenYuan_YingFa = df_wage_detail.loc[
            df_wage_detail["员工子组名称"] == "0", ["应发工资"]].sum().sum()

        # 各类保险

        float_endowment_insurance = df_wage_detail["养老保险员工实缴"].sum()
        float_unemployement_insurance = df_wage_detail["失业保险员工实缴"].sum()
        float_medical_insurance = df_wage_detail["医疗保险员工实缴"].sum()
        float_hourse_insurance = df_wage_detail["住房公积金员工实缴"].sum()
        float_big_ill_insurance = df_wage_detail["大额保险-个人"].sum()
        float_tex = df_wage_detail["个人所得税"].sum()
        float_endowment_insurance_extra = df_wage_detail["养老保险-个人补缴"].sum()
        float_unemployement_insurance_extra = df_wage_detail["失业保险-个人补缴"].sum()
        float_medical_insurance_extra = df_wage_detail["医疗保险-个人补缴"].sum()
        float_hourse_insurance_extra = df_wage_detail["住房公积金-个人补缴"].sum()

        float_in_hand_money = df_wage_detail["实发工资"].sum()

        workbook = openpyxl.load_workbook(str_file_name, data_only=True)
        worksheet = workbook.active

        worksheet["a1"] = "河钢乐亭钢铁有限公司{}工资汇总表".format(self.date[:4] + "年" + self.date[4:] + "月")
        # 各类人数
        worksheet["c6"] = int_high_position_count + int_middle_position_count + int_other_position_count
        worksheet["c7"] = int_technology_people_count
        worksheet["c8"] = int_produce_people_count
        worksheet["c18"] = int_extra_people_count
        worksheet["c5"] = worksheet["c6"].value + worksheet["c7"].value + worksheet["c8"].value + worksheet["c18"].value

        # 基础工资
        worksheet["d6"] = float_JingYingGuanLi_GWGZ
        worksheet["d7"] = float_ZHuanYeJiSHu_GWGZ
        worksheet["d8"] = float_ShengChanCaoZuo_GWGZ
        worksheet["d18"] = float_QiTaRenYuan_GWGZ
        worksheet["d5"] = worksheet["d6"].value + worksheet["d7"].value + worksheet["d8"].value + worksheet["d18"].value

        # 加班工资
        worksheet["e6"] = float_JingYingGuanLi_JBGZ
        worksheet["e7"] = float_ZHuanYeJiSHu_JBGZ
        worksheet["e8"] = float_ShengChanCaoZuo_JBGZ
        worksheet["e18"] = float_QiTaRenYuan_JBGZ
        worksheet["e5"] = worksheet["e6"].value + worksheet["e7"].value + worksheet["e8"].value + worksheet["e18"].value

        # 岗位补贴
        worksheet["f6"] = float_JingYingGuanLi_GWBT
        worksheet["f7"] = float_ZHuanYeJiSHu_GWBT
        worksheet["f8"] = float_ShengChanCaoZuo_GWBT
        worksheet["f18"] = float_QiTaRenYuan_GWBT
        worksheet["f5"] = worksheet["f6"].value + worksheet["f7"].value + worksheet["f8"].value + worksheet["f18"].value

        # 补发工资
        worksheet["g6"] = float_JingYingGuanLi_BuFa
        worksheet["g7"] = float_ZHuanYeJiSHu_BuFa
        worksheet["g8"] = float_ShengChanCaoZuo_BuFa
        worksheet["g18"] = float_QiTaRenYuan_BuFa
        worksheet["g5"] = worksheet["g6"].value + worksheet["g7"].value + worksheet["g8"].value + worksheet["g18"].value

        # 补扣
        worksheet["h6"] = float_JingYingGuanLi_BuKou
        worksheet["h7"] = float_ZHuanYeJiSHu_BuKou
        worksheet["h8"] = float_ShengChanCaoZuo_BuKou
        worksheet["h18"] = float_QiTaRenYuan_BuKou
        worksheet["h5"] = worksheet["h6"].value + worksheet["h7"].value + worksheet["h8"].value + worksheet["h18"].value

        # 班中餐，含误餐
        worksheet["i6"] = float_JingYingGuanLi_BanZHongCan
        worksheet["i7"] = float_ZHuanYeJiSHu_BanZHongCan
        worksheet["i8"] = float_ShengChanCaoZuo_BanZHongCan
        worksheet["i18"] = float_QiTaRenYuan_BanZHongCan
        worksheet["i5"] = worksheet["i6"].value + worksheet["i7"].value + worksheet["i8"].value + worksheet["i18"].value

        # 交通补贴
        worksheet["j6"] = float_JingYingGuanLi_JiaoTong
        worksheet["j7"] = float_ZHuanYeJiSHu_JiaoTong
        worksheet["j8"] = float_ShengChanCaoZuo_JiaoTong
        worksheet["j18"] = float_QiTaRenYuan_JiaoTong
        worksheet["j5"] = worksheet["j6"].value + worksheet["j7"].value + worksheet["j8"].value + worksheet["j18"].value

        # 异地津补贴（含工作补贴）
        worksheet["k6"] = float_JingYingGuanLi_YiDiJinTie
        worksheet["k7"] = float_ZHuanYeJiSHu_YiDiJinTie
        worksheet["k8"] = float_ShengChanCaoZuo_YiDiJinTie
        worksheet["k18"] = float_QiTaRenYuan_YiDiJinTie
        worksheet["k5"] = worksheet["k6"].value + worksheet["k7"].value + worksheet["k8"].value + worksheet["k18"].value

        # 异地差旅费
        worksheet["l6"] = float_JingYingGuanLi_YiDiCHaiLueFei
        worksheet["l7"] = float_ZHuanYeJiSHu_YiDiCHaiLueFei
        worksheet["l8"] = float_ShengChanCaoZuo_YiDiCHaiLueFei
        worksheet["l18"] = float_QiTaRenYuan_YiDiCHaiLueFei
        worksheet["l5"] = worksheet["l6"].value + worksheet["l7"].value + worksheet["l8"].value + worksheet["l18"].value

        # 外租房
        worksheet["m6"] = float_JingYingGuanLi_WaiZuFang
        worksheet["m7"] = float_ZHuanYeJiSHu_WaiZuFang
        worksheet["m8"] = float_ShengChanCaoZuo_WaiZuFang
        worksheet["m18"] = float_QiTaRenYuan_WaiZuFang
        worksheet["m5"] = worksheet["m6"].value + worksheet["m7"].value + worksheet["m8"].value + worksheet["m18"].value

        # 应发工资
        worksheet["n6"] = float_JingYingGuanLi_YingFa
        worksheet["n7"] = float_ZHuanYeJiSHu_YingFa
        worksheet["n8"] = float_ShengChanCaoZuo_YingFa
        worksheet["n18"] = float_QiTaRenYuan_YingFa
        worksheet["n5"] = worksheet["n6"].value + worksheet["n7"].value + worksheet["n8"].value + worksheet["n18"].value

        # 各类保险

        worksheet["p5"] = float_endowment_insurance
        worksheet["p6"] = float_unemployement_insurance
        worksheet["p7"] = float_medical_insurance
        worksheet["p10"] = float_hourse_insurance
        worksheet["p8"] = float_big_ill_insurance
        worksheet["p9"] = float_hourse_insurance_extra
        worksheet["p11"] = float_endowment_insurance_extra
        worksheet["p12"] = float_unemployement_insurance_extra
        worksheet["p13"] = float_medical_insurance_extra
        worksheet["p14"] = float_tex
        worksheet["p18"] = float_endowment_insurance + float_unemployement_insurance + float_medical_insurance \
                           + float_hourse_insurance + float_big_ill_insurance + float_hourse_insurance_extra \
                           + float_endowment_insurance_extra + float_unemployement_insurance_extra \
                           + float_medical_insurance_extra + float_tex
        # 实发工资
        worksheet["n19"] = worksheet["n5"].value - worksheet["p18"].value

        worksheet["c19"]=strCHangeMoneyFormat(worksheet["n5"].value - worksheet["p18"].value)


        # 关闭文件并保存
        workbook.close()
        workbook.save(str_file_name)

        # 做进一步核实
        if worksheet["n19"].value != float_in_hand_money:
            print(worksheet["n19"].value, float_in_hand_money)
            raise DataFormatError("工资明细表中实发工资{0}与工资汇总表中实发工资金额{1}不一致,推测为工资明细表中有人实发工资小于0，请检查后重新生成汇总表。".format(float_in_hand_money,worksheet["n19"].value))

    def start(self):

        str_file_name = self.strReturnFileName()

        if str_file_name:
            try:
                shutil.copyfile(r"..\conf\汇总表", str_file_name)
                self.excelWageSummary(str_file_name)
            except DataFormatError as e:
                QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", str(e), QMessageBox.Ok)
                # os.unlink(str_file_name)
            except Exception as e:
                QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "发生未知错误，内容如下{}".format(str(e)), QMessageBox.Ok)

            else:
                QMessageBox.information(self.window, "河钢乐亭薪酬管理系统", "工资汇总表生成完毕！", QMessageBox.Ok)
