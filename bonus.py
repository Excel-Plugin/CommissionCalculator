# -*- coding: utf-8 -*-
# 业务员提成明细表

import InterfaceModule as IM
import os
from datetime import datetime

class Bonus:
    def __init__(self):
        self.header = ["业务", "开票日期", "客户编号", "客户名称",
                       "开票金额（含税）", "发票号码", "到期时间", "款期", "付款日",
                       "付款金额（含税）", "付款未税金额", "到款天数", "未税服务费", "客户类型",
                       "提成比例", "提成金额", "我司单价", "公司指导价合计", "实际差价",
                       "成品代码", "品名", "规格", "数量", "单位",
                       "单价", "含税金额", "重量", "单桶公斤数量", "指导价",
                       "单号", "出货时间", "出货地点"]
        pass

    def calc_bonus(self):
        ratio=1.17 #税率为1.17
        excel=IM.Easyexcel(os.getcwd() + r"\about\2018年04道普业务提成明细.xlsx", "57578970", "57578971")
        header_dict0, sheet_data0 = excel.get_sheet("应收款4月份（数据源表）")
        print(header_dict0)
        print(sheet_data0)
        print(len(sheet_data0))
        result=[]
        for i in range(len(sheet_data0)):
            temp=[]
            for j in range(15):
                temp.append([])
            result.append(temp)
        for i in range (len(sheet_data0)):
            result[i][header_dict0['业务']] = sheet_data0[i][header_dict0['业务']]
            result[i][header_dict0['开票日期']] = sheet_data0[i][header_dict0['开票日期']]
            result[i][header_dict0['客户编号']] = sheet_data0[i][header_dict0['客户编号']]
            result[i][header_dict0['客户名称']] = sheet_data0[i][header_dict0['客户名称']]
            result[i][header_dict0['金额']] = sheet_data0[i][header_dict0['金额']]
            result[i][header_dict0['发票号码']] = sheet_data0[i][header_dict0['发票号码']]
            result[i][header_dict0['到期时间']] = sheet_data0[i][header_dict0['到期时间']]
            result[i][header_dict0['款期']] = sheet_data0[i][header_dict0['款期']]
            result[i][header_dict0['付款日']] = sheet_data0[i][header_dict0['付款日']]
            result[i][header_dict0['付款金额']] = sheet_data0[i][header_dict0['付款金额']]
            if result[i][header_dict0['付款金额']]=="未税":
                result[i][10]=result[i][header_dict0['付款金额']]
            else:
                result[i][10]=round(float(result[i][header_dict0['付款金额']])/ratio,2)

            result[i][11]=0
            result[i][12]=""
            result[i][13]="正常计算"
            result[i][14]=0.01




        print (result)

        pass


if __name__  == '__main__':
        test=Bonus()
        test.calc_bonus()