# -*- coding: utf-8 -*-
# 业务员提成明细表

import InterfaceModule as IM
import os


class Bonus:
    def __init__(self):
        pass

    def calc_bonus(self):
        excel=IM.Easyexcel(os.getcwd() + r"\about\2018年04道普业务提成明细.xlsx", "57578970", "57578971")
        header_dict0, sheet_data0 = excel.get_sheet("应收款4月份（数据源表）")
        print(header_dict0)
        print(sheet_data0)
        print(len(sheet_data0))
        result=[]
        for i in range(len(sheet_data0)):
            temp=[]
            for j in range(12):
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



        print (result)

        pass


if __name__  == '__main__':
        test=Bonus()
        test.calc_bonus()