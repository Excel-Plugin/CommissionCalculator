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
            for j in range(3):
                temp.append([])
            result.append(temp)
        for i in range (len(sheet_data0)):
            result[i][header_dict0['业务']]=sheet_data0[i][header_dict0['业务']]
        print (result)

        pass


if __name__  == '__main__':
        test=Bonus()
        test.calc_bonus()