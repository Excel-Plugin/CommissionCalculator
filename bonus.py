# -*- coding: utf-8 -*-
# 业务员提成明细表

import InterfaceModule as IM
import os

class bonus:
    def __init__(self):
        pass
    def calc_bonus(self):
        excel=IM.Easyexcel(os.getcwd() + r"\about\2018年04道普业务提成明细.xlsx", "57578970", "57578971")
        header_dict0, sheet_data0 = excel.get_sheet("应收款4月份（数据源表）")
        pass