import InterfaceModule
from datetime import datetime

class CalcRatio:
    def __init__(self):
        self.header=["开始时间","结束时间","规则名","业务员","固定比例","0-60"
            , "61-120","121-150","151-180","切削液","切削油"
            , "其他","售后占比"]
        im=InterfaceModule.Easyexcel(os.getcwd() + r"\about\2018年04道普业务提成明细.xlsx", "57578970", "57578971")
        self.ruleTitle, self.rules=im.get_sheet("规则")
        self.rst_dict = {}
        for i, attr in enumerate(self.header):
            self.rst_dict[attr] = i
        pass

    def calc(self,time,ruleName,days,goodsName,after_sales_name=None,salesName=None):
        sales_ratio1=0

        after_sales_ratio=0
        #开票时间、规则名、销售员名字、付款天数、货物名、售后人员名
        for rule in self.rules:
            if(rule[self.rst_dict['规则名']]==ruleName):
                if (rule[self.rst_dict['开始时间']] is not "None") and rule[self.rst_dict['开始时间']] > time:
                    break
                if (rule[self.rst_dict['结束时间']] is not "None") and rule[self.rst_dict['结束时间']] < time:
                    break
                if  days>180:
                    break
                if (rule[self.rst_dict['0-60']] is not 'None'):
                    if 0<=days and days<=60:
                        sales_ratio1=rule[self.rst_dict['0-60']]
                        break
                if (rule[self.rst_dict['61-120']] is not 'None'):
                    if 61<=days and days<=120:
                        sales_ratio1=rule[self.rst_dict['61-120']]
                        break
                if (rule[self.rst_dict['121-150']] is not 'None'):
                    if 121<=days and days<=150:
                        sales_ratio1=rule[self.rst_dict['121-150']]
                        break
                if (rule[self.rst_dict['151-180']] is not 'None'):
                    if 151<=days and days<=180:
                        sales_ratio1=rule[self.rst_dict['151-180']]
                        break
                if (rule[self.rst_dict['固定比例']] is not 'None'):

                        sales_ratio1=rule[self.rst_dict['固定比例']]
                        break
                if "切削液" in goodsName:
                    sales_ratio1=rule[self.rst_dict['切削液']]
                    after_sales_ratio=rule[self.rst_dict['售后占比']]
                    break
                if "切削油" in goodsName:
                    sales_ratio1=rule[self.rst_dict['切削油']]
                    after_sales_ratio=rule[self.rst_dict['售后占比']]
                    break
                sales_ratio1 = rule[self.rst_dict['其他']]
                after_sales_ratio = rule[self.rst_dict['售后占比']]

                break
            else:
                continue



        return  sales_ratio1,after_sales_ratio
        #返回值>1 表示每桶多少钱，返回值<1 表示比例
        pass