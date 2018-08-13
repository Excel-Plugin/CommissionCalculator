# 测试类 用于测试
import InterfaceModule as IM
import os
import CalcRatio
import  bonus
test=IM.Easyexcel(os.getcwd() +r"\about\2018年06道普业务提成明细.xlsx")
sht1_head,sht1=test.get_sheet("主管表")
print(sht1)




