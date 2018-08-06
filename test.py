# 测试类 用于测试
import InterfaceModule as IM
import os
import CalcRatio
test=IM.Easyexcel(os.getcwd() +r"\about\2018年06道普业务提成明细.xlsx", "57578970", "57578971")
sht1_head,sht1=test.get_sheet("规则")
print(sht1_head)
print(sht1)
test2=CalcRatio.CalcRatio(sht1_head,sht1)
a,b=test2.calc(None,"大客户1%",121,"切削液")
print(a)
print(b)