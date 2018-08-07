# 测试类 用于测试
import InterfaceModule as IM
import os
import CalcRatio
import  bonus
test=IM.Easyexcel(os.getcwd() +r"\about\2018年06道普业务提成明细.xlsx", "57578970", "57578971")
sht1_head,sht1=test.get_sheet("规则")

test2=CalcRatio.CalcRatio(sht1_head,sht1)
a,b=test2.calc(None,"1%",121,"切削液")


sht2_head,sht2=test.get_sheet("指导价5月（新）")


sht3_head,sht3=test.get_sheet("数据源表")
sht4_head,sht4=test.get_sheet("客户编号")
print(sht2)



test3=bonus.Bonus()
h1,r1=test3.calc_commission(sht3_head,sht3,sht4_head,sht4)
