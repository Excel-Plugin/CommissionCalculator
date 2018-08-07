# 测试类 用于测试
import InterfaceModule as IM
import os
import CalcRatio
import  bonus
test=IM.Easyexcel(os.getcwd() +r"\about\2018年06道普业务提成明细.xlsx", "57578970", "57578971")
sht1_head,sht1=test.get_sheet("规则")
print(sht1)
test2=CalcRatio.CalcRatio(sht1_head,sht1)
a,b=test2.calc('2018-01-01 00:00:00+00:00',"正常计算",190,"切切")
print(a)
print(b)

sht2_head,sht2=test.get_sheet("指导价5月（新）")


sht3_head,sht3=test.get_sheet("数据源表")

clt_dict, clt_data = test.get_sheet("客户编号")
client_dict = {}  # 映射关系：客户编号->该客户对应行
for row in clt_data:
            client_dict[row[clt_dict['客户编号']]] = row
price_dict={}
for row in sht2:
            price_dict[row[sht2_head['编号']]]=row[sht2_head['指导单价(未税)\n元/KG']]

sht4_head,sht4=test.get_sheet("售后员")
place=[]
for i in sht4:
    if i[1]!='None':
        place.append(i[1])



test3=bonus.Bonus()
h1,r1,r2=test3.calc_commission(sht3_head,sht3,clt_dict,client_dict,sht1_head,sht1,price_dict,place)
print(len(r1))
for i in r1:
    print(i)


