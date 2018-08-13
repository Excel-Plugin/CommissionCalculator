# 测试类 用于测试
import InterfaceModule as IM
import os
import CalcRatio
import  bonus
excel=IM.Easyexcel(os.getcwd() +r"\about\2018年06道普业务提成明细.xlsx")
src_dict, src_data = excel.get_sheet("数据源表")

rul_dict, rul_data = excel.get_sheet("规则")

clt_dict, clt_data = excel.get_sheet("客户编号")
client_dict = {}  # 映射关系：客户编号->该客户对应行
for row in clt_data:
    client_dict[row[clt_dict['客户编号']]] = row

sht2_head, sht2 = excel.get_sheet("指导价")
price = []
for row in sht2:
            price.append([row[sht2_head['编号']],row[sht2_head['指导单价(未税)\n元/KG']],row[sht2_head['备注']],row[sht2_head['出货开始时间']],row[sht2_head['出货结束时间']]])
sht4_head, sht4= excel.get_sheet("主管表")
slr_dict, slr_data = excel.get_sheet("售后员")
excel.close()  # 关闭输入文件

places = []  # 售后员表中的地点名
for row in slr_data:
    if row[1] != 'None':
        places.append([row[1], row[5], row[6]])
bs=bonus.Bonus(price)
h1, r1, r2 = bs.calc_commission(src_dict, src_data, clt_dict, client_dict, rul_dict, rul_data, places, sht4)
print(r1)




