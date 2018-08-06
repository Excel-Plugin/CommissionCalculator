# -*- coding: utf-8 -*-
from datetime import datetime

# 来自售后员提成明细表
# 注意：此处的地名一定要与数据源表中的地名完全一致
# TODO: 售后员表有两种类型，目前只考虑了第一种
default_psn2plc = {"戴梦菲": ["龙华", "观澜","纳诺-观澜"],
                   "李飞": ["济源","纳诺-济源"],
                   "卢伟": ["郑州港区", "郑州加工区", "鹤壁", "锜昌", "建泰"],
                   "周文斌": ["廊坊", "太原", "烟台"]}


class Bonus(object):

    def __init__(self, psn2plc=default_psn2plc):
        self.plc2psn = {}
        for psn, plcs in psn2plc.items():
            for plc in plcs:
                self.plc2psn[plc] = psn
        # 表头各属性名称，按顺序放置
        self.header = [ "业务", "开票日期", "客户编号", "客户名称",
                       "开票金额（含税）", "发票号码", "到期时间", "款期", "付款日",
                       "付款金额（含税）", "付款未税金额", "到款天数", "未税服务费", "客户类型",
                       "提成比例", "提成金额", "我司单价", "公司指导价合计", "实际差价",
                       "成品代码", "品名", "规格", "数量", "单位",
                       "单价", "含税金额", "重量", "单桶公斤数量", "指导价",
                       "单号", "出货时间", "出货地点"]
        self.rst_dict = {}
        for i, attr in enumerate(self.header):
            self.rst_dict[attr] = i

    def calc_commission(self, src_dict, src_data, clt_dict, client_dict):
        """根据数据源表计算各售后服务员提成"""
        # TODO: 写入Excel的时候记得把所有float型数据按照保留两位小数显示
        # TODO: 添加汇总行
        result = []  # 结果表数据
        for rcd in src_data:

            row = ["" for _ in range(0, len(self.rst_dict))]  # 注意这里不能用[]*len(self.rst_dict)（复制的是引用）

            row[self.rst_dict['业务']] = rcd[src_dict['业务']]
            row[self.rst_dict['开票日期']] = rcd[src_dict['开票日期']]
            row[self.rst_dict['客户编号']] = rcd[src_dict['客户编号']]
            row[self.rst_dict['客户名称']] = rcd[src_dict['客户名称']]
            row[self.rst_dict['开票金额（含税）']] = rcd[src_dict['金额']]
            row[self.rst_dict['发票号码']] = rcd[src_dict['发票号码']]
            row[self.rst_dict['到期时间']] = rcd[src_dict['到期时间']]
            row[self.rst_dict['款期']] = rcd[src_dict['款期']]
            row[self.rst_dict['付款日']] = rcd[src_dict['付款日']]
            row[self.rst_dict['付款金额（含税）']] = rcd[src_dict['付款金额']]
            # 注意此处可能因为编码不同导致相等关系不成立
            if rcd[src_dict['发票号码']] == "未税":
                row[self.rst_dict['付款未税金额']] = float(rcd[src_dict['付款金额']])
                continue
            else:
                row[self.rst_dict['付款未税金额']] = float(rcd[src_dict['付款金额']]) / float(rcd[src_dict['税率']])
            # 值格式为'2018-04-23 00:00:00+00:00'，所以要split(' ')[0]
            # 这里的付款日格式可能形如'2018-3-31/2018-4-4'，计算时只使用最后的日期，所以要split('/')[-1]
            row[self.rst_dict['到款天数']] = \
                (datetime.strptime(rcd[src_dict['付款日']].split(' ')[0].split('/')[-1], "%Y-%m-%d")
                 - datetime.strptime(rcd[src_dict['开票日期']].split(' ')[0], "%Y-%m-%d")).days
            row[self.rst_dict['未税服务费']] = ""  # 不需要计算
            row[self.rst_dict['提成比例']] = 0  # TODO: 添加提成比例
            row[self.rst_dict['客户类型']] = client_dict[rcd[src_dict['客户编号']]][clt_dict['提成计算方式']]
            row[self.rst_dict['提成金额']] = float(rcd[src_dict['数量（桶）']])*row[self.rst_dict['提成比例']]
            row[self.rst_dict['我司单价']] = ""  # 不需要计算
            row[self.rst_dict['公司指导价合计']] = ""  # 不需要计算
            row[self.rst_dict['实际差价']] = ""  # 不需要计算
            row[self.rst_dict['成品代码']] = rcd[src_dict['成品代码']]
            row[self.rst_dict['品名']] = rcd[src_dict['品名']]
            row[self.rst_dict['规格']] = rcd[src_dict['规格']]
            row[self.rst_dict['数量']] = rcd[src_dict['数量（桶）']]
            row[self.rst_dict['单位']] = rcd[src_dict['单位']]
            row[self.rst_dict['单价']] = rcd[src_dict['单价']]
            row[self.rst_dict['含税金额']] = rcd[src_dict['含税金额']]
            row[self.rst_dict['数量']] = rcd[src_dict['数量（桶）']]
            row[self.rst_dict['重量']] = rcd[src_dict['重量（公斤）']]
            row[self.rst_dict['单桶公斤数量']] = rcd[src_dict['单桶重量']]
            row[self.rst_dict['指导价']] = "指导价"  # 不需要计算
            row[self.rst_dict['单号']] = rcd[src_dict['单号']]
            row[self.rst_dict['出货时间']] = rcd[src_dict['出货时间']]
            row[self.rst_dict['出货地点']] = rcd[src_dict['出货地点']]
            result.append(row)
        return self.header, result  # TODO: 由于不知道接口是否支持直接写入int,float，所以暂且没有将非str类型进行转换
