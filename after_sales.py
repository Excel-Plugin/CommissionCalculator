# -*- coding: utf-8 -*-
from datetime import datetime

# 来自售后员提成明细表
# 注意：此处的地名一定要与数据源表中的地名完全一致
psn2plc = {"戴梦菲": ["龙华", "观澜"],
           "李飞": ["济源"],
           "卢伟": ["郑州港区", "郑州加工区", "鹤壁", "锜昌", "建泰"],
           "周文斌": ["廊坊", "太原", "烟台"]}


class AfterSales(object):

    def __init__(self, psn2plc):
        self.plc2psn = {}
        for psn, plcs in psn2plc:
            for plc in plcs:
                self.plc2psn[plc] = psn
        # 表头各属性名称，按顺序放置
        self.header = ["售后", "业务", "开票日期", "客户编号", "客户名称",
                       "开票金额（含税）", "发票号码", "到期时间", "款期", "付款日",
                       "付款金额（含税）", "付款未税金额", "到款天数", "未税服务费", "客户类型",
                       "提成比例", "提成金额", "我司单价", "公司指导价合计", "实际差价",
                       "成品代码", "品名", "规格", "数量", "单位",
                       "单价", "含税金额", "重量", "单桶公斤数量", "指导价",
                       "单号", "出货时间", "出货地点"]
        self.rst_dict = {}
        for i, attr in enumerate(self.header):
            self.rst_dict[attr] = i

    def calc_commission(self, src_dict, src_data):
        """根据数据源表计算各售后服务员提成"""
        # TODO: 写入Excel的时候记得把所有float型数据按照保留两位小数显示
        result = []  # 结果表数据
        for rcd in src_data:
            place = None  # 该行记录对应的出货地点
            for plc in rcd[src_dict['出货地点']].split('-'):
                if plc in self.plc2psn:
                    place = plc
                    break
            if place is None:
                continue
            row = ["" for _ in range(0, len(self.rst_dict))]  # 注意这里不能用[]*len(self.rst_dict)（复制的是引用）
            row[self.rst_dict['售后']] = self.plc2psn[rcd[src_dict['出货地点']]]
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
            else:
                row[self.rst_dict['付款未税金额']] = float(rcd[src_dict['付款金额']]) / 1.17
            # 值格式为'2018-04-23 00:00:00+00:00'，所以要split(' ')[0]
            # 注意：这里的付款日格式可能形如'2018-3-31/2018-4-4'，但是这些记录的出货地点都是拆分付款，所以正常情况下不会在结果表中
            row[self.rst_dict['到款天数']] = \
                (datetime.strptime(row[self.rst_dict['到期时间']].split(' ')[0], "%Y-%m-%d")
                 - datetime.strptime(rcd[self.rst_dict['付款日']].split(' ')[0], "%Y-%m-%d")).days
