# -*- coding: utf-8 -*-
from datetime import datetime


class Saler(object):
    """用于记录售后员的相关信息，相关数据都以其对应的数据类型存储，不应直接输入到Excel中"""

    def __init__(self, name):
        self.name = name
        self.places = {}  # 出货地点的集合
        self.clients = {}  # 客户的集合

    def add_a_row(self, slr_dict, slr_row):
        """注意：单条记录中出货地点与客户编号应有且只有一个存在值"""
        if (slr_row[slr_dict['出货地点']] != 'None' and slr_row[slr_dict['客户编号']] != 'None') \
                or (slr_row[slr_dict['出货地点']] != 'None' and slr_row[slr_dict['客户编号']] != 'None'):
            raise Exception("售后员表中单行'出货地点'与'客户编号'只能有且仅有一个存在值！")
        row = slr_row
        # 值格式为'2018-04-23 00:00:00+00:00'，所以要split(' ')[0]
        row[slr_dict['开始时间']] = \
            datetime.strptime(slr_row[slr_dict['开始时间']].split(' ')[0], '%Y-%m-%d')
        row[slr_dict['结束时间']] = \
            datetime.strptime(slr_row[slr_dict['结束时间']].split(' ')[0], '%Y-%m-%d')
        if slr_row[slr_dict['出货地点']] != 'None':
            self.places[slr_row[slr_dict['出货地点']]] = row
        if slr_row[slr_dict['客户编号']] != 'None':
            row[slr_dict['提成比例']] = float(row[slr_dict['提成比例']])
            self.clients[slr_row[slr_dict['客户编号']]] = row


class AfterSales(object):
    def __init__(self, slr_dict, slr_data):
        self.salers = {}
        self.slr_dict = slr_dict
        for row in slr_data:
            if row[slr_dict['售后员']] not in self.salers:
                saler = Saler(row[slr_dict['售后员']])
                saler.add_a_row(slr_dict, row)
                self.salers[row[slr_dict['售后员']]] = saler
            else:
                self.salers[row[slr_dict['售后员']]].add_a_row(slr_dict, row)
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
        print("AfterSales gened")

    def calc_commission(self, src_dict, src_data, clt_dict, client_dict):
        """根据数据源表计算各售后服务员提成"""
        result = []  # 结果表数据
        for rcd_num, rcd in enumerate(src_data):
            print(str(rcd_num) + '/' + str(len(src_data)-1))
            place = rcd[src_dict['出货地点']]  # 该行记录对应的出货地点
            shipment = datetime.strptime(rcd[src_dict['出货时间']].split(' ')[0], "%Y-%m-%d")
            saler = None  # 该行记录对应的售后
            # 注意：这里默认一个出货地点和一个客户编号只会对应一个售后员
            for slr in self.salers.values():  # 匹配出货地点
                for plc in slr.places:  # 该售货员的地名出现在数据源表出货地点中
                    if plc in place \
                            and slr.places[plc][self.slr_dict['开始时间']] <= shipment <= slr.places[plc][
                                self.slr_dict['结束时间']]:
                        saler = slr
            if saler is None:  # 出货地点匹配失败，匹配客户编号
                number = rcd[src_dict['客户编号']]  # 该行记录对应的客户编号
                for slr in self.salers.values():
                    if number in slr.clients \
                            and slr.clients[number][self.slr_dict['开始时间']] <= shipment <= slr.clients[number][
                                self.slr_dict['结束时间']]:
                        saler = slr
                        break
            if saler is None:  # 没有对应的售后
                continue
            row = ["" for _ in range(0, len(self.rst_dict))]  # 注意这里不能用[]*len(self.rst_dict)（复制的是引用）
            row[self.rst_dict['售后']] = saler.name
            row[self.rst_dict['业务']] = rcd[src_dict['业务']]
            row[self.rst_dict['开票日期']] = rcd[src_dict['开票日期']]
            row[self.rst_dict['客户编号']] = rcd[src_dict['客户编号']]
            row[self.rst_dict['客户名称']] = rcd[src_dict['客户名称']]
            row[self.rst_dict['开票金额（含税）']] = float(rcd[src_dict['金额']])
            row[self.rst_dict['发票号码']] = rcd[src_dict['发票号码']]
            row[self.rst_dict['到期时间']] = rcd[src_dict['到期时间']]
            row[self.rst_dict['款期']] = rcd[src_dict['款期']]
            row[self.rst_dict['付款日']] = rcd[src_dict['付款日']]
            row[self.rst_dict['付款金额（含税）']] = float(rcd[src_dict['付款金额']])
            # 注意此处可能因为编码不同导致相等关系不成立
            if rcd[src_dict['发票号码']] == "未税":
                row[self.rst_dict['付款未税金额']] = float(rcd[src_dict['付款金额']])
                continue
            else:
                row[self.rst_dict['付款未税金额']] = float(rcd[src_dict['付款金额']]) / (1+float(rcd[src_dict['税率']]))
            # 值格式为'2018-04-23 00:00:00+00:00'，所以要split(' ')[0]
            # 这里的付款日格式可能形如'2018-3-31/2018-4-4'，计算时只使用最后的日期，所以要split('/')[-1]
            row[self.rst_dict['到款天数']] = \
                (datetime.strptime(rcd[src_dict['付款日']].split(' ')[0].split('/')[-1], "%Y-%m-%d")
                 - datetime.strptime(rcd[src_dict['开票日期']].split(' ')[0], "%Y-%m-%d")).days
            row[self.rst_dict['未税服务费']] = ""  # 不需要计算
            row[self.rst_dict['提成比例']] = 0  # TODO: 添加提成比例
            # 注意：这里使用的是“提成计算方式”而不是“客户类型”
            row[self.rst_dict['客户类型']] = client_dict[rcd[src_dict['客户编号']]][clt_dict['客户类型']]
            row[self.rst_dict['提成金额']] = float(rcd[src_dict['数量（桶）']]) * row[self.rst_dict['提成比例']]
            row[self.rst_dict['我司单价']] = ""  # 不需要计算
            row[self.rst_dict['公司指导价合计']] = ""  # 不需要计算
            row[self.rst_dict['实际差价']] = ""  # 不需要计算
            row[self.rst_dict['成品代码']] = rcd[src_dict['成品代码']]
            row[self.rst_dict['品名']] = rcd[src_dict['品名']]
            row[self.rst_dict['规格']] = rcd[src_dict['规格']]
            row[self.rst_dict['数量']] = float(rcd[src_dict['数量（桶）']])
            row[self.rst_dict['单位']] = rcd[src_dict['单位']]
            row[self.rst_dict['单价']] = rcd[src_dict['单价']]
            row[self.rst_dict['含税金额']] = rcd[src_dict['含税金额']]
            row[self.rst_dict['重量']] = rcd[src_dict['重量（公斤）']]
            row[self.rst_dict['单桶公斤数量']] = rcd[src_dict['单桶重量']]
            row[self.rst_dict['指导价']] = "指导价"  # 不需要计算
            row[self.rst_dict['单号']] = rcd[src_dict['单号']]
            row[self.rst_dict['出货时间']] = rcd[src_dict['出货时间']]
            row[self.rst_dict['出货地点']] = rcd[src_dict['出货地点']]
            result.append(row)
        print("result gened")

        # 计算售后员汇总
        result.sort(key=lambda row: row[self.rst_dict['售后']])  # 按照售后员人名进行排序
        print("排序完成")
        tmp = ["" for _ in range(0, len(self.rst_dict))]  # 当前售后员累计
        tmp[self.rst_dict['售后']] = result[0][self.rst_dict['售后']]
        tmp[self.rst_dict['开票金额（含税）']] = 0
        tmp[self.rst_dict['付款金额（含税）']] = 0
        tmp[self.rst_dict['付款未税金额']] = 0
        tmp[self.rst_dict['提成金额']] = 0
        tmp[self.rst_dict['数量']] = 0
        i = 0
        while i < len(result):
            row = result[i]
            print(str(i)+'/'+str(len(result)))
            if tmp[self.rst_dict['售后']] != row[self.rst_dict['售后']]:
                tmp[self.rst_dict['售后']] += " 汇总"
                result.insert(i, tmp)
                i += 1
                tmp = ["" for _ in range(0, len(self.rst_dict))]  # 重置tmp，不能重用原先的（insert进去的是tmp的引用）
                tmp[self.rst_dict['售后']] = row[self.rst_dict['售后']]
                tmp[self.rst_dict['开票金额（含税）']] = 0
                tmp[self.rst_dict['付款金额（含税）']] = 0
                tmp[self.rst_dict['付款未税金额']] = 0
                tmp[self.rst_dict['提成金额']] = 0
                tmp[self.rst_dict['数量']] = 0
            tmp[self.rst_dict['开票金额（含税）']] += row[self.rst_dict['开票金额（含税）']]
            tmp[self.rst_dict['付款金额（含税）']] += row[self.rst_dict['付款金额（含税）']]
            tmp[self.rst_dict['付款未税金额']] += row[self.rst_dict['付款未税金额']]
            tmp[self.rst_dict['提成金额']] += row[self.rst_dict['提成金额']]
            tmp[self.rst_dict['数量']] += row[self.rst_dict['数量']]
            i += 1
        tmp[self.rst_dict['售后']] += " 汇总"
        result.append(tmp)
        print("汇总完成")
        return self.header, result
