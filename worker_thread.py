from PyQt5.QtCore import QThread
from InterfaceModule import Easyexcel
from after_sales import AfterSales
from CalcRatio import CalcRatio
import os
import logging
import bonus

class WorkerThread(QThread):

    def __init__(self, signal):
        super(WorkerThread, self).__init__()
        self.__signal = signal
        self.__files = []

    def setFiles(self, files):
        self.__files = files

    # 线程是否已就绪
    def isReady(self):
        return len(self.__files) > 0  # files不能为空

    def run(self):
        try:
            return self.__work()
        except Exception as err:
            print(err)
            logging.exception(err)

    # 调用这个函数来更新UI上的进度至progress（progress取值范围应为0-100）
    def __updateProgress(self, progress):
        self.__signal.emit(progress)

    # TODO:逻辑还没完成
    def __work(self):
        excel = Easyexcel(self.__files[0], "57578970", "57578971")
        src_dict, src_data = excel.get_sheet("数据源表")
        rul_dict, rul_data = excel.get_sheet("规则")
        calc_ratio = CalcRatio(rul_dict, rul_data)

        # 注意1：这里默认客户编号表里面所有行都没有空属性且文件结尾前没有空行
        # 注意2：这里默认客户编号表里所有客户类型都在规则表的"规则名"列中
        clt_dict, clt_data = excel.get_sheet("客户编号")
        client_dict = {}  # 映射关系：客户编号->该客户对应行
        for row in clt_data:
            client_dict[row[clt_dict['客户编号']]] = row
        self.__updateProgress(30)

        sht2_head, sht2 = excel.get_sheet("指导价5月（新）")
        price_dict = {}
        for row in sht2:
            price_dict[row[sht2_head['编号']]] = row[sht2_head['指导单价(未税)\n元/KG']]

        slr_dict, slr_data = excel.get_sheet("售后员")

        place = []
        for i in slr_data:
            if i[1] != 'None':
                place.append(i[1])


        after_sales = AfterSales(slr_dict, slr_data)
        as_header, as_content = after_sales.calc_commission(src_dict, src_data, clt_dict, client_dict, calc_ratio)

        bs=bonus.Bonus()
        h1, r1, r2 = bs.calc_commission(src_dict, src_data, clt_dict, client_dict, rul_dict, rul_data, price_dict, place)

        print("计算完成")
        self.__updateProgress(90)

        targetfile = "test.xlsx"
        if os.path.isfile(targetfile):
            print("覆盖原文件")
            os.remove(targetfile)
        ex = Easyexcel(os.getcwd() + "\\" + targetfile)
        ex.create_sheet("test")
        ex.set_sheet("test", as_header, as_content)
        ex.create_sheet("test1")
        ex.set_sheet("test", h1, r1)
        ex.create_sheet("test2")
        ex.set_sheet("test", h1, r2)
        print("写入完成")
        self.__updateProgress(100)


if __name__ == '__main__':
    w = WorkerThread(0)
    header, content = w.run()
    print(header)
    print(content)
