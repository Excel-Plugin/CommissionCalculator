from PyQt5.QtCore import QThread
from InterfaceModule import Easyexcel
from after_sales import AfterSales
import os


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
        return self.__work()

    # 调用这个函数来更新UI上的进度至progress（progress取值范围应为0-100）
    def __updateProgress(self, progress):
        self.__signal.emit(progress)

    # TODO:逻辑还没完成
    def __work(self):
        excel = Easyexcel(os.getcwd() + r"\about\2018年04道普业务提成明细.xlsx", "57578970", "57578971")
        src_dict, src_data = excel.get_sheet("应收款4月份（数据源表）")

        # 注意1：这里默认客户编号表里面所有行都没有空属性且文件结尾前没有空行
        # 注意2：这里默认客户编号表里所有客户类型都在规则表的"规则名"列中
        clt_dict, clt_data = excel.get_sheet("客户编号")
        client_dict = {}  # 映射关系：客户编号->该客户对应行
        for row in clt_data:
            client_dict[row[clt_dict['客户编号']]] = row

        after_sales = AfterSales()
        as_header, as_content = after_sales.calc_commission(src_dict, src_data, clt_dict, client_dict)
        return as_header, as_content



if __name__ == '__main__':
    w = WorkerThread(0)
    header, content = w.run()
    print(header)
    print(content)
