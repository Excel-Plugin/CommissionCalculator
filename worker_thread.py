from PyQt5.QtCore import QThread
from PyQt5.QtWidgets import QMessageBox
from InterfaceModule import Easyexcel
from after_sales import AfterSales
from CalcRatio import CalcRatio
import os
import logging
import bonus

class WorkerThread(QThread):

    def __init__(self, signal, progressText):
        super(WorkerThread, self).__init__()
        self.__signal = signal
        self.__progressText = progressText
        self.__files = []
        self.__readword = None
        self.__writeword = None

    def setFiles(self, files):
        self.__files = files

    def setPassWord(self, read, write):
        self.__readword, self.__writeword = read, write

    # 线程是否已就绪
    def isReady(self):
        return len(self.__files) > 0  # files不能为空

    # TODO: 子线程中无法弹出窗口
    def run(self):
        try:
            return self.__work()
        except Exception as err:
            print(err)
            self.__progressText.setText(str(err))
            # QMessageBox.warning(None, "出现错误", str(err))
            logging.exception(err)

    # 调用这个函数来更新UI上的进度至progress（progress取值范围应为0-100）
    def __updateProgress(self, progress):
        self.__signal.emit(progress)

    def __work(self):
        targetfile = "业务提成明细表.xlsx"
        if os.path.isfile(targetfile):
            os.remove(targetfile)
            # reply = QMessageBox.question(None, "是否覆盖原文件", "目标文件'业务提成明细表.xlsx'已存在，是否重新生成？")
            # if reply == QMessageBox.Yes:
            #     os.remove(targetfile)
            # else:
            #     return

        self.__progressText.setText("正在读取输入文件")
        self.__updateProgress(3)
        excel = Easyexcel(self.__files[0], False, self.__readword, self.__writeword)
        self.__updateProgress(5)
        self.__progressText.setText("正在读取数据源表")
        src_dict, src_data = excel.get_sheet("数据源表")
        self.__updateProgress(25)
        self.__progressText.setText("正在读取规则表")
        rul_dict, rul_data = excel.get_sheet("规则")
        calc_ratio = CalcRatio(rul_dict, rul_data)
        self.__updateProgress(30)

        # 注意1：这里默认客户编号表里面所有行都没有空属性且文件结尾前没有空行
        # 注意2：这里默认客户编号表里所有客户类型都在规则表的"规则名"列中
        self.__progressText.setText("正在读取客户编号表")
        clt_dict, clt_data = excel.get_sheet("客户编号")
        client_dict = {}  # 映射关系：客户编号->该客户对应行
        for row in clt_data:
            client_dict[row[clt_dict['客户编号']]] = row
        self.__updateProgress(40)

        self.__progressText.setText("正在读取指导价表")
        sht2_head, sht2 = excel.get_sheet("指导价")
        price = []
        for row in sht2:
            price.append([row[sht2_head['编号']],row[sht2_head['指导单价(未税)\n元/KG']],row[sht2_head['备注']],row[sht2_head['出货开始时间']],row[sht2_head['出货结束时间']]])
        self.__updateProgress(45)

        sht4_head, sht4= excel.get_sheet("主管表")

        self.__progressText.setText("正在读取售后员表")
        slr_dict, slr_data = excel.get_sheet("售后员")
        excel.close()  # 关闭输入文件
        self.__updateProgress(50)

        self.__progressText.setText("正在计算：业务员提成明细（售后）")
        places = []  # 售后员表中的地点名
        for row in slr_data:
            if row[1] != 'None':
                places.append([row[1],row[5],row[6]])

        after_sales = AfterSales(slr_dict, slr_data)
        as_header, as_content = after_sales.calc_commission(src_dict, src_data, clt_dict, client_dict, calc_ratio)
        self.__updateProgress(55)

        self.__progressText.setText("正在计算：业务员提成")
        bs=bonus.Bonus(price)
        h1, r1, r2 = bs.calc_commission(src_dict, src_data, clt_dict, client_dict, rul_dict, rul_data, places, sht4)
        self.__updateProgress(60)

        self.__progressText.setText("正在写入："+targetfile)
        ex = Easyexcel(os.getcwd() + "\\" + targetfile, False)
        self.__progressText.setText("正在写入：业务员提成明细（售后）")
        ex.create_sheet("业务员提成明细（售后）")
        ex.set_sheet("业务员提成明细（售后）", as_header, as_content)
        self.__updateProgress(70)
        self.__progressText.setText("正在写入：业务员提成明细")
        ex.create_sheet("业务员提成明细")
        ex.set_sheet("业务员提成明细", h1, r1)
        self.__updateProgress(85)
        self.__progressText.setText("正在写入：业务员提成打印")
        ex.create_sheet("业务员提成打印")
        ex.set_sheet("业务员提成打印", h1, r2)
        self.__updateProgress(100)
        print("写入完成")

        ex.save()
        ex.close()
        self.__files.clear()  # 清除所有文件


if __name__ == '__main__':
    w = WorkerThread(0)
    header, content = w.run()
    print(header)
    print(content)
