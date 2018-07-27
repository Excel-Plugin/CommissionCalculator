# coding:utf-8
from win32com.client import Dispatch
import win32com.client


class Easyexcel:
    def __init__(self, filename=None, access_password=None, write_res_password=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.xlApp.Visible = True
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(Filename=filename, UpdateLinks=2, ReadOnly=False, Format=None,
                                                    Password=access_password, WriteResPassword=write_res_password)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def get_sheet(self, sheet_name):
        A = []
        for i in range(1, 10000):
            check0=0
            for j in range (1,10):

                if str(self.xlBook.Worksheets(sheet_name).Cells(i, j)) == 'None' and i>10 :
                    check0=check0+1

            if check0>=9:
                break
            check0=0
            # print(self.xlBook.Worksheets(sheet_name).Rows(i))
            B = []
            for j in range(1, 100):
                if (str(self.xlBook.Worksheets(sheet_name).Cells(i, j)) == 'None') and (
                        str(self.xlBook.Worksheets(sheet_name).Cells(i, j + 1)) == 'None') and (
                        str(self.xlBook.Worksheets(sheet_name).Cells(i, j + 2)) == 'None'):
                            break
                B.append(str(self.xlBook.Worksheets(sheet_name).Cells(i, j)))
                print(str(self.xlBook.Worksheets(sheet_name).Cells(i, j)))
            A.append(B)
        # print(A)
        return A

    def close(self):
        self.xlBook.Close(self.filename)
        del self.xlApp

    def save(self):
        self.xlBook.Save()

    def setSheet(self, sheet_name, content):
        sht = self.xlBook.Worksheets(sheet_name)
        for i in range(len(content)):
            for j in range(len(content[0])):
                sht.Cells(i + 1, j + 1).Value = content[i][j]
        self.save()

    def createSheet(self, sheet_name):
        sht = self.xlBook.Worksheets
        sht.Add(After='Sheet1').Name = sheet_name
        pass
    def dataTableName(self):
        dtn={"业务":0,"开票日期":1,"客户编号":2,"客户名称":3,"金额":4,"发票号码":5,"到期时间":6}
        pass


# 测试一下
test = Easyexcel(r"C:\Project\RMB\昆山项目1\2018年04道普业务提成明细.xlsx","57578970","57578971")
K = test.get_sheet('应收款4月份（数据源表）')
print(K)
