# -*- coding: utf-8 -*-
import win32com.client
import os
import pickle


def cache(get_sheet):
    """Easyexcel类中get_sheet函数的装饰器，用于自动缓存被使用过的sheet
    注意1：若被处理的Excel表不再是’about/2018年04道普业务提成明细.xlsx’，这里需要修改
    注意2：若sheet被修改了，将cached_sheets下对应的pickle文件删除即可，之后调用get_sheet时会自动生成"""
    def inner(self, sheet_name):
        if os.path.exists("cached_sheets/" + sheet_name + ".pickle"):
            print("exists")
            with open("cached_sheets/" + sheet_name + ".pickle", "rb") as f:
                return pickle.load(f)
        else:
            print("gen")
            header_dict, sheet_data = get_sheet(self, sheet_name)
            with open("cached_sheets/" + sheet_name + ".pickle", "wb") as f:
                pickle.dump((header_dict, sheet_data), f)
            return header_dict, sheet_data

    return inner


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

    def get_a_row(self, sheet_name, r, col_num=-1):
        """col_num<0,根据末尾连续空格数决定此行是否终止(用于读取表头);col_num>=0,读入长度为col_num的一行(用于读取普通数据)
        注意：在指定col_num情况下，该行只要有一个值不是None就会被返回整行，只有全为None时才会返回[]"""
        row = []
        if col_num < 0:
            c = 1  # col，本行要读入的列号，从1开始计数
            while str(self.xlBook.Worksheets(sheet_name).Cells(r, c)) != 'None' \
                    or str(self.xlBook.Worksheets(sheet_name).Cells(r, c + 1)) != 'None' \
                    or str(self.xlBook.Worksheets(sheet_name).Cells(r, c + 2)) != 'None':
                row.append(str(self.xlBook.Worksheets(sheet_name).Cells(r, c)))
                c += 1
        else:
            for c in range(1, col_num + 1):
                row.append(str(self.xlBook.Worksheets(sheet_name).Cells(r, c)))
            if row.count('None') >= col_num:  # 若该行全为None，则返回空行；反之只要有一个非None的值就正常返回row
                row = []
        return row

    @cache
    def get_sheet(self, sheet_name):
        """读取Excel表中的一个sheet，返回表头各属性对应索引dict和数据表
        注意：这里默认所有sheet都是矩阵，即所有行长度都等于表头长度"""

        # 读取表头，属性-索引字典保存在header_dict中
        r = 1
        while len(self.get_a_row(sheet_name, r)) <= 0:
            r += 1
        header = self.get_a_row(sheet_name, r)  # 表头
        r += 1
        header_dict = {}
        for i, name in enumerate(header):
            header_dict[name] = i

        # 读取表中数据到data中
        sheet_data = []
        len_ = len(header_dict)  # 表头长度
        row = self.get_a_row(sheet_name, r, len_)
        while row:
            sheet_data.append(row)
            r += 1
            row = self.get_a_row(sheet_name, r, len_)

        return header_dict, sheet_data

    def close(self):
        self.xlBook.Close(self.filename)
        del self.xlApp

    def save(self):
        self.xlBook.Save()

    def set_sheet(self, sheet_name, content):
        sht = self.xlBook.Worksheets(sheet_name)
        for i in range(len(content)):
            for j in range(len(content[0])):
                sht.Cells(i + 1, j + 1).Value = content[i][j]
        self.save()

    def create_sheet(self, sheet_name):
        sht = self.xlBook.Worksheets
        sht.Add(After='Sheet1').Name = sheet_name
        pass


if __name__ == '__main__':
    excel = Easyexcel(os.getcwd() + r"\about\2018年04道普业务提成明细.xlsx", "57578970", "57578971")
    header_dict, sheet_data = excel.get_sheet("应收款4月份（数据源表）")
    print(len(header_dict))
    print(len(sheet_data))
    print(header_dict)
    print("!!!!!")
    print(sheet_data)
