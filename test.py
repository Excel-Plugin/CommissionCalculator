import xlrd as xlrd

filename,password = "abc.xlsx", '123'
xlwb = xlrd.open_workbook()
print(xlwb.Sheets(1).Cells(1,1))